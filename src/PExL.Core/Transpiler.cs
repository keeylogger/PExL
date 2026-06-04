using System;
using System.Collections.Generic;
using PExL.Core.Diagnostics;
using PExL.Core.Emit;
using PExL.Core.Lexing;
using PExL.Core.Parsing;
using PExL.Core.Parsing.Ast;

namespace PExL.Core
{
    public sealed class CellFormula
    {
        /// <summary>Target cell/range (A1 notation), or null for a bare preview expression.</summary>
        public string? Target { get; set; }

        /// <summary>The emitted Excel formula, including the leading '='.</summary>
        public string Formula { get; set; } = "";
    }

    /// <summary>
    /// A document-level global ("Global") declared with <c>MakeGlobal(EXPR) :: NAME</c>.
    /// Maps to a native Excel Defined Name so it persists in the workbook and is
    /// reusable from any cell — even for recipients who don't have the add-in.
    /// </summary>
    public sealed class GlobalDef
    {
        /// <summary>The name (becomes an Excel Defined Name).</summary>
        public string Name { get; set; } = "";

        /// <summary>The Excel "refers to" formula, including the leading '='.</summary>
        public string Formula { get; set; } = "";
    }

    /// <summary>
    /// A console-style command that performs an IDE action rather than producing a
    /// formula (e.g. <c>ShowGlobals()</c>). The add-in executes these; pure
    /// transpilation just records them.
    /// </summary>
    public sealed class ConsoleCommand
    {
        public string Name { get; set; } = "";
        public List<string> Args { get; } = new List<string>();
    }

    public sealed class TranspileResult
    {
        public List<CellFormula> Cells { get; } = new List<CellFormula>();
        public List<GlobalDef> Globals { get; } = new List<GlobalDef>();
        public List<ConsoleCommand> Commands { get; } = new List<ConsoleCommand>();
    }

    /// <summary>
    /// Top-level entry point: PExL source text in, native Excel formulas out.
    /// The transpiler has no Excel dependency, so it is fully unit-testable.
    /// </summary>
    public static class Transpiler
    {
        // Verbs that act as IDE/console commands instead of emitting a formula.
        private static readonly HashSet<string> CommandVerbs = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "showGlobals", "showGlobal", "globals", "listGlobals"
        };

        public static TranspileResult Transpile(string source) => Transpile(source, null);

        /// <summary>
        /// Transpile, treating <paramref name="knownGlobals"/> as already-declared
        /// document globals. This lets a snippet reference a global by name (e.g.
        /// <c>Pi * 2</c>) without re-declaring it with <c>MakeGlobal</c> — the host
        /// supplies the names that already live in the workbook.
        /// </summary>
        public static TranspileResult Transpile(string source, IEnumerable<string>? knownGlobals)
        {
            var tokens = new Lexer(source).Tokenize();
            var program = new Parser(tokens).ParseProgram();
            var emitter = new FormulaEmitter();
            var result = new TranspileResult();

            if (knownGlobals != null)
                foreach (var g in knownGlobals)
                    if (!string.IsNullOrWhiteSpace(g)) emitter.DefineGlobal(g);

            foreach (var stmt in program.Statements)
            {
                // MakeGlobal(EXPR) :: NAME  — declare a persistent workbook global.
                if (TryGetGlobalInner(stmt.Expression, out var inner))
                {
                    if (string.IsNullOrEmpty(stmt.BindName))
                        throw new PExLException(
                            "MakeGlobal(...) must be named, e.g. MakeGlobal(0.2) :: TaxRate",
                            stmt.Expression.Line, stmt.Expression.Column);

                    ValidateGlobalName(stmt.BindName!, stmt.Expression.Line, stmt.Expression.Column);
                    emitter.DefineGlobal(stmt.BindName!);
                    // A bare cell/range global is absolutized so the workbook name
                    // doesn't shift when referenced from different cells. Constants
                    // and formulas are emitted as written.
                    Expr toEmit = inner!;
                    if (inner is ReferenceExpr)
                    {
                        var fx = new VerbCall { Verb = "fixed", Line = inner.Line, Column = inner.Column };
                        fx.Positional.Add(inner);
                        toEmit = fx;
                    }
                    var gf = emitter.Emit(toEmit).Formula;
                    result.Globals.Add(new GlobalDef { Name = stmt.BindName!, Formula = "=" + gf });

                    // Optional convenience: `MakeGlobal(...) :: N -> C1` also drops `=N` into C1.
                    if (stmt.OutputTarget != null)
                        result.Cells.Add(new CellFormula { Target = stmt.OutputTarget, Formula = "=" + stmt.BindName });
                    continue;
                }

                // ShowGlobals() and friends — console commands, no formula emitted.
                if (TryGetCommand(stmt, out var command))
                {
                    result.Commands.Add(command!);
                    continue;
                }

                if (stmt.BindName != null)
                    emitter.Bind(stmt.BindName, stmt.Expression);

                if (stmt.OutputTarget != null)
                {
                    var f = emitter.Emit(stmt.Expression).Formula;
                    result.Cells.Add(new CellFormula { Target = stmt.OutputTarget, Formula = "=" + f });
                }
                else if (stmt.BindName == null)
                {
                    var f = emitter.Emit(stmt.Expression).Formula;
                    result.Cells.Add(new CellFormula { Target = null, Formula = "=" + f });
                }
            }

            return result;
        }

        // Reject names Excel will refuse as a Defined Name, with a friendly message
        // instead of an opaque COM error at apply time.
        private static void ValidateGlobalName(string name, int line, int column)
        {
            if (name.Length > 255)
                throw new PExLException("Global name is too long (Excel allows at most 255 characters).", line, column);
            if (string.Equals(name, "C", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "R", StringComparison.OrdinalIgnoreCase))
                throw new PExLException($"'{name}' is reserved by Excel and can't be used as a global name.", line, column);
            // A1 / AB12 / R1C1 styles collide with cell references.
            if (System.Text.RegularExpressions.Regex.IsMatch(name, "^[A-Za-z]{1,3}[0-9]+$") ||
                System.Text.RegularExpressions.Regex.IsMatch(name, "^[Rr][0-9]+[Cc][0-9]+$"))
                throw new PExLException($"'{name}' looks like a cell reference; choose a different global name.", line, column);
        }

        private static bool TryGetGlobalInner(Expr expr, out Expr? inner)
        {
            inner = null;
            if (expr is VerbCall vc && string.Equals(vc.Verb, "makeGlobal", StringComparison.OrdinalIgnoreCase))
            {
                if (vc.Positional.Count != 1)
                    throw new PExLException(
                        "MakeGlobal(...) takes exactly one value, e.g. MakeGlobal(A2:A100) :: SalesQ1",
                        vc.Line, vc.Column);
                inner = vc.Positional[0];
                return true;
            }
            return false;
        }

        private static bool TryGetCommand(Statement stmt, out ConsoleCommand? command)
        {
            command = null;
            if (stmt.Expression is VerbCall vc && CommandVerbs.Contains(vc.Verb))
            {
                command = new ConsoleCommand { Name = "showGlobals" };
                foreach (var a in vc.Positional)
                    if (a is StringLit s) command.Args.Add(s.Value);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Convenience for single-expression scenarios (and tests): returns the
        /// emitted formula (with '=') of the last statement that produces a value.
        /// </summary>
        public static string ToFormula(string source)
        {
            var result = Transpile(source);
            return result.Cells.Count > 0 ? result.Cells[result.Cells.Count - 1].Formula : string.Empty;
        }
    }
}
