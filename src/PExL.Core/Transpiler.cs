using System.Collections.Generic;
using PExL.Core.Emit;
using PExL.Core.Lexing;
using PExL.Core.Parsing;

namespace PExL.Core
{
    public sealed class CellFormula
    {
        /// <summary>Target cell/range (A1 notation), or null for a bare preview expression.</summary>
        public string? Target { get; set; }

        /// <summary>The emitted Excel formula, including the leading '='.</summary>
        public string Formula { get; set; } = "";
    }

    public sealed class TranspileResult
    {
        public List<CellFormula> Cells { get; } = new List<CellFormula>();
    }

    /// <summary>
    /// Top-level entry point: PExL source text in, native Excel formulas out.
    /// The transpiler has no Excel dependency, so it is fully unit-testable.
    /// </summary>
    public static class Transpiler
    {
        public static TranspileResult Transpile(string source)
        {
            var tokens = new Lexer(source).Tokenize();
            var program = new Parser(tokens).ParseProgram();
            var emitter = new FormulaEmitter();
            var result = new TranspileResult();

            foreach (var stmt in program.Statements)
            {
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
