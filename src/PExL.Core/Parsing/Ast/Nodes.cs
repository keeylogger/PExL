using System.Collections.Generic;

namespace PExL.Core.Parsing.Ast
{
    /// <summary>Base class for every PExL syntax-tree node.</summary>
    public abstract class Node
    {
        public int Line { get; set; }
        public int Column { get; set; }
    }

    /// <summary>A whole script: an ordered list of statements.</summary>
    public sealed class ProgramNode : Node
    {
        public List<Statement> Statements { get; } = new List<Statement>();
    }

    /// <summary>EXPRESSION [:: name] [-> target].</summary>
    public sealed class Statement : Node
    {
        public Expr Expression { get; set; } = default!;
        public string? BindName { get; set; }
        public string? OutputTarget { get; set; } // A1 / A1:B2 (raw reference text)
    }

    public abstract class Expr : Node { }

    public sealed class NumberLit : Expr { public string Raw { get; set; } = ""; }
    public sealed class StringLit : Expr { public string Value { get; set; } = ""; }
    public sealed class BoolLit : Expr { public bool Value { get; set; } }
    public sealed class EmptyLit : Expr { }
    public sealed class DateLit : Expr { public string Raw { get; set; } = ""; public string? Locale { get; set; } }

    /// <summary>A literal Excel reference (cell, range, sheet-qualified).</summary>
    public sealed class ReferenceExpr : Expr { public string Text { get; set; } = ""; }

    /// <summary>A reference to a value previously bound with <c>::</c>.</summary>
    public sealed class NameRef : Expr { public string Name { get; set; } = ""; }

    /// <summary>A named/prepositional argument (e.g. <c>within A1:A100</c>).</summary>
    public sealed class NamedArg
    {
        public string Label { get; set; } = "";
        public Expr Value { get; set; } = default!;
    }

    /// <summary>verb[.modifier]([positional...]) [preposition value ...].</summary>
    public sealed class VerbCall : Expr
    {
        public string Verb { get; set; } = "";
        public string? Modifier { get; set; }
        public List<Expr> Positional { get; } = new List<Expr>();
        public List<NamedArg> Named { get; } = new List<NamedArg>();

        public Expr? FindNamed(string label)
        {
            foreach (var a in Named)
                if (string.Equals(a.Label, label, System.StringComparison.OrdinalIgnoreCase))
                    return a.Value;
            return null;
        }
    }

    public sealed class Unary : Expr
    {
        public string Op { get; set; } = ""; // "-" or "not"
        public Expr Operand { get; set; } = default!;
    }

    public sealed class Binary : Expr
    {
        public string Op { get; set; } = ""; // ^ * / + - = <> > < >= <= and or
        public Expr Left { get; set; } = default!;
        public Expr Right { get; set; } = default!;
    }

    /// <summary>if COND then THEN [else ELSE].</summary>
    public sealed class IfExpr : Expr
    {
        public Expr Condition { get; set; } = default!;
        public Expr Then { get; set; } = default!;
        public Expr? Else { get; set; }
    }
}
