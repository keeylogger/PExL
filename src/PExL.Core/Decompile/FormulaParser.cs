using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using PExL.Core.Diagnostics;

namespace PExL.Core.Decompile
{
    // ---------------- AST ----------------

    /// <summary>Base node for a parsed native Excel formula.</summary>
    public abstract class FNode { }

    public sealed class FNum : FNode { public string Raw = ""; }
    public sealed class FStr : FNode { public string Value = ""; }
    public sealed class FBool : FNode { public bool Value; }
    public sealed class FRef : FNode { public string Text = ""; }
    public sealed class FName : FNode { public string Name = ""; }   // bare name / named range
    public sealed class FCall : FNode { public string Name = ""; public List<FNode> Args = new List<FNode>(); }
    public sealed class FUnary : FNode { public string Op = ""; public FNode Operand = default!; }
    public sealed class FPostfix : FNode { public string Op = ""; public FNode Operand = default!; }
    public sealed class FBinary : FNode { public string Op = ""; public FNode Left = default!; public FNode Right = default!; }
    public sealed class FArray : FNode { public List<List<FNode>> Rows = new List<List<FNode>>(); }

    // ---------------- tokens ----------------

    internal enum FTok
    {
        Number, String, Ref, Name, Bool,
        LParen, RParen, LBrace, RBrace, Comma, Semicolon,
        Plus, Minus, Star, Slash, Caret, Amp, Percent,
        Eq, NotEq, Gt, Lt, Gte, Lte,
        End
    }

    internal struct FToken
    {
        public FTok Type;
        public string Text;
        public int Pos;
        public FToken(FTok t, string s, int p) { Type = t; Text = s; Pos = p; }
    }

    // ---------------- lexer ----------------

    internal sealed class FormulaLexer
    {
        // Same A1 grammar PExL recognizes (cells, ranges, whole columns/rows, sheet-qualified).
        private static readonly Regex RefRegex = new Regex(
            @"\G(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?" +
            @"(?:" +
                @"\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?" +
                @"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}" +
                @"|\$?[0-9]+:\$?[0-9]+" +
            @")",
            RegexOptions.Compiled);

        private readonly string _src;
        private int _pos;

        public FormulaLexer(string src) { _src = src ?? string.Empty; }

        private char Cur => _pos < _src.Length ? _src[_pos] : '\0';
        private char Peek(int n = 1) => _pos + n < _src.Length ? _src[_pos + n] : '\0';

        public List<FToken> Tokenize()
        {
            var toks = new List<FToken>();
            FToken t;
            do { t = Next(); toks.Add(t); } while (t.Type != FTok.End);
            return toks;
        }

        private FToken Next()
        {
            while (Cur == ' ' || Cur == '\t' || Cur == '\r' || Cur == '\n') _pos++;
            int start = _pos;
            if (_pos >= _src.Length) return new FToken(FTok.End, "", start);

            char c = Cur;

            // two-char comparisons
            if (c == '>' && Peek() == '=') { _pos += 2; return new FToken(FTok.Gte, ">=", start); }
            if (c == '<' && Peek() == '=') { _pos += 2; return new FToken(FTok.Lte, "<=", start); }
            if (c == '<' && Peek() == '>') { _pos += 2; return new FToken(FTok.NotEq, "<>", start); }

            if (c == '"') return ScanString(start);

            // references take priority over identifiers/numbers, but a name that is
            // immediately followed by '(' is a function call, not a cell reference.
            if (char.IsLetter(c) || c == '$' || c == '\'' || char.IsDigit(c))
            {
                var m = RefRegex.Match(_src, _pos);
                if (m.Success && m.Index == _pos)
                {
                    int after = _pos + m.Length;
                    char next = after < _src.Length ? _src[after] : '\0';
                    if (next != '(')
                    {
                        _pos = after;
                        return new FToken(FTok.Ref, m.Value, start);
                    }
                }
            }

            if (char.IsDigit(c)) return ScanNumber(start);
            if (char.IsLetter(c) || c == '_') return ScanName(start);

            _pos++;
            switch (c)
            {
                case '(': return new FToken(FTok.LParen, "(", start);
                case ')': return new FToken(FTok.RParen, ")", start);
                case '{': return new FToken(FTok.LBrace, "{", start);
                case '}': return new FToken(FTok.RBrace, "}", start);
                case ',': return new FToken(FTok.Comma, ",", start);
                case ';': return new FToken(FTok.Semicolon, ";", start);
                case '+': return new FToken(FTok.Plus, "+", start);
                case '-': return new FToken(FTok.Minus, "-", start);
                case '*': return new FToken(FTok.Star, "*", start);
                case '/': return new FToken(FTok.Slash, "/", start);
                case '^': return new FToken(FTok.Caret, "^", start);
                case '&': return new FToken(FTok.Amp, "&", start);
                case '%': return new FToken(FTok.Percent, "%", start);
                case '=': return new FToken(FTok.Eq, "=", start);
                case '>': return new FToken(FTok.Gt, ">", start);
                case '<': return new FToken(FTok.Lt, "<", start);
            }
            throw new PExLException($"Unexpected character '{c}' in formula");
        }

        private FToken ScanString(int start)
        {
            _pos++; // opening quote
            var sb = new StringBuilder();
            while (true)
            {
                if (_pos >= _src.Length) throw new PExLException("Unterminated string in formula");
                char c = Cur;
                if (c == '"')
                {
                    if (Peek() == '"') { sb.Append('"'); _pos += 2; continue; }
                    _pos++; break;
                }
                sb.Append(c); _pos++;
            }
            return new FToken(FTok.String, sb.ToString(), start);
        }

        private FToken ScanNumber(int start)
        {
            var sb = new StringBuilder();
            while (char.IsDigit(Cur)) { sb.Append(Cur); _pos++; }
            if (Cur == '.' && char.IsDigit(Peek())) { sb.Append('.'); _pos++; while (char.IsDigit(Cur)) { sb.Append(Cur); _pos++; } }
            if (Cur == 'e' || Cur == 'E')
            {
                sb.Append(Cur); _pos++;
                if (Cur == '+' || Cur == '-') { sb.Append(Cur); _pos++; }
                while (char.IsDigit(Cur)) { sb.Append(Cur); _pos++; }
            }
            return new FToken(FTok.Number, sb.ToString(), start);
        }

        private FToken ScanName(int start)
        {
            var sb = new StringBuilder();
            while (char.IsLetterOrDigit(Cur) || Cur == '_' || Cur == '.') { sb.Append(Cur); _pos++; }
            string s = sb.ToString();
            if (string.Equals(s, "TRUE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, "FALSE", StringComparison.OrdinalIgnoreCase))
                return new FToken(FTok.Bool, s, start);
            return new FToken(FTok.Name, s, start);
        }
    }

    // ---------------- parser ----------------

    /// <summary>
    /// Recursive-descent parser for native Excel formulas. Produces a small,
    /// generic <see cref="FNode"/> tree that <see cref="PExLWriter"/> turns back
    /// into readable PExL.
    /// </summary>
    internal sealed class FormulaParser
    {
        private readonly List<FToken> _t;
        private int _i;

        public FormulaParser(List<FToken> tokens) { _t = tokens; }

        private FToken Cur => _t[_i];
        private bool Is(FTok t) => Cur.Type == t;
        private FToken Eat(FTok t, string what)
        {
            if (Cur.Type != t) throw new PExLException($"Expected {what} in formula but found '{Cur.Text}'");
            return _t[_i++];
        }

        public FNode ParseAll()
        {
            var e = ParseExpr();
            if (!Is(FTok.End)) throw new PExLException($"Unexpected '{Cur.Text}' after formula");
            return e;
        }

        // = <> > < >= <=   (lowest)
        private FNode ParseExpr()
        {
            var left = ParseConcat();
            while (true)
            {
                string? op = Cur.Type switch
                {
                    FTok.Eq => "=",
                    FTok.NotEq => "<>",
                    FTok.Gt => ">",
                    FTok.Lt => "<",
                    FTok.Gte => ">=",
                    FTok.Lte => "<=",
                    _ => null
                };
                if (op == null) break;
                _i++;
                var right = ParseConcat();
                left = new FBinary { Op = op, Left = left, Right = right };
            }
            return left;
        }

        private FNode ParseConcat()
        {
            var left = ParseAdditive();
            while (Is(FTok.Amp))
            {
                _i++;
                var right = ParseAdditive();
                left = new FBinary { Op = "&", Left = left, Right = right };
            }
            return left;
        }

        private FNode ParseAdditive()
        {
            var left = ParseMultiplicative();
            while (Is(FTok.Plus) || Is(FTok.Minus))
            {
                string op = Is(FTok.Plus) ? "+" : "-";
                _i++;
                var right = ParseMultiplicative();
                left = new FBinary { Op = op, Left = left, Right = right };
            }
            return left;
        }

        private FNode ParseMultiplicative()
        {
            var left = ParsePower();
            while (Is(FTok.Star) || Is(FTok.Slash))
            {
                string op = Is(FTok.Star) ? "*" : "/";
                _i++;
                var right = ParsePower();
                left = new FBinary { Op = op, Left = left, Right = right };
            }
            return left;
        }

        private FNode ParsePower()
        {
            var left = ParseUnary();
            if (Is(FTok.Caret))
            {
                _i++;
                var right = ParsePower(); // right-associative
                return new FBinary { Op = "^", Left = left, Right = right };
            }
            return left;
        }

        private FNode ParseUnary()
        {
            if (Is(FTok.Minus)) { _i++; return new FUnary { Op = "-", Operand = ParseUnary() }; }
            if (Is(FTok.Plus)) { _i++; return ParseUnary(); }
            return ParsePostfix();
        }

        private FNode ParsePostfix()
        {
            var e = ParsePrimary();
            while (Is(FTok.Percent)) { _i++; e = new FPostfix { Op = "%", Operand = e }; }
            return e;
        }

        private FNode ParsePrimary()
        {
            var tok = Cur;
            switch (tok.Type)
            {
                case FTok.Number: _i++; return new FNum { Raw = tok.Text };
                case FTok.String: _i++; return new FStr { Value = tok.Text };
                case FTok.Bool: _i++; return new FBool { Value = string.Equals(tok.Text, "TRUE", StringComparison.OrdinalIgnoreCase) };
                case FTok.Ref: _i++; return new FRef { Text = tok.Text };
                case FTok.LParen:
                    _i++;
                    var inner = ParseExpr();
                    Eat(FTok.RParen, "')'");
                    return inner;
                case FTok.LBrace: return ParseArray();
                case FTok.Name:
                    _i++;
                    if (Is(FTok.LParen)) return ParseCall(tok.Text);
                    return new FName { Name = tok.Text };
                default:
                    throw new PExLException($"Unexpected '{tok.Text}' in formula");
            }
        }

        private FNode ParseCall(string name)
        {
            Eat(FTok.LParen, "'('");
            var call = new FCall { Name = name };
            if (!Is(FTok.RParen))
            {
                call.Args.Add(ParseExpr());
                while (Is(FTok.Comma)) { _i++; call.Args.Add(ParseExpr()); }
            }
            Eat(FTok.RParen, "')'");
            return call;
        }

        private FNode ParseArray()
        {
            Eat(FTok.LBrace, "'{'");
            var arr = new FArray();
            var row = new List<FNode>();
            arr.Rows.Add(row);
            if (!Is(FTok.RBrace))
            {
                row.Add(ParseExpr());
                while (Is(FTok.Comma) || Is(FTok.Semicolon))
                {
                    bool newRow = Is(FTok.Semicolon);
                    _i++;
                    if (newRow) { row = new List<FNode>(); arr.Rows.Add(row); }
                    row.Add(ParseExpr());
                }
            }
            Eat(FTok.RBrace, "'}'");
            return arr;
        }
    }
}
