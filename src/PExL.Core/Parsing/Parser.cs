using System;
using System.Collections.Generic;
using PExL.Core.Diagnostics;
using PExL.Core.Lexing;
using PExL.Core.Parsing.Ast;
using PExL.Core.StdLib;

namespace PExL.Core.Parsing
{
    /// <summary>
    /// Recursive-descent parser for PExL. Produces a <see cref="ProgramNode"/>.
    /// Honors the forgiving layer: filler words are skipped, prepositional labels
    /// and bare positional arguments are both accepted, and verb synonyms resolve
    /// via <see cref="VerbRegistry"/>.
    /// </summary>
    public sealed class Parser
    {
        private static readonly HashSet<string> Filler = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "the", "value", "please", "a", "an"
        };
        private static readonly HashSet<string> Flags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "descending", "ascending"
        };

        private readonly List<Token> _tokens;
        private int _i;

        public Parser(List<Token> tokens)
        {
            _tokens = tokens;
        }

        private Token Cur => _tokens[_i];
        private Token Peek(int n = 1)
        {
            int j = _i + n;
            return j < _tokens.Count ? _tokens[j] : _tokens[_tokens.Count - 1];
        }

        private bool Is(TokenType t) => Cur.Type == t;
        private bool IsKeyword(string kw) => Cur.Type == TokenType.Identifier && Cur.Text == kw;

        private Token Eat(TokenType t, string what)
        {
            if (Cur.Type != t)
                throw new PExLException($"Expected {what} but found '{Cur.Text}'", Cur.Line, Cur.Column);
            var tok = Cur;
            _i++;
            return tok;
        }

        private void SkipNewlines()
        {
            while (Is(TokenType.NewLine)) _i++;
        }

        private void SkipFiller()
        {
            while (Cur.Type == TokenType.Identifier && Filler.Contains(Cur.Text)) _i++;
        }

        public ProgramNode ParseProgram()
        {
            var program = new ProgramNode();
            SkipNewlines();
            while (!Is(TokenType.EndOfInput))
            {
                var stmt = ParseStatement();
                if (stmt != null) program.Statements.Add(stmt);
                // statement separators
                if (Is(TokenType.NewLine) || Is(TokenType.Comma)) _i++;
                SkipNewlines();
            }
            return program;
        }

        private Statement ParseStatement()
        {
            // check ... block
            if (IsKeyword("check"))
                return ParseCheck();

            var line = Cur.Line; var col = Cur.Column;
            var expr = ParseExpression();
            var stmt = new Statement { Expression = expr, Line = line, Column = col };

            if (Is(TokenType.Bind))
            {
                _i++;
                SkipFiller();
                stmt.BindName = Eat(TokenType.Identifier, "a name after '::'").Text;
            }
            if (Is(TokenType.Arrow))
            {
                _i++;
                SkipFiller();
                if (Is(TokenType.CellRef)) stmt.OutputTarget = Eat(TokenType.CellRef, "a target cell/range after '->'").Text;
                else stmt.OutputTarget = Eat(TokenType.Identifier, "a target after '->'").Text;
            }
            return stmt;
        }

        // ---- expression precedence ----

        private Expr ParseExpression()
        {
            if (IsKeyword("if")) return ParseIf();
            return ParsePipe();
        }

        private Expr ParsePipe()
        {
            var left = ParseOr();
            while (Is(TokenType.Pipe))
            {
                _i++;
                SkipFiller();
                var call = ParseVerbCallAtom(allowBareArgs: false);
                call.Positional.Insert(0, left);
                left = call;
            }
            return left;
        }

        private Expr ParseOr()
        {
            var left = ParseAnd();
            while (IsKeyword("or"))
            {
                _i++;
                var right = ParseAnd();
                left = new Binary { Op = "or", Left = left, Right = right };
            }
            return left;
        }

        private Expr ParseAnd()
        {
            var left = ParseComparison();
            while (IsKeyword("and"))
            {
                _i++;
                var right = ParseComparison();
                left = new Binary { Op = "and", Left = left, Right = right };
            }
            return left;
        }

        private Expr ParseComparison()
        {
            var left = ParseAdditive();
            while (true)
            {
                string? op = null;
                switch (Cur.Type)
                {
                    case TokenType.Eq: op = "="; break;
                    case TokenType.NotEq: op = "<>"; break;
                    case TokenType.Gt: op = ">"; break;
                    case TokenType.Lt: op = "<"; break;
                    case TokenType.Gte: op = ">="; break;
                    case TokenType.Lte: op = "<="; break;
                }
                if (op != null)
                {
                    _i++;
                    var right = ParseAdditive();
                    left = new Binary { Op = op, Left = left, Right = right };
                    continue;
                }
                // English aliases: is / is not
                if (IsKeyword("is"))
                {
                    _i++;
                    string o = "=";
                    if (IsKeyword("not")) { _i++; o = "<>"; }
                    var right = ParseAdditive();
                    left = new Binary { Op = o, Left = left, Right = right };
                    continue;
                }
                // infix contains / startsWith / endsWith
                if (Cur.Type == TokenType.Identifier &&
                    (string.Equals(Cur.Text, "contains", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(Cur.Text, "startsWith", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(Cur.Text, "endsWith", StringComparison.OrdinalIgnoreCase)))
                {
                    var verb = VerbRegistry.Canonical(Cur.Text);
                    _i++;
                    var right = ParseAdditive();
                    var call = new VerbCall { Verb = verb };
                    call.Positional.Add(left);
                    call.Positional.Add(right);
                    left = call;
                    continue;
                }
                break;
            }
            return left;
        }

        private Expr ParseAdditive()
        {
            var left = ParseMultiplicative();
            while (Is(TokenType.Plus) || Is(TokenType.Minus))
            {
                var op = Is(TokenType.Plus) ? "+" : "-";
                _i++;
                var right = ParseMultiplicative();
                left = new Binary { Op = op, Left = left, Right = right };
            }
            return left;
        }

        private Expr ParseMultiplicative()
        {
            var left = ParsePower();
            while (Is(TokenType.Star) || Is(TokenType.Slash))
            {
                var op = Is(TokenType.Star) ? "*" : "/";
                _i++;
                var right = ParsePower();
                left = new Binary { Op = op, Left = left, Right = right };
            }
            return left;
        }

        private Expr ParsePower()
        {
            var left = ParseUnary();
            if (Is(TokenType.Caret))
            {
                _i++;
                var right = ParsePower(); // right associative
                return new Binary { Op = "^", Left = left, Right = right };
            }
            return left;
        }

        private Expr ParseUnary()
        {
            if (Is(TokenType.Minus))
            {
                _i++;
                return new Unary { Op = "-", Operand = ParseUnary() };
            }
            if (IsKeyword("not"))
            {
                _i++;
                return new Unary { Op = "not", Operand = ParseUnary() };
            }
            return ParsePrimary();
        }

        private Expr ParsePrimary()
        {
            SkipFiller();
            var tok = Cur;
            switch (tok.Type)
            {
                case TokenType.Number:
                    _i++;
                    return new NumberLit { Raw = tok.Text, Line = tok.Line, Column = tok.Column };
                case TokenType.String:
                    _i++;
                    return new StringLit { Value = tok.Text, Line = tok.Line, Column = tok.Column };
                case TokenType.DateLiteral:
                    _i++;
                    return new DateLit { Raw = tok.Text, Line = tok.Line, Column = tok.Column };
                case TokenType.CellRef:
                    _i++;
                    return new ReferenceExpr { Text = tok.Text, Line = tok.Line, Column = tok.Column };
                case TokenType.LParen:
                    _i++;
                    var inner = ParseExpression();
                    Eat(TokenType.RParen, "')'");
                    return inner;
                case TokenType.Identifier:
                    if (tok.Text == "if") return ParseIf();
                    if (tok.Text == "true") { _i++; return new BoolLit { Value = true }; }
                    if (tok.Text == "false") { _i++; return new BoolLit { Value = false }; }
                    if (tok.Text == "empty") { _i++; return new EmptyLit(); }
                    if (VerbRegistry.IsVerb(tok.Text) || Peek().Type == TokenType.LParen || Peek().Type == TokenType.Dot)
                        return ParseVerbCallAtom(allowBareArgs: true);
                    // otherwise: a reference to a previously bound name
                    _i++;
                    return new NameRef { Name = tok.Text, Line = tok.Line, Column = tok.Column };
                default:
                    throw new PExLException($"Unexpected token '{tok.Text}'", tok.Line, tok.Column);
            }
        }

        private VerbCall ParseVerbCallAtom(bool allowBareArgs)
        {
            var nameTok = Eat(TokenType.Identifier, "a verb");
            var call = new VerbCall { Verb = VerbRegistry.Canonical(nameTok.Text), Line = nameTok.Line, Column = nameTok.Column };

            // legacy.* : legacy.vlookup(...) -> capture the real function name as modifier
            // .modifier
            if (Is(TokenType.Dot))
            {
                _i++;
                call.Modifier = Eat(TokenType.Identifier, "a modifier after '.'").Text;
            }

            // (positional args)
            if (Is(TokenType.LParen))
            {
                _i++;
                if (!Is(TokenType.RParen))
                {
                    call.Positional.Add(ParseExpression());
                    while (Is(TokenType.Comma))
                    {
                        _i++;
                        call.Positional.Add(ParseExpression());
                    }
                }
                Eat(TokenType.RParen, "')'");
            }

            // trailing bare positionals, prepositional args, method-chains, and flags
            while (true)
            {
                // .label(value) method-style named arg
                if (Is(TokenType.Dot) && Peek().Type == TokenType.Identifier)
                {
                    _i++;
                    var label = Eat(TokenType.Identifier, "a method name").Text;
                    Expr val;
                    if (Is(TokenType.LParen))
                    {
                        _i++;
                        val = Is(TokenType.RParen) ? new EmptyLit() : ParseExpression();
                        Eat(TokenType.RParen, "')'");
                    }
                    else val = new BoolLit { Value = true };
                    call.Named.Add(new NamedArg { Label = VerbRegistry.CanonicalPreposition(label), Value = val });
                    continue;
                }

                if (Cur.Type == TokenType.Identifier)
                {
                    // prepositional label
                    if (VerbRegistry.IsPreposition(Cur.Text))
                    {
                        var label = VerbRegistry.CanonicalPreposition(Cur.Text);
                        _i++;
                        Expr val;
                        if (Is(TokenType.LParen))
                        {
                            _i++;
                            val = Is(TokenType.RParen) ? new EmptyLit() : ParseExpression();
                            Eat(TokenType.RParen, "')'");
                        }
                        else
                        {
                            val = ParseOr();
                        }
                        call.Named.Add(new NamedArg { Label = label, Value = val });
                        continue;
                    }
                    // flag word
                    if (Flags.Contains(Cur.Text))
                    {
                        var flag = Cur.Text.ToLowerInvariant();
                        _i++;
                        call.Named.Add(new NamedArg { Label = flag, Value = new BoolLit { Value = true } });
                        continue;
                    }
                }

                // bare positional argument (English form like: find D2 within ...)
                if (allowBareArgs && StartsAtom())
                {
                    call.Positional.Add(ParseOr());
                    continue;
                }

                break;
            }

            return call;
        }

        private bool StartsAtom()
        {
            switch (Cur.Type)
            {
                case TokenType.Number:
                case TokenType.String:
                case TokenType.DateLiteral:
                case TokenType.CellRef:
                case TokenType.LParen:
                    return true;
                case TokenType.Identifier:
                    if (Filler.Contains(Cur.Text)) return true;
                    if (VerbRegistry.IsPreposition(Cur.Text)) return false;
                    if (Cur.Text == "and" || Cur.Text == "or" || Cur.Text == "then" ||
                        Cur.Text == "else" || Cur.Text == "is") return false;
                    return true; // verb or bound name
                default:
                    return false;
            }
        }

        private Expr ParseIf()
        {
            var ifTok = Eat(TokenType.Identifier, "'if'");
            var cond = ParseOr();
            if (!IsKeyword("then"))
                throw new PExLException("Expected 'then' in if-expression", Cur.Line, Cur.Column);
            _i++;
            var thenExpr = ParseOr();
            Expr? elseExpr = null;
            if (IsKeyword("else"))
            {
                _i++;
                elseExpr = ParseOr();
            }
            return new IfExpr { Condition = cond, Then = thenExpr, Else = elseExpr, Line = ifTok.Line, Column = ifTok.Column };
        }

        // Flexible check block. All of these are accepted:
        //   check SUBJECT:            check:                  check
        //     if > 90 then "A"          if c1 then "A"          c1 then "A"
        //     else "C"                  else "C"                c2 then "B"
        //                                                       else "C"
        //   -> TARGET   (optional; :: name also allowed)
        // A subject is optional; the colon and the leading 'if' on each branch are
        // optional. A branch is any line that contains 'then'.
        private Statement ParseCheck()
        {
            var checkTok = Eat(TokenType.Identifier, "'check'");
            SkipNewlines();

            Expr? subject = null;
            Expr? pendingCond = null;

            if (Is(TokenType.Colon))
            {
                _i++; // 'check:' with no subject
            }
            else if (!IsKeyword("if") && !IsKeyword("else") &&
                     !Is(TokenType.Arrow) && !Is(TokenType.Bind) && !Is(TokenType.EndOfInput))
            {
                var e = ParseOr();
                if (Is(TokenType.Colon)) { subject = e; _i++; }
                else if (IsKeyword("then")) { pendingCond = e; } // it was the first branch, not a subject
                else { subject = e; }                            // subject without a colon
            }
            SkipNewlines();

            var ifs = new VerbCall { Verb = "ifs", Line = checkTok.Line, Column = checkTok.Column };

            while (true)
            {
                if (pendingCond == null)
                {
                    SkipNewlines();
                    if (!BranchAhead()) break;
                    if (IsKeyword("if")) _i++;
                }

                Expr cond = pendingCond ?? ParseBranchCondition(subject);
                pendingCond = null;
                if (!IsKeyword("then"))
                    throw new PExLException("Expected 'then' in a check branch", Cur.Line, Cur.Column);
                _i++;
                Expr result = ParseOr();
                ifs.Positional.Add(cond);
                ifs.Positional.Add(result);
                SkipNewlines();
            }

            if (IsKeyword("else"))
            {
                _i++;
                ifs.Named.Add(new NamedArg { Label = "else", Value = ParseOr() });
                SkipNewlines();
            }

            var stmt = new Statement { Expression = ifs, Line = checkTok.Line, Column = checkTok.Column };
            if (Is(TokenType.Bind))
            {
                _i++;
                SkipFiller();
                stmt.BindName = Eat(TokenType.Identifier, "a name after '::'").Text;
            }
            if (Is(TokenType.Arrow))
            {
                _i++;
                SkipFiller();
                if (Is(TokenType.CellRef)) stmt.OutputTarget = Eat(TokenType.CellRef, "a target cell/range after '->'").Text;
                else stmt.OutputTarget = Eat(TokenType.Identifier, "a target after '->'").Text;
            }
            return stmt;
        }

        /// <summary>True when the tokens ahead begin another check branch.</summary>
        private bool BranchAhead()
        {
            if (IsKeyword("else") || Is(TokenType.Arrow) || Is(TokenType.Bind) || Is(TokenType.EndOfInput))
                return false;
            if (IsKeyword("if")) return true;
            return LineHasThen();
        }

        /// <summary>Scan to the end of the current line for a 'then' keyword.</summary>
        private bool LineHasThen()
        {
            for (int j = _i; j < _tokens.Count; j++)
            {
                var t = _tokens[j];
                if (t.Type == TokenType.NewLine || t.Type == TokenType.EndOfInput ||
                    t.Type == TokenType.Arrow || t.Type == TokenType.Bind)
                    return false;
                if (t.Type == TokenType.Identifier && t.Text == "then")
                    return true;
            }
            return false;
        }

        private Expr ParseBranchCondition(Expr? subject)
        {
            // implicit-subject comparison: "if > 90"
            string? op = null;
            switch (Cur.Type)
            {
                case TokenType.Eq: op = "="; break;
                case TokenType.NotEq: op = "<>"; break;
                case TokenType.Gt: op = ">"; break;
                case TokenType.Lt: op = "<"; break;
                case TokenType.Gte: op = ">="; break;
                case TokenType.Lte: op = "<="; break;
            }
            if (op != null && subject != null)
            {
                _i++;
                var rhs = ParseAdditive();
                return new Binary { Op = op, Left = subject, Right = rhs };
            }
            // implicit-subject verb: "if contains "x""
            if (subject != null && Cur.Type == TokenType.Identifier &&
                (string.Equals(Cur.Text, "contains", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(Cur.Text, "startsWith", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(Cur.Text, "endsWith", StringComparison.OrdinalIgnoreCase)))
            {
                var verb = VerbRegistry.Canonical(Cur.Text);
                _i++;
                var arg = ParseAdditive();
                var call = new VerbCall { Verb = verb };
                call.Positional.Add(subject);
                call.Positional.Add(arg);
                return call;
            }
            // explicit full condition
            return ParseOr();
        }
    }
}
