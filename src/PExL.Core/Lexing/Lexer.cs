using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using PExL.Core.Diagnostics;

namespace PExL.Core.Lexing
{
    /// <summary>
    /// Hand-written lexer for PExL. Whitespace is insignificant except newlines,
    /// which separate statements. Excel A1 references are recognized as a single
    /// token so the parser can treat them as atoms.
    /// </summary>
    public sealed class Lexer
    {
        private static readonly Regex ReferenceRegex = new Regex(
            @"\G(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?" +
            @"(?:" +
                @"\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?" + // cell or cell:cell
                @"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}" +                        // whole columns A:A
                @"|\$?[0-9]+:\$?[0-9]+" +                                      // whole rows 1:1
            @")",
            RegexOptions.Compiled);

        private readonly string _src;
        private int _pos;
        private int _line = 1;
        private int _col = 1;

        public Lexer(string source)
        {
            _src = source ?? string.Empty;
        }

        public List<Token> Tokenize()
        {
            var tokens = new List<Token>();
            Token t;
            do
            {
                t = Next();
                tokens.Add(t);
            } while (t.Type != TokenType.EndOfInput);
            return tokens;
        }

        private char Cur => _pos < _src.Length ? _src[_pos] : '\0';
        private char Peek(int n = 1) => _pos + n < _src.Length ? _src[_pos + n] : '\0';

        private void Advance(int n = 1)
        {
            for (int i = 0; i < n; i++)
            {
                if (_pos >= _src.Length) return;
                if (_src[_pos] == '\n') { _line++; _col = 1; }
                else { _col++; }
                _pos++;
            }
        }

        private Token Next()
        {
            SkipTrivia();

            int line = _line, col = _col, start = _pos;
            if (_pos >= _src.Length)
                return new Token(TokenType.EndOfInput, string.Empty, start, line, col);

            char c = Cur;

            if (c == '\n')
            {
                Advance();
                return new Token(TokenType.NewLine, "\\n", start, line, col);
            }

            // multi-char operators
            if (c == '|' && Peek() == '>') { Advance(2); return new Token(TokenType.Pipe, "|>", start, line, col); }
            if (c == ':' && Peek() == ':') { Advance(2); return new Token(TokenType.Bind, "::", start, line, col); }
            if (c == '-' && Peek() == '>') { Advance(2); return new Token(TokenType.Arrow, "->", start, line, col); }
            if (c == '>' && Peek() == '=') { Advance(2); return new Token(TokenType.Gte, ">=", start, line, col); }
            if (c == '<' && Peek() == '=') { Advance(2); return new Token(TokenType.Lte, "<=", start, line, col); }
            if (c == '<' && Peek() == '>') { Advance(2); return new Token(TokenType.NotEq, "<>", start, line, col); }
            if (c == '!' && Peek() == '=') { Advance(2); return new Token(TokenType.NotEq, "!=", start, line, col); }

            // symbol forms of boolean operators -> emit as identifier keywords
            if (c == '&' && Peek() == '&') { Advance(2); return new Token(TokenType.Identifier, "and", start, line, col); }
            if (c == '|' && Peek() == '|') { Advance(2); return new Token(TokenType.Identifier, "or", start, line, col); }
            if (c == '!') { Advance(); return new Token(TokenType.Identifier, "not", start, line, col); }

            // strings
            if (c == '"') return ScanString('"', line, col);
            if (c == '`') return ScanString('`', line, col);
            if (c == '#') return ScanDate(line, col);

            // references (try before identifiers/numbers)
            if (char.IsLetter(c) || c == '$' || c == '\'' || char.IsDigit(c))
            {
                var refTok = TryScanReference(line, col);
                if (refTok.HasValue) return refTok.Value;
            }

            if (char.IsDigit(c)) return ScanNumber(line, col);
            if (char.IsLetter(c) || c == '_') return ScanIdentifier(line, col);

            // single-char tokens
            Advance();
            switch (c)
            {
                case '.': return new Token(TokenType.Dot, ".", start, line, col);
                case ',': return new Token(TokenType.Comma, ",", start, line, col);
                case ':': return new Token(TokenType.Colon, ":", start, line, col);
                case '(': return new Token(TokenType.LParen, "(", start, line, col);
                case ')': return new Token(TokenType.RParen, ")", start, line, col);
                case '+': return new Token(TokenType.Plus, "+", start, line, col);
                case '-': return new Token(TokenType.Minus, "-", start, line, col);
                case '*': return new Token(TokenType.Star, "*", start, line, col);
                case '/': return new Token(TokenType.Slash, "/", start, line, col);
                case '^': return new Token(TokenType.Caret, "^", start, line, col);
                case '=': return new Token(TokenType.Eq, "=", start, line, col);
                case '>': return new Token(TokenType.Gt, ">", start, line, col);
                case '<': return new Token(TokenType.Lt, "<", start, line, col);
            }

            throw new PExLException($"Unexpected character '{c}'", line, col);
        }

        private void SkipTrivia()
        {
            while (_pos < _src.Length)
            {
                char c = Cur;
                if (c == ' ' || c == '\t' || c == '\r') { Advance(); continue; }
                if (c == '/' && Peek() == '/')
                {
                    while (_pos < _src.Length && Cur != '\n') Advance();
                    continue;
                }
                break;
            }
        }

        private Token? TryScanReference(int line, int col)
        {
            var m = ReferenceRegex.Match(_src, _pos);
            if (!m.Success || m.Index != _pos) return null;
            // Guard: a plain integer (no colon, no letters) is a number, not a ref.
            int start = _pos;
            Advance(m.Length);
            return new Token(TokenType.CellRef, m.Value, start, line, col);
        }

        private Token ScanNumber(int line, int col)
        {
            int start = _pos;
            var sb = new StringBuilder();
            while (char.IsDigit(Cur)) { sb.Append(Cur); Advance(); }
            if (Cur == '.' && char.IsDigit(Peek()))
            {
                sb.Append('.'); Advance();
                while (char.IsDigit(Cur)) { sb.Append(Cur); Advance(); }
            }
            if (Cur == 'e' || Cur == 'E')
            {
                sb.Append(Cur); Advance();
                if (Cur == '+' || Cur == '-') { sb.Append(Cur); Advance(); }
                while (char.IsDigit(Cur)) { sb.Append(Cur); Advance(); }
            }
            if (Cur == '%')
            {
                sb.Append('%'); Advance();
            }
            return new Token(TokenType.Number, sb.ToString(), start, line, col);
        }

        private Token ScanIdentifier(int line, int col)
        {
            int start = _pos;
            var sb = new StringBuilder();
            while (char.IsLetterOrDigit(Cur) || Cur == '_') { sb.Append(Cur); Advance(); }
            return new Token(TokenType.Identifier, sb.ToString(), start, line, col);
        }

        private Token ScanString(char quote, int line, int col)
        {
            int start = _pos;
            Advance(); // opening quote
            var sb = new StringBuilder();
            while (true)
            {
                if (_pos >= _src.Length)
                    throw new PExLException("Unterminated string literal", line, col);
                char c = Cur;
                if (c == quote)
                {
                    // doubled quote inside a normal string => literal quote
                    if (quote == '"' && Peek() == '"') { sb.Append('"'); Advance(2); continue; }
                    Advance();
                    break;
                }
                sb.Append(c);
                Advance();
            }
            return new Token(TokenType.String, sb.ToString(), start, line, col);
        }

        private Token ScanDate(int line, int col)
        {
            int start = _pos;
            Advance(); // opening #
            var sb = new StringBuilder();
            while (_pos < _src.Length && Cur != '#')
            {
                sb.Append(Cur);
                Advance();
            }
            if (_pos >= _src.Length)
                throw new PExLException("Unterminated date literal (missing closing '#')", line, col);
            Advance(); // closing #
            return new Token(TokenType.DateLiteral, sb.ToString(), start, line, col);
        }
    }
}
