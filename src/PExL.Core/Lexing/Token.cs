namespace PExL.Core.Lexing
{
    public enum TokenType
    {
        // literals / atoms
        Number,
        String,        // "..." or `...`
        DateLiteral,   // #2024-01-01#
        Identifier,    // verbs, keywords, bind-names
        CellRef,       // A1, $A$1, Sheet1!A1, A1:B2, 'My Sheet'!A1

        // operators
        Pipe,          // |>
        Bind,          // ::
        Arrow,         // ->
        Dot,           // .
        Comma,         // ,
        Colon,         // :
        LParen,        // (
        RParen,        // )

        Plus, Minus, Star, Slash, Caret,
        Eq, NotEq, Gt, Lt, Gte, Lte,

        // structure
        NewLine,
        EndOfInput
    }

    public readonly struct Token
    {
        public TokenType Type { get; }
        public string Text { get; }
        public int Position { get; }
        public int Line { get; }
        public int Column { get; }

        public Token(TokenType type, string text, int position, int line, int column)
        {
            Type = type;
            Text = text;
            Position = position;
            Line = line;
            Column = column;
        }

        public override string ToString() => $"{Type}('{Text}')@{Line}:{Column}";
    }
}
