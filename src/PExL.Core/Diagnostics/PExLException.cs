using System;

namespace PExL.Core.Diagnostics
{
    /// <summary>
    /// Raised when PExL source cannot be lexed, parsed, or emitted. Carries a
    /// human-readable message and (when known) the source position so the editor
    /// can surface a squiggle.
    /// </summary>
    public sealed class PExLException : Exception
    {
        public int Line { get; }
        public int Column { get; }

        public PExLException(string message, int line = 0, int column = 0)
            : base(message)
        {
            Line = line;
            Column = column;
        }
    }
}
