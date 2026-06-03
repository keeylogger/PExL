namespace PExL.Core.Emit
{
    public enum SplitStrategy { First, Last, All }

    /// <summary>
    /// The result of emitting an expression. Usually just a formula fragment, but
    /// <c>split</c> produces a <see cref="SplitInfo"/> instead so that downstream
    /// extractors (fromLeft/fromRight/at/spill) can emit the leanest formula.
    /// </summary>
    public sealed class EmitResult
    {
        public string Formula { get; }
        public SplitInfo? Split { get; }

        public EmitResult(string formula)
        {
            Formula = formula;
        }

        public EmitResult(SplitInfo split)
        {
            Split = split;
            Formula = split.AsArrayFormula(); // fallback when used as a plain value
        }

        public override string ToString() => Formula;
    }

    /// <summary>Carries the source text, delimiter, and strategy of a split.</summary>
    public sealed class SplitInfo
    {
        public string Source { get; }
        public string Delimiter { get; }
        public SplitStrategy Strategy { get; }

        public SplitInfo(string source, string delimiter, SplitStrategy strategy)
        {
            Source = source;
            Delimiter = delimiter;
            Strategy = strategy;
        }

        public string AsArrayFormula() => $"TEXTSPLIT({Source},{Delimiter})";
    }
}
