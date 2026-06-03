using PExL.Core.Decompile;
using PExL.Core.Diagnostics;

namespace PExL.Core
{
    /// <summary>
    /// Reverse of <see cref="Transpiler"/>: takes a native Excel formula and
    /// rewrites it as readable PExL. Known functions map to PExL verbs; anything
    /// unrecognized degrades gracefully to <c>raw(...)</c> / <c>legacy.*</c> so the
    /// output still recompiles to an equivalent formula.
    /// </summary>
    public static class Decompiler
    {
        /// <summary>
        /// Translate a single Excel formula to PExL.
        /// </summary>
        /// <param name="formula">The formula text (with or without a leading '=').</param>
        /// <param name="target">Optional cell to append as <c>-&gt; target</c>.</param>
        public static string ToPExL(string formula, string? target = null)
        {
            string body = Clean(formula);
            if (body.Length == 0)
                return "// (the selected cell has no formula)";

            string pexl;
            try
            {
                var tokens = new FormulaLexer(body).Tokenize();
                var tree = new FormulaParser(tokens).ParseAll();
                pexl = new PExLWriter().Write(tree);
            }
            catch (PExLException ex)
            {
                return "// Could not fully translate this formula:\n//   " + ex.Message +
                       "\nraw(\"" + body.Replace("\"", "'") + "\")";
            }

            if (!string.IsNullOrEmpty(target))
                pexl += (pexl.IndexOf('\n') >= 0 ? "\n-> " : " -> ") + target;
            return pexl;
        }

        private static string Clean(string formula)
        {
            if (formula == null) return string.Empty;
            string s = formula.Trim();
            if (s.StartsWith("=")) s = s.Substring(1).Trim();
            // Excel sometimes shows a leading implicit-intersection '@'.
            if (s.StartsWith("@")) s = s.Substring(1).Trim();
            return s;
        }
    }
}
