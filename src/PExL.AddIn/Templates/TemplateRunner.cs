using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using PExL.AddIn.Interop;
using PExL.AddIn.TaskPane;
using PExL.Core;

namespace PExL.AddIn.Templates
{
    /// <summary>
    /// One-click ribbon actions. Simple transforms/generators run straight on the
    /// active cell (no editor); range, multi-range, spilling and wizard templates
    /// open the editor seeded with the user's current selection so nothing
    /// references a random A1/A2 anymore.
    /// </summary>
    public static class TemplateRunner
    {
        public static void Run(string id)
        {
            // Wizards collect input on the UI thread first.
            if (id == "tplPivot") { RunPivot(); return; }
            if (id == "tplDatePicker") { RunDatePicker(); return; }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                dynamic app = ExcelDnaUtil.Application;
                var sel = Capture(app);
                var plan = Plan(id, sel);

                if (plan.Direct)
                {
                    try
                    {
                        var result = Transpiler.Transpile(plan.Code);
                        ExcelInjector.WriteCells(app, result.Cells, plan.Code);
                        return;
                    }
                    catch { /* fall through to the editor so the user can see/fix it */ }
                }
                OpenEditor(plan.Code);
            });
        }

        /// <summary>Read the active cell formula and rewrite it as PExL in the editor.</summary>
        public static void RunTranslate()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                string code;
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic cell = app.ActiveCell;
                    string addr = cell.Address[false, false];
                    string formula = ReadFormula(cell);
                    code = Decompiler.ToPExL(formula, addr);
                }
                catch (Exception ex)
                {
                    code = "// Could not read the selected cell:\n//   " + ex.Message;
                }
                OpenEditor(code);
            });
        }

        private static string ReadFormula(dynamic cell)
        {
            try { string f = (string)cell.Formula2; if (!string.IsNullOrEmpty(f)) return f; } catch { }
            try { return (string)cell.Formula ?? string.Empty; } catch { }
            return string.Empty;
        }

        private static void RunDatePicker()
        {
            using (var dlg = new DatePickerForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                string lit = "#" + dlg.SelectedDate.ToString("yyyy-MM-dd") + "#";
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    dynamic app = ExcelDnaUtil.Application;
                    var sel = Capture(app);
                    string code = lit + " -> " + sel.Active;
                    try
                    {
                        var result = Transpiler.Transpile(code);
                        ExcelInjector.WriteCells(app, result.Cells, code);
                    }
                    catch { OpenEditor(code); }
                });
            }
        }

        private static void RunPivot()
        {
            string snippet;
            using (var dlg = new PivotBuilderForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                snippet = dlg.BuildPExL();
            }
            OpenEditor(snippet);
        }

        private static void OpenEditor(string code)
        {
            TaskPaneManager.ShowEditor();
            TaskPaneManager.Editor?.SetCode(code);
        }

        // ---------------- planning ----------------

        private struct PlanResult { public bool Direct; public string Code; }
        private static PlanResult Direct(string code) => new PlanResult { Direct = true, Code = code };
        private static PlanResult Editor(string code) => new PlanResult { Direct = false, Code = code };

        private static PlanResult Plan(string id, Sel s)
        {
            switch (id)
            {
                // ---- one-click: write straight to the selected cell ----
                case "tplTrimClean":
                    return s.Left != null
                        ? Direct(s.Left + " |> trim |> clean -> " + s.Active)
                        : Editor("// Trim spaces and strip nonprintable characters\n" + s.Active + " |> trim |> clean");

                case "tplProperCase":
                    return s.Left != null
                        ? Direct(s.Left + " |> trim |> proper -> " + s.Active)
                        : Editor("// Title-case names, after trimming\n" + s.Active + " |> trim |> proper");

                case "tplExtractEmail":
                    return s.Left != null
                        ? Direct(s.Left + " |> split.Last(\"@\") |> fromRight -> " + s.Active)
                        : Editor("// Pull the domain out of an email address\n" + s.Active + " |> split.Last(\"@\") |> fromRight");

                case "tplSafeDivide":
                    return (s.Left != null && s.Left2 != null)
                        ? Direct("ifError(" + s.Left2 + " / " + s.Left + ", \"-\") -> " + s.Active)
                        : Editor("// Divide, but show a dash instead of #DIV/0!\nifError(" + s.Active + " / " + (s.Right ?? "B2") + ", \"-\")");

                case "tplToday":
                    return Direct("today() -> " + s.Active);

                case "tplAge":
                {
                    string dateCell = s.Left ?? "C5";
                    return Editor(
                        "// Difference between " + dateCell + " and today, written to " + s.Active + "\n" +
                        "// Order-proof: min()/max() keep it working even if the date is in the future\n" +
                        "dateDiff.years(min(" + dateCell + ", today()), max(" + dateCell + ", today())) -> " + s.Active + "\n" +
                        "// Pick the unit you need:\n" +
                        "//   dateDiff.days(" + dateCell + ", today())   -> " + s.Active + "\n" +
                        "//   dateDiff.months(" + dateCell + ", today()) -> " + s.Active + "\n" +
                        "//   dateDiff.years(" + dateCell + ", today())  -> " + s.Active);
                }

                // ---- editor-seeded: reference the current selection ----
                case "tplSplitColumns":
                    return Editor(
                        "// Split " + s.Active + " on a delimiter into two columns\n" +
                        s.Active + " |> split.First(\"-\") :: parts\n" +
                        "parts |> fromLeft  -> " + (s.Right ?? "C2") + "\n" +
                        "parts |> fromRight -> " + (s.Right2 ?? "D2") + "\n" +
                        "// Variants:\n" +
                        "//   split on something else: " + s.Active + " |> split.First(\",\")\n" +
                        "//   split on the LAST match: " + s.Active + " |> split.Last(\"-\")");

                case "tplCombine":
                    return Editor(
                        "// Combine cells with a separator\n" +
                        "combine(" + s.Active + ", " + (s.Right ?? "B2") + ") with(\" \") -> " + (s.Right2 ?? "C2") + "\n" +
                        "// Variants:\n" +
                        "//   different separator: combine(" + s.Active + ", " + (s.Right ?? "B2") + ") with(\", \")\n" +
                        "//   more cells:          combine(" + s.Active + ", " + (s.Right ?? "B2") + ", " + (s.Right2 ?? "C2") + ") with(\" \")");

                case "tplCsvSplit":
                    return Editor(
                        "// Spread a comma-separated value across columns\n" +
                        s.Active + " |> split(\",\") |> spill -> " + (s.Right ?? "B2") + "\n" +
                        "// Variants:\n" +
                        "//   split on a pipe:      " + s.Active + " |> split(\"|\") |> spill -> " + (s.Right ?? "B2") + "\n" +
                        "//   split into rows:      " + s.Active + " |> split(\",\") |> spillDown -> " + (s.Right ?? "B2"));

                case "tplRemoveDuplicates":
                    return Editor(
                        "// Unique values from " + s.Src + "\n" +
                        "unique(" + s.Src + ") -> " + s.Landing + "\n" +
                        "// Variant — keep only values that appear exactly once:\n" +
                        "//   unique(" + s.Src + ") exactlyOnce -> " + s.Landing);

                case "tplSortBy":
                    return Editor(
                        "// Sort " + s.Src + " by its first column (descending)\n" +
                        "sort(" + s.Src + ") by(1) descending -> " + s.Landing + "\n" +
                        "// Variants:\n" +
                        "//   ascending:        sort(" + s.Src + ") by(1) ascending -> " + s.Landing + "\n" +
                        "//   by another column: sort(" + s.Src + ") by(2) descending -> " + s.Landing);

                case "tplFilterRows":
                    return Editor(
                        "// Keep only rows that match a condition\n" +
                        "filter(" + s.Src + ") where(" + s.Src + " = \"value\") -> " + s.Landing + "\n" +
                        "// Variants — any comparison works:\n" +
                        "//   greater than:  filter(" + s.Src + ") where(" + s.Src + " > 100) -> " + s.Landing + "\n" +
                        "//   not equal to:  filter(" + s.Src + ") where(" + s.Src + " <> \"West\") -> " + s.Landing + "\n" +
                        "//   contains text: filter(" + s.Src + ") where(" + s.Src + " contains \"abc\") -> " + s.Landing);

                case "tplTopN":
                    return Editor(
                        "// Top 10 rows of " + s.Src + " by its first column\n" +
                        "sort(" + s.Src + ") by(1) descending :: ranked\n" +
                        "take(ranked, 10) -> " + s.Landing + "\n" +
                        "// Variants:\n" +
                        "//   bottom 10:   sort(" + s.Src + ") by(1) ascending :: ranked\n" +
                        "//   top 5:       take(ranked, 5) -> " + s.Landing);

                case "tplFillBlanks":
                    return Editor(
                        "// Fill a blank cell from the value above it\n" +
                        (s.Left != null
                            ? "if " + s.Left + " is empty then " + (s.Above ?? s.Left) + " else " + s.Left + " -> " + s.Active + "\n"
                            : "if " + s.Active + " is empty then " + (s.Above ?? s.Active) + " else " + s.Active + "\n") +
                        "// Variant — fill with a fixed default instead:\n" +
                        "//   if " + (s.Left ?? s.Active) + " is empty then \"N/A\" else " + (s.Left ?? s.Active) + " -> " + s.Active);

                case "tplLookupBuilder":
                    return Editor(
                        "// Look " + s.Active + " up in another sheet\n" +
                        "find " + s.Active + " within Sheet2!A:A thenReturn Sheet2!B:B ifMissing \"N/A\" -> " + (s.Right ?? "C2") + "\n" +
                        "// Variants:\n" +
                        "//   return a number column: find " + s.Active + " within Sheet2!A:A thenReturn Sheet2!C:C ifMissing 0 -> " + (s.Right ?? "C2") + "\n" +
                        "//   blank when missing:     find " + s.Active + " within Sheet2!A:A thenReturn Sheet2!B:B ifMissing \"\" -> " + (s.Right ?? "C2"));

                case "tplCheck":
                {
                    string subj = s.Left ?? s.Active;
                    string tgt = s.Left != null ? s.Active : (s.Right ?? "C2");
                    return Editor(
                        "// Multi-branch logic - the first matching line wins\n" +
                        "check\n" +
                        "  " + subj + " >= 90 then \"A\"\n" +
                        "  " + subj + " >= 80 then \"B\"\n" +
                        "  " + subj + " >= 70 then \"C\"\n" +
                        "  else \"F\"\n" +
                        "-> " + tgt + "\n" +
                        "// Variant - name a subject once, then compare with just the operator:\n" +
                        "//   check " + subj + ":\n" +
                        "//     if >= 90 then \"A\"\n" +
                        "//     else \"F\"\n" +
                        "//   -> " + tgt);
                }

                case "tplQuickStats":
                    return Editor(QuickStats(s));

                case "tplSumWhere":
                    return Editor(
                        "// Conditional total\n" +
                        "sumWhere(" + s.Src + ", " + s.Src + " = \"West\") -> " + s.Landing + "\n" +
                        "// Variants:\n" +
                        "//   numeric condition:  sumWhere(" + s.Src + ", " + s.Src + " > 100) -> " + s.Landing + "\n" +
                        "//   count instead:      countWhere(" + s.Src + ", " + s.Src + " = \"West\") -> " + s.Landing);

                default:
                    return Editor("// Template '" + id + "' is on the roadmap.");
            }
        }

        private static string QuickStats(Sel s)
        {
            string[] verbs = { "count", "sum", "avg", "min", "max" };
            var sb = new StringBuilder();
            sb.Append("// Quick stats for ").Append(s.Src).Append('\n');
            for (int i = 0; i < verbs.Length; i++)
            {
                sb.Append(verbs[i]).Append('(').Append(s.Src).Append(')');
                string? tgt = (s.StatTargets != null && s.StatTargets.Length > i) ? s.StatTargets[i] : null;
                if (tgt != null) sb.Append(" -> ").Append(tgt);
                if (i < verbs.Length - 1) sb.Append('\n');
            }
            return sb.ToString();
        }

        // ---------------- selection capture ----------------

        private sealed class Sel
        {
            public string Active = "A1";
            public string Selection = "A1";
            public bool Multi;
            public string? Left;
            public string? Left2;
            public string? Right;
            public string? Right2;
            public string? Above;
            public string? RightOfSel;
            public string[]? StatTargets;

            public string Src => Multi ? Selection : Active;
            public string Landing => RightOfSel ?? Right ?? Active;
        }

        private static Sel Capture(dynamic app)
        {
            var s = new Sel();
            try
            {
                dynamic active = app.ActiveCell;
                s.Active = active.Address[false, false];
                s.Left = TryAddr(() => active.Offset[0, -1]);
                s.Left2 = TryAddr(() => active.Offset[0, -2]);
                s.Right = TryAddr(() => active.Offset[0, 1]);
                s.Right2 = TryAddr(() => active.Offset[0, 2]);
                s.Above = TryAddr(() => active.Offset[-1, 0]);
            }
            catch { /* no active cell */ }

            try
            {
                dynamic sel = app.Selection;
                s.Selection = sel.Address[false, false];
                s.Multi = s.Selection.IndexOf(':') >= 0 || s.Selection.IndexOf(',') >= 0;

                try
                {
                    dynamic topRight = sel.Cells[1, (int)sel.Columns.Count];
                    dynamic land = topRight.Offset[0, 1];
                    s.RightOfSel = land.Address[false, false];
                    var stats = new List<string>();
                    for (int i = 0; i < 5; i++)
                    {
                        string? a = TryAddr(() => land.Offset[i, 0]);
                        if (a == null) break;
                        stats.Add(a);
                    }
                    if (stats.Count > 0) s.StatTargets = stats.ToArray();
                }
                catch { /* leave landing/stat targets unset */ }
            }
            catch { /* selection isn't a range */ }

            return s;
        }

        private static string? TryAddr(Func<dynamic> get)
        {
            try { dynamic r = get(); return (string)r.Address[false, false]; }
            catch { return null; }
        }
    }
}
