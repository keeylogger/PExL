using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using PExL.Core;

namespace PExL.AddIn.Interop
{
    /// <summary>
    /// Writes transpiled formulas into the grid via COM and preserves the
    /// original PExL source with each target cell. When <paramref name="asStatic"/>
    /// is set, the formula is evaluated and replaced with its resulting values.
    /// </summary>
    public static class ExcelInjector
    {
        public static void Inject(IReadOnlyList<CellFormula> cells, string originalCode, bool asStatic = false)
            => Inject(cells, null, originalCode, asStatic);

        public static void Inject(IReadOnlyList<CellFormula> cells, IReadOnlyList<GlobalDef>? globals, string originalCode, bool asStatic = false)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    if (globals != null && globals.Count > 0)
                        GlobalStore.CreateMany(app, globals);
                    WriteCells(app, cells, originalCode, asStatic);
                }
                catch
                {
                    // Surface nothing into Excel; the editor already showed a preview.
                }
            });
        }

        /// <summary>
        /// Write the formulas using an already-acquired Excel application object.
        /// Call this only from inside an Excel macro context (e.g. within an
        /// <see cref="ExcelAsyncUtil.QueueAsMacro"/> callback).
        /// </summary>
        public static void WriteCells(dynamic app, IReadOnlyList<CellFormula> cells, string originalCode, bool asStatic = false)
        {
            var entry = new UndoEntry { Label = Summarize(originalCode) };
            string sheetKey = string.Empty;

            foreach (var cf in cells)
            {
                dynamic range = cf.Target != null ? app.Range[cf.Target] : app.ActiveCell;

                var snap = CapturePrior(range);

                // Formula2 keeps dynamic-array semantics and accepts invariant (comma) syntax.
                range.Formula2 = cf.Formula;

                if (asStatic)
                {
                    dynamic frozen = FreezeToValues(app, range);
                    if (snap != null)
                    {
                        try { if (frozen != null) snap.ClearAddress = (string)frozen.Address; } catch { /* best effort */ }
                    }
                }
                else
                {
                    CodeStore.Save(app, range, originalCode);
                }

                if (snap != null)
                {
                    entry.Cells.Add(snap);
                    if (string.IsNullOrEmpty(sheetKey)) sheetKey = KeyFromSnapshot(snap);
                }
            }

            UndoManager.Push(sheetKey, entry);
            Ribbon.RefreshUndo();
        }

        /// <summary>
        /// Restore the most recent PExL write on the active sheet only, on the
        /// macro thread (history is tracked per sheet).
        /// </summary>
        public static void UndoLast()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    string key = ActiveSheetKey(app);
                    if (UndoManager.TryPop(key, out var entry) && entry != null)
                        RestoreEntry(app, entry);
                }
                catch
                {
                    // Best effort; nothing to surface into Excel.
                }
                Ribbon.RefreshUndo();
            });
        }

        /// <summary>True when the active sheet has at least one undoable PExL change.</summary>
        public static bool ActiveSheetHasUndo()
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                return UndoManager.CountFor(ActiveSheetKey(app)) > 0;
            }
            catch
            {
                // If we cannot read the active sheet, fall back to "anything pending".
                return UndoManager.TotalCount > 0;
            }
        }

        private static string ActiveSheetKey(dynamic app)
        {
            string book = string.Empty, sheet = string.Empty;
            try { dynamic ws = app.ActiveSheet; sheet = (string)ws.Name; book = (string)ws.Parent.Name; }
            catch { /* ignore */ }
            return MakeKey(book, sheet);
        }

        private static string KeyFromSnapshot(CellSnapshot snap) => MakeKey(snap.Book, snap.Sheet);

        private static string MakeKey(string? book, string? sheet) => (book ?? string.Empty) + "\u0001" + (sheet ?? string.Empty);

        private static CellSnapshot? CapturePrior(dynamic range)
        {
            try
            {
                var snap = new CellSnapshot();
                try { dynamic ws = range.Worksheet; snap.Sheet = (string)ws.Name; snap.Book = (string)ws.Parent.Name; } catch { /* ignore */ }
                try { snap.Address = (string)range.Address; } catch { /* ignore */ }
                try { snap.PriorFormula = range.Formula2; } catch { snap.PriorFormula = null; }
                return string.IsNullOrEmpty(snap.Address) ? null : snap;
            }
            catch
            {
                return null;
            }
        }

        private static void RestoreEntry(dynamic app, UndoEntry entry)
        {
            // Restore in reverse so a multi-cell action unwinds cleanly.
            for (int i = entry.Cells.Count - 1; i >= 0; i--)
            {
                var snap = entry.Cells[i];
                try
                {
                    dynamic ws = ResolveSheet(app, snap);

                    if (!string.IsNullOrEmpty(snap.ClearAddress))
                    {
                        try { ws.Range[snap.ClearAddress].ClearContents(); } catch { /* ignore */ }
                    }

                    dynamic range = ws.Range[snap.Address];
                    if (IsBlank(snap.PriorFormula))
                        range.ClearContents();
                    else
                        range.Formula2 = snap.PriorFormula;
                }
                catch
                {
                    // Skip any cell we can no longer resolve.
                }
            }
        }

        private static dynamic ResolveSheet(dynamic app, CellSnapshot snap)
        {
            if (!string.IsNullOrEmpty(snap.Book) && !string.IsNullOrEmpty(snap.Sheet))
            {
                try { return app.Workbooks[snap.Book].Worksheets[snap.Sheet]; } catch { /* fall through */ }
            }
            if (!string.IsNullOrEmpty(snap.Sheet))
            {
                try { return app.Worksheets[snap.Sheet]; } catch { /* fall through */ }
            }
            return app.ActiveSheet;
        }

        private static bool IsBlank(object? priorFormula)
        {
            if (priorFormula == null) return true;
            if (priorFormula is string s) return s.Length == 0;
            return false;
        }

        private static string Summarize(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return "PExL change";
            foreach (var line in code.Replace("\r", "").Split('\n'))
            {
                var t = line.Trim();
                if (t.Length == 0 || t.StartsWith("//")) continue;
                return t.Length > 40 ? t.Substring(0, 40) + "..." : t;
            }
            return "PExL change";
        }

        /// <summary>
        /// Replace a just-written formula with its calculated result. Handles
        /// spilled dynamic arrays by snapshotting the whole spill range.
        /// </summary>
        private static dynamic FreezeToValues(dynamic app, dynamic range)
        {
            try
            {
                app.Calculate();
                dynamic target = range;
                try
                {
                    // If the formula spilled, capture the entire spill area.
                    dynamic se = range.SpillingToRange;
                    if (se != null) target = se;
                }
                catch { /* not a spilling cell */ }

                var values = target.Value2;
                target.ClearContents();
                target.Value2 = values;
                return target;
            }
            catch
            {
                try { range.Value2 = range.Value2; } catch { /* best effort */ }
                return range;
            }
        }
    }
}
