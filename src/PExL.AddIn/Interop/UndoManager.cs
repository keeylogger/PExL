using System.Collections.Generic;

namespace PExL.AddIn.Interop
{
    /// <summary>
    /// The prior state of one written cell (or range), captured just before a
    /// PExL write so it can be restored later.
    /// </summary>
    public sealed class CellSnapshot
    {
        public string? Book;           // owning workbook name (for resolution)
        public string? Sheet;          // owning worksheet name
        public string? Address;        // anchor cell/range to restore
        public object? PriorFormula;   // Formula2 before the write (string, 2D array, or null/empty when blank)
        public string? ClearAddress;   // optional area to clear first (e.g. a frozen spill range)
    }

    /// <summary>One undoable PExL action: a label plus the cells it touched.</summary>
    public sealed class UndoEntry
    {
        public string Label = "PExL change";
        public List<CellSnapshot> Cells = new List<CellSnapshot>();
    }

    /// <summary>
    /// A capped history of PExL writes, kept per worksheet. Excel discards its
    /// native undo stack on programmatic COM writes, so we keep our own; keying
    /// by sheet means Undo only ever touches the sheet the user is looking at.
    /// </summary>
    public static class UndoManager
    {
        private const int MaxDepth = 100;
        private static readonly Dictionary<string, List<UndoEntry>> _stacks =
            new Dictionary<string, List<UndoEntry>>();

        /// <summary>Undo entries available for one sheet.</summary>
        public static int CountFor(string sheetKey)
        {
            if (string.IsNullOrEmpty(sheetKey)) return 0;
            lock (_stacks)
            {
                return _stacks.TryGetValue(sheetKey, out var list) ? list.Count : 0;
            }
        }

        /// <summary>Total entries across every sheet (used as a safe fallback).</summary>
        public static int TotalCount
        {
            get
            {
                lock (_stacks)
                {
                    int n = 0;
                    foreach (var kv in _stacks) n += kv.Value.Count;
                    return n;
                }
            }
        }

        public static void Push(string sheetKey, UndoEntry entry)
        {
            if (entry == null || entry.Cells.Count == 0 || string.IsNullOrEmpty(sheetKey)) return;
            lock (_stacks)
            {
                if (!_stacks.TryGetValue(sheetKey, out var list))
                {
                    list = new List<UndoEntry>();
                    _stacks[sheetKey] = list;
                }
                list.Add(entry);
                while (list.Count > MaxDepth) list.RemoveAt(0);
            }
        }

        public static bool TryPop(string sheetKey, out UndoEntry? entry)
        {
            entry = null;
            if (string.IsNullOrEmpty(sheetKey)) return false;
            lock (_stacks)
            {
                if (!_stacks.TryGetValue(sheetKey, out var list) || list.Count == 0) return false;
                entry = list[list.Count - 1];
                list.RemoveAt(list.Count - 1);
                if (list.Count == 0) _stacks.Remove(sheetKey);
                return true;
            }
        }

        public static void Clear()
        {
            lock (_stacks) { _stacks.Clear(); }
        }
    }
}
