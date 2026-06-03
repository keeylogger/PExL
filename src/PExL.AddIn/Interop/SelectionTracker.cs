using System;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace PExL.AddIn.Interop
{
    /// <summary>
    /// Watches the current Excel selection while the editor pane is open and
    /// raises <see cref="SelectionChanged"/> on every change. Uses a lightweight
    /// UI-thread poll (rather than a COM event sink) so clicking a cell or
    /// dragging a range reaches the editor in near real time.
    ///
    /// Two independent concerns:
    ///   * Polling   - on while the editor pane is visible (drives recall).
    ///   * InsertMode - whether a change should be inserted into the editor as a
    ///                  reference. When off, the editor uses the change to offer
    ///                  loading any PExL saved for that cell.
    /// </summary>
    public static class SelectionTracker
    {
        private static Timer? _timer;
        private static Timer? _sheetTimer;
        private static bool _polling;
        private static string _last = string.Empty;
        private static string _lastSheet = string.Empty;

        public static Action<string>? SelectionChanged;

        public static bool InsertMode { get; private set; }

        public static void Install()
        {
            _timer = new Timer { Interval = 150 };
            _timer.Tick += Tick;

            // Always-on, low-frequency watch for the active sheet so the Undo
            // button (which tracks history per sheet) re-enables/greys correctly
            // when the user switches tabs.
            _sheetTimer = new Timer { Interval = 800 };
            _sheetTimer.Tick += SheetTick;
            _sheetTimer.Start();
        }

        public static void Uninstall()
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Tick -= Tick;
                _timer.Dispose();
                _timer = null;
            }
            if (_sheetTimer != null)
            {
                _sheetTimer.Stop();
                _sheetTimer.Tick -= SheetTick;
                _sheetTimer.Dispose();
                _sheetTimer = null;
            }
        }

        private static void SheetTick(object? sender, EventArgs e)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic ws = app.ActiveSheet;
                    string key = (string)ws.Parent.Name + "\u0001" + (string)ws.Name;
                    if (key != _lastSheet)
                    {
                        _lastSheet = key;
                        Ribbon.RefreshUndo();
                    }
                }
                catch { /* no active sheet / COM busy - ignore */ }
            });
        }

        /// <summary>Start/stop polling - called when the editor pane is shown/hidden.</summary>
        public static void SetPolling(bool on)
        {
            _polling = on;
            if (_timer == null) return;
            if (on) _timer.Start(); else _timer.Stop();
        }

        /// <summary>Toggle whether selection changes are inserted as references.</summary>
        public static void SetInsertMode(bool on) => InsertMode = on;

        private static void Tick(object? sender, EventArgs e)
        {
            if (!_polling) return;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic selection = app.Selection;
                    string address = selection.Address[false, false];
                    if (!string.IsNullOrEmpty(address) && address != _last)
                    {
                        _last = address;
                        SelectionChanged?.Invoke(address);
                    }
                }
                catch { /* selection not a range, or COM busy - ignore */ }
            });
        }
    }
}
