using ExcelDna.Integration.CustomUI;

namespace PExL.AddIn.TaskPane
{
    /// <summary>
    /// Owns the single PExL editor task pane instance and shows/hides it.
    /// </summary>
    public static class TaskPaneManager
    {
        private static dynamic? _ctp;
        private static EditorControl? _editor;

        private static dynamic? _docsCtp;
        private static DocsControl? _docs;

        public static EditorControl? Editor => _editor;

        public static void ShowEditor()
        {
            if (_ctp == null)
            {
                _ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(EditorControl), "PExL Editor");
                _editor = _ctp.ContentControl as EditorControl;
                try { _ctp.Width = 480; } catch { /* ignore sizing failures */ }
            }
            _ctp.Visible = true;

            // Poll the selection while the pane is open so the editor can offer
            // to recall PExL saved on a cell (and insert refs in insert mode).
            Interop.SelectionTracker.SetPolling(true);
        }

        public static void ShowDocs()
        {
            if (_docsCtp == null)
            {
                _docsCtp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(DocsControl), "PExL Docs");
                _docs = _docsCtp.ContentControl as DocsControl;
                try { _docsCtp.Width = 460; } catch { /* ignore sizing failures */ }
            }
            _docsCtp.Visible = true;
        }
    }
}
