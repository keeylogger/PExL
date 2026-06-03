using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PExL.AddIn.Interop;
using PExL.Core;

namespace PExL.AddIn.TaskPane
{
    /// <summary>
    /// The WebView2-hosted Monaco editor. Must be COM-visible so Excel-DNA can
    /// instantiate it inside a Custom Task Pane.
    /// </summary>
    [ComVisible(true)]
    public sealed class EditorControl : UserControl
    {
        private readonly WebView2 _web;
        private bool _ready;
        private string _lastCode = string.Empty;
        private string? _pendingCode;

        /// <summary>Replace the editor contents (used by templates to pre-fill PExL).</summary>
        public void SetCode(string code)
        {
            if (!_ready) { _pendingCode = code; return; }
            Post(new { type = "setCode", code });
        }

        public EditorControl()
        {
            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);
            SelectionTracker.SelectionChanged += OnExcelSelectionChanged;
            _ = InitializeAsync();
        }

        private async System.Threading.Tasks.Task InitializeAsync()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? AppDomain.CurrentDomain.BaseDirectory;

                // Point WebView2 at the architecture-specific loader if bundled (Excel-DNA gotcha).
                string nativeDir = Path.Combine(baseDir, "runtimes", "win-x64", "native");
                if (Directory.Exists(nativeDir))
                {
                    try { CoreWebView2Environment.SetLoaderDllFolderPath(nativeDir); } catch { /* best effort */ }
                }

                string userData = Path.Combine(Path.GetTempPath(), "PExL.WebView2");
                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                await _web.EnsureCoreWebView2Async(env);

                _web.CoreWebView2.WebMessageReceived += OnWebMessage;

                string? indexPath = ResolveIndexHtml(baseDir);
                if (indexPath != null)
                    _web.CoreWebView2.Navigate(new Uri(indexPath).AbsoluteUri);
                else
                    _web.CoreWebView2.NavigateToString(FallbackHtml(
                        "Editor assets (web/index.html) were not found next to the add-in."));

                _ready = true;
                if (_pendingCode != null) { Post(new { type = "setCode", code = _pendingCode }); _pendingCode = null; }
            }
            catch (Exception ex)
            {
                try { _web.CoreWebView2?.NavigateToString(FallbackHtml(ex.Message)); } catch { /* ignore */ }
            }
        }

        /// <summary>
        /// Locate web/index.html. Works whether the add-in is loaded unpacked
        /// (web/ beside the .xll) or packed from a publish/ subfolder (web/ one
        /// level up), and falls back to the source tree during development.
        /// </summary>
        private static string? ResolveIndexHtml(string baseDir)
        {
            string[] candidates =
            {
                Path.Combine(baseDir, "web", "index.html"),
                Path.Combine(baseDir, "..", "web", "index.html"),
                Path.Combine(baseDir, "..", "..", "PExL.Editor.Web", "index.html"),
                Path.Combine(baseDir, "..", "..", "..", "..", "PExL.Editor.Web", "index.html"),
            };
            foreach (var c in candidates)
            {
                try { if (File.Exists(c)) return Path.GetFullPath(c); } catch { /* ignore */ }
            }
            return null;
        }

        private void OnWebMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.TryGetWebMessageAsString();
                var root = JObject.Parse(json);
                string type = (string?)root["type"] ?? string.Empty;

                switch (type)
                {
                    case "compile":
                        HandleCompile(GetCode(root), apply: false, forApply: false, asStatic: false);
                        break;
                    case "previewForApply":
                        HandleCompile(GetCode(root), apply: false, forApply: true, asStatic: false);
                        break;
                    case "apply":
                        HandleCompile(GetCode(root), apply: true, forApply: false,
                            asStatic: (bool?)root["asStatic"] ?? false);
                        break;
                    case "insertMode":
                        SelectionTracker.SetInsertMode((bool?)root["on"] ?? false);
                        break;
                }
            }
            catch (Exception ex)
            {
                Post(new { type = "error", message = ex.Message });
            }
        }

        private static string GetCode(JObject root) =>
            (string?)root["code"] ?? string.Empty;

        private void HandleCompile(string code, bool apply, bool forApply, bool asStatic)
        {
            _lastCode = code;
            try
            {
                var result = Transpiler.Transpile(code);
                var cells = new object[result.Cells.Count];
                for (int i = 0; i < result.Cells.Count; i++)
                    cells[i] = new { target = result.Cells[i].Target, formula = result.Cells[i].Formula };

                Post(new { type = forApply ? "previewForApply" : "preview", cells });

                if (apply)
                {
                    ExcelInjector.Inject(result.Cells, code, asStatic);
                    Post(new { type = "applied", count = result.Cells.Count, asStatic });
                }
            }
            catch (PExL.Core.Diagnostics.PExLException px)
            {
                Post(new { type = "error", message = px.Message, line = px.Line, column = px.Column });
            }
            catch (Exception ex)
            {
                Post(new { type = "error", message = ex.Message });
            }
        }

        /// <summary>
        /// Fired (on the Excel main thread) whenever the selection changes while
        /// the pane is open. In insert mode the address is dropped into the editor;
        /// otherwise we offer to load any PExL previously saved for that cell.
        /// </summary>
        private void OnExcelSelectionChanged(string address)
        {
            if (!_ready || IsDisposed) return;
            try
            {
                if (SelectionTracker.InsertMode)
                {
                    PostOnUi(new { type = "insertRef", address });
                    return;
                }

                // Recall only makes sense for a single cell.
                if (address.IndexOf(':') >= 0 || address.IndexOf(',') >= 0) return;

                dynamic app = ExcelDnaUtil.Application;
                dynamic range = app.Range[address];
                string? code = CodeStore.TryGet(app, range);
                if (!string.IsNullOrEmpty(code))
                    PostOnUi(new { type = "storedCode", address, code });
            }
            catch { /* COM busy / not a range - ignore */ }
        }

        private void PostOnUi(object payload)
        {
            if (IsDisposed) return;
            try
            {
                if (InvokeRequired) BeginInvoke(new Action(() => Post(payload)));
                else Post(payload);
            }
            catch { /* ignore */ }
        }

        private void Post(object payload)
        {
            if (!_ready || _web.CoreWebView2 == null) return;
            try { _web.CoreWebView2.PostWebMessageAsJson(JsonConvert.SerializeObject(payload)); } catch { /* ignore */ }
        }

        private static string FallbackHtml(string message) =>
            "<html><body style='font-family:Segoe UI;padding:16px;color:#444'>" +
            "<h3>PExL editor</h3><p>WebView2 could not start the rich editor.</p>" +
            "<pre style='white-space:pre-wrap;color:#a00'>" + System.Net.WebUtility.HtmlEncode(message) + "</pre>" +
            "</body></html>";

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                SelectionTracker.SelectionChanged -= OnExcelSelectionChanged;
                _web.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
