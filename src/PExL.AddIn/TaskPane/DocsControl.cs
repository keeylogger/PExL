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
    /// WebView2-hosted documentation pane: searchable cards with a live "Run"
    /// playground (compile preview only) and a "Try it yourself!" demo-sheet
    /// generator. Must be COM-visible so Excel-DNA can host it in a task pane.
    /// </summary>
    [ComVisible(true)]
    public sealed class DocsControl : UserControl
    {
        private readonly WebView2 _web;
        private bool _ready;

        public DocsControl()
        {
            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);
            _ = InitializeAsync();
        }

        private async System.Threading.Tasks.Task InitializeAsync()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? AppDomain.CurrentDomain.BaseDirectory;

                string nativeDir = Path.Combine(baseDir, "runtimes", "win-x64", "native");
                if (Directory.Exists(nativeDir))
                {
                    try { CoreWebView2Environment.SetLoaderDllFolderPath(nativeDir); } catch { /* best effort */ }
                }

                string userData = Path.Combine(Path.GetTempPath(), "PExL.WebView2");
                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                await _web.EnsureCoreWebView2Async(env);

                _web.CoreWebView2.WebMessageReceived += OnWebMessage;

                string? docsPath = ResolveDocsHtml(baseDir);
                if (docsPath != null)
                    _web.CoreWebView2.Navigate(new Uri(docsPath).AbsoluteUri);
                else
                    _web.CoreWebView2.NavigateToString(FallbackHtml(
                        "Documentation assets (web/docs.html) were not found next to the add-in."));

                _ready = true;
            }
            catch (Exception ex)
            {
                try { _web.CoreWebView2?.NavigateToString(FallbackHtml(ex.Message)); } catch { /* ignore */ }
            }
        }

        private static string? ResolveDocsHtml(string baseDir)
        {
            string[] candidates =
            {
                Path.Combine(baseDir, "web", "docs.html"),
                Path.Combine(baseDir, "..", "web", "docs.html"),
                Path.Combine(baseDir, "..", "..", "PExL.Editor.Web", "docs.html"),
                Path.Combine(baseDir, "..", "..", "..", "..", "PExL.Editor.Web", "docs.html"),
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
                        HandleCompile((string?)root["id"] ?? string.Empty, (string?)root["code"] ?? string.Empty, apply: false);
                        break;
                    case "apply":
                        HandleCompile((string?)root["id"] ?? string.Empty, (string?)root["code"] ?? string.Empty, apply: true);
                        break;
                    case "tryIt":
                        DemoData.Create((string?)root["demo"] ?? string.Empty, (string?)root["code"] ?? string.Empty);
                        break;
                }
            }
            catch (Exception ex)
            {
                Post(new { type = "error", message = ex.Message });
            }
        }

        /// <summary>
        /// Compile a snippet. When <paramref name="apply"/> is false this is a
        /// preview only ("View Excel formula"); when true the formula is written
        /// to the active sheet ("Run", same as the editor's Apply &#8594; cell).
        /// </summary>
        private void HandleCompile(string id, string code, bool apply)
        {
            try
            {
                var result = Transpiler.Transpile(code);

                if (apply)
                {
                    ExcelInjector.Inject(result.Cells, code);
                    Post(new { type = "applied", id, count = result.Cells.Count });
                    return;
                }

                var cells = new object[result.Cells.Count];
                for (int i = 0; i < result.Cells.Count; i++)
                    cells[i] = new { target = result.Cells[i].Target, formula = result.Cells[i].Formula };
                Post(new { type = "preview", id, cells });
            }
            catch (PExL.Core.Diagnostics.PExLException px)
            {
                Post(new { type = "error", id, message = px.Message, line = px.Line, column = px.Column });
            }
            catch (Exception ex)
            {
                Post(new { type = "error", id, message = ex.Message });
            }
        }

        private void Post(object payload)
        {
            if (!_ready || _web.CoreWebView2 == null) return;
            try { _web.CoreWebView2.PostWebMessageAsJson(JsonConvert.SerializeObject(payload)); } catch { /* ignore */ }
        }

        private static string FallbackHtml(string message) =>
            "<html><body style='font-family:Segoe UI;padding:16px;color:#444'>" +
            "<h3>PExL documentation</h3><p>WebView2 could not start the docs pane.</p>" +
            "<pre style='white-space:pre-wrap;color:#a00'>" + System.Net.WebUtility.HtmlEncode(message) + "</pre>" +
            "</body></html>";

        protected override void Dispose(bool disposing)
        {
            if (disposing) _web.Dispose();
            base.Dispose(disposing);
        }
    }
}
