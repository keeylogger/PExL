using System.Xml;

namespace PExL.AddIn.Interop
{
    /// <summary>
    /// Persists the original PExL source per cell inside the workbook as a
    /// CustomXMLPart, so selecting a cell can repopulate the editor with the
    /// readable code instead of the raw formula. All operations are best-effort
    /// and never throw into Excel.
    /// </summary>
    public static class CodeStore
    {
        private const string Ns = "urn:pexl:code-store";

        public static void Save(dynamic app, dynamic range, string code)
        {
            try
            {
                string key = KeyFor(range);
                dynamic wb = app.ActiveWorkbook;
                var doc = LoadDoc(wb);
                var root = doc.DocumentElement!;

                XmlElement? existing = null;
                foreach (XmlNode node in root.ChildNodes)
                {
                    if (node is XmlElement el && el.GetAttribute("addr") == key) { existing = el; break; }
                }
                if (existing == null)
                {
                    existing = doc.CreateElement("cell", Ns);
                    existing.SetAttribute("addr", key);
                    root.AppendChild(existing);
                }
                existing.InnerText = code;

                ReplacePart(wb, doc.OuterXml);
            }
            catch { /* best effort */ }
        }

        public static string? TryGet(dynamic app, dynamic range)
        {
            try
            {
                string key = KeyFor(range);
                dynamic wb = app.ActiveWorkbook;
                var doc = LoadDoc(wb);
                foreach (XmlNode node in doc.DocumentElement!.ChildNodes)
                {
                    if (node is XmlElement el && el.GetAttribute("addr") == key)
                        return el.InnerText;
                }
            }
            catch { /* ignore */ }
            return null;
        }

        private static string KeyFor(dynamic range)
        {
            string sheet = range.Worksheet.Name;
            string addr = range.Address[false, false];
            return sheet + "!" + addr;
        }

        private static XmlDocument LoadDoc(dynamic wb)
        {
            var doc = new XmlDocument();
            try
            {
                dynamic parts = wb.CustomXMLParts.SelectByNamespace(Ns);
                if (parts.Count > 0)
                {
                    string xml = parts[1].XML;
                    doc.LoadXml(xml);
                    return doc;
                }
            }
            catch { /* fall through to fresh doc */ }

            var root = doc.CreateElement("pexl", Ns);
            doc.AppendChild(root);
            return doc;
        }

        private static void ReplacePart(dynamic wb, string xml)
        {
            try
            {
                dynamic parts = wb.CustomXMLParts.SelectByNamespace(Ns);
                while (parts.Count > 0)
                {
                    parts[1].Delete();
                    parts = wb.CustomXMLParts.SelectByNamespace(Ns);
                }
            }
            catch { /* ignore */ }
            wb.CustomXMLParts.Add(xml, System.Type.Missing);
        }
    }
}
