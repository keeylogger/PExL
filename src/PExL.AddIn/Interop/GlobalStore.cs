using System.Collections.Generic;
using System.Xml;

namespace PExL.AddIn.Interop
{
    /// <summary>One PExL global, as shown in the ShowGlobals() manager.</summary>
    public sealed class GlobalInfo
    {
        public string Name { get; set; } = "";
        public string RefersTo { get; set; } = "";
        public string Value { get; set; } = "";
    }

    /// <summary>
    /// Creates and manages PExL "globals" — native Excel Defined Names (workbook
    /// scope) that PExL also records in a CustomXMLPart registry. The registry is
    /// what lets the ShowGlobals() manager list and edit ONLY the names PExL
    /// created, never the user's own Excel named ranges. All ops are best-effort.
    /// </summary>
    public static class GlobalStore
    {
        private const string Ns = "urn:pexl:globals";

        /// <summary>Create or update a workbook-scoped global and register it.</summary>
        public static void Create(dynamic app, string name, string refersTo)
        {
            dynamic wb = app.ActiveWorkbook;
            // Replace any existing same-named global cleanly.
            try { wb.Names[name].Delete(); } catch { /* not present */ }
            wb.Names.Add(name, refersTo);
            Register(wb, name);
        }

        public static void CreateMany(dynamic app, IEnumerable<PExL.Core.GlobalDef> globals)
        {
            if (globals == null) return;
            foreach (var g in globals)
            {
                try { Create(app, g.Name, g.Formula); } catch { /* skip one bad name */ }
            }
        }

        /// <summary>List the PExL globals (intersection of registry and live names).</summary>
        public static List<GlobalInfo> List(dynamic app)
        {
            var list = new List<GlobalInfo>();
            try
            {
                dynamic wb = app.ActiveWorkbook;
                foreach (var name in RegisteredNames(wb))
                {
                    var info = new GlobalInfo { Name = name };
                    try { info.RefersTo = (string)wb.Names[name].RefersTo; }
                    catch { continue; } // dropped from Excel by the user — skip
                    info.Value = TryEvaluate(app, name);
                    list.Add(info);
                }
            }
            catch { /* ignore */ }
            return list;
        }

        /// <summary>Change what a PExL global refers to. RefersTo must include '='.</summary>
        public static bool Update(dynamic app, string name, string refersTo)
        {
            try
            {
                dynamic wb = app.ActiveWorkbook;
                if (!IsRegistered(wb, name)) return false;
                wb.Names[name].RefersTo = refersTo;
                return true;
            }
            catch { return false; }
        }

        /// <summary>Rename a PExL global; Excel updates dependent references.</summary>
        public static bool Rename(dynamic app, string oldName, string newName)
        {
            try
            {
                dynamic wb = app.ActiveWorkbook;
                if (!IsRegistered(wb, oldName)) return false;
                wb.Names[oldName].Name = newName; // renames and rewires references
                Unregister(wb, oldName);
                Register(wb, newName);
                return true;
            }
            catch { return false; }
        }

        /// <summary>Delete a PExL global (only if PExL created it).</summary>
        public static bool Delete(dynamic app, string name)
        {
            try
            {
                dynamic wb = app.ActiveWorkbook;
                if (!IsRegistered(wb, name)) return false;
                try { wb.Names[name].Delete(); } catch { /* already gone */ }
                Unregister(wb, name);
                return true;
            }
            catch { return false; }
        }

        // ---- registry (CustomXMLPart) ----

        private static string TryEvaluate(dynamic app, string name)
        {
            try
            {
                dynamic v = app.Evaluate(name);
                if (v == null) return "";
                string s = v.ToString();
                return s.Length > 120 ? s.Substring(0, 120) + "…" : s;
            }
            catch { return ""; }
        }

        private static IEnumerable<string> RegisteredNames(dynamic wb)
        {
            var doc = LoadDoc(wb);
            var names = new List<string>();
            foreach (XmlNode node in doc.DocumentElement!.ChildNodes)
                if (node is XmlElement el && el.GetAttribute("name").Length > 0)
                    names.Add(el.GetAttribute("name"));
            return names;
        }

        private static bool IsRegistered(dynamic wb, string name)
        {
            foreach (var n in RegisteredNames(wb))
                if (string.Equals(n, name, System.StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        private static void Register(dynamic wb, string name)
        {
            var doc = LoadDoc(wb);
            var root = doc.DocumentElement!;
            foreach (XmlNode node in root.ChildNodes)
                if (node is XmlElement el && string.Equals(el.GetAttribute("name"), name, System.StringComparison.OrdinalIgnoreCase))
                    return; // already there
            var add = doc.CreateElement("global", Ns);
            add.SetAttribute("name", name);
            root.AppendChild(add);
            ReplacePart(wb, doc.OuterXml);
        }

        private static void Unregister(dynamic wb, string name)
        {
            var doc = LoadDoc(wb);
            var root = doc.DocumentElement!;
            XmlElement? target = null;
            foreach (XmlNode node in root.ChildNodes)
                if (node is XmlElement el && string.Equals(el.GetAttribute("name"), name, System.StringComparison.OrdinalIgnoreCase))
                { target = el; break; }
            if (target != null)
            {
                root.RemoveChild(target);
                ReplacePart(wb, doc.OuterXml);
            }
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
            catch { /* fall through */ }

            var root = doc.CreateElement("globals", Ns);
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
