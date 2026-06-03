using System.Collections.Generic;
using ExcelDna.Integration;

namespace PExL.AddIn.Interop
{
    /// <summary>
    /// Creates throwaway worksheets with dummy data tailored to a documentation
    /// example, so users can press "Try it yourself!" and immediately experiment.
    /// </summary>
    public static class DemoData
    {
        private sealed class DemoSpec
        {
            public string Title = "PExL Demo";
            public List<KeyValuePair<string, object>> Cells = new List<KeyValuePair<string, object>>();
            public string Select = "A1";
            public string NoteCell = "A1";
        }

        public static void Create(string demo, string code)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic wb = app.ActiveWorkbook;
                    if (wb == null) wb = app.Workbooks.Add();

                    var spec = Spec(demo ?? string.Empty);

                    dynamic ws = wb.Worksheets.Add();
                    try { ws.Name = UniqueName(wb, spec.Title); } catch { /* keep default name */ }

                    foreach (var c in spec.Cells)
                        ws.Range[c.Key].Value2 = c.Value;

                    if (!string.IsNullOrEmpty(code))
                    {
                        string oneLine = code.Replace("\r", " ").Replace("\n", "  ");
                        ws.Range[spec.NoteCell].Value2 = "Try in the editor:  " + oneLine;
                    }

                    ws.Activate();
                    try { ws.Range[spec.Select].Select(); } catch { /* ignore */ }
                    try { ws.Columns.AutoFit(); } catch { /* ignore */ }
                }
                catch
                {
                    // Best effort; nothing to surface.
                }
            });
        }

        private static string UniqueName(dynamic wb, string baseName)
        {
            string name = Trunc(baseName);
            int n = 2;
            while (Exists(wb, name))
            {
                name = Trunc(baseName + " " + n);
                n++;
            }
            return name;
        }

        private static string Trunc(string s) => s.Length > 31 ? s.Substring(0, 31) : s;

        private static bool Exists(dynamic wb, string name)
        {
            foreach (dynamic ws in wb.Worksheets)
            {
                try { if (string.Equals((string)ws.Name, name, System.StringComparison.OrdinalIgnoreCase)) return true; }
                catch { }
            }
            return false;
        }

        private static DemoSpec Cell(this DemoSpec s, string addr, object val)
        {
            s.Cells.Add(new KeyValuePair<string, object>(addr, val));
            return s;
        }

        /// <summary>Write a column of values down from <paramref name="startRow"/>.</summary>
        private static DemoSpec Col(this DemoSpec s, string col, int startRow, params object[] vals)
        {
            for (int i = 0; i < vals.Length; i++)
                s.Cells.Add(new KeyValuePair<string, object>(col + (startRow + i), vals[i]));
            return s;
        }

        private static DemoSpec Spec(string demo)
        {
            var s = new DemoSpec { NoteCell = "H1" };
            switch (demo.ToLowerInvariant())
            {
                case "trim":
                case "proper":
                    s.Title = "PExL Demo - Text";
                    s.Cell("A1", "Raw name").Cell("B1", "Cleaned");
                    s.Col("A", 2,
                        "  john   SMITH ", "MARY o'brien ", " josé  garcía ", "anne-marie JONES",
                        "  PETER o'toole", "li  wei ", " ahmed   KHAN", "Sofia ROSSI ",
                        "  hans müller ", "olga  IVANOVA", " kwame mensah ", "  emma   THOMPSON");
                    s.Select = "B2";
                    break;

                case "email":
                    s.Title = "PExL Demo - Email";
                    s.Cell("A1", "Email").Cell("B1", "Domain");
                    s.Col("A", 2,
                        "alice.morgan@contoso.com", "ben.carter@northwind.co.uk", "carla.dias@fabrikam.com",
                        "david.lee@adventure-works.com", "elena.petrova@contoso.com", "frank.mueller@fabrikam.de",
                        "grace.kim@northwind.co.uk", "hassan.ali@contoso.com", "ingrid.olsen@fabrikam.com",
                        "james.wright@adventure-works.com", "keiko.tanaka@contoso.jp", "luis.fernandez@fabrikam.es");
                    s.Select = "B2";
                    break;

                case "split":
                    s.Title = "PExL Demo - Split";
                    s.Cell("A1", "Order code").Cell("B1", "Region").Cell("C1", "Number");
                    s.Col("A", 2,
                        "WEST-2024-0017", "EAST-2024-0093", "NORTH-2023-0481", "SOUTH-2024-0142",
                        "WEST-2023-0205", "EAST-2024-0310", "NORTH-2024-0056", "SOUTH-2023-0729",
                        "WEST-2024-0088", "EAST-2023-0164", "NORTH-2024-0402", "SOUTH-2024-0011");
                    s.Select = "A2";
                    break;

                case "csv":
                    s.Title = "PExL Demo - CSV";
                    s.Cell("A1", "Pasted row (Product, Sales, Region, Status)");
                    s.Col("A", 2,
                        "Widget, 1200, West, In stock", "Gadget, 840, East, Low stock",
                        "Gizmo, 1500, North, In stock", "Doohickey, 400, South, Backorder",
                        "Sprocket, 980, West, In stock", "Cog, 1320, East, In stock",
                        "Lever, 220, North, Low stock", "Piston, 760, South, In stock",
                        "Valve, 1610, West, In stock", "Gasket, 530, East, Backorder",
                        "Bearing, 1180, North, In stock", "Flange, 690, South, Low stock");
                    s.Select = "A2";
                    break;

                case "combine":
                    s.Title = "PExL Demo - Combine";
                    s.Cell("A1", "First").Cell("B1", "Last").Cell("C1", "Full name");
                    s.Col("A", 2,
                        "Ada", "Alan", "Grace", "Linus", "Margaret", "Dennis",
                        "Katherine", "Tim", "Barbara", "Edsger", "Donald", "Radia");
                    s.Col("B", 2,
                        "Lovelace", "Turing", "Hopper", "Torvalds", "Hamilton", "Ritchie",
                        "Johnson", "Berners-Lee", "Liskov", "Dijkstra", "Knuth", "Perlman");
                    s.Select = "C2";
                    break;

                case "unique":
                    s.Title = "PExL Demo - Unique";
                    s.Cell("A1", "Region").Cell("C1", "Distinct");
                    s.Col("A", 2,
                        "West", "East", "West", "North", "South", "East",
                        "West", "North", "East", "South", "West", "North");
                    s.Select = "A2:A13";
                    break;

                case "sort":
                case "topn":
                    s.Title = "PExL Demo - Sort";
                    s.Cell("A1", "Product").Cell("B1", "Sales").Cell("D1", "Result");
                    s.Col("A", 2,
                        "Widget", "Gadget", "Gizmo", "Doohickey", "Sprocket", "Cog",
                        "Lever", "Piston", "Valve", "Gasket", "Bearing", "Flange");
                    s.Col("B", 2, 1200, 840, 1500, 400, 980, 1320, 220, 760, 1610, 530, 1180, 690);
                    s.Select = "A2:B13";
                    break;

                case "filter":
                    s.Title = "PExL Demo - Filter";
                    s.Cell("A1", "Region").Cell("B1", "Amount").Cell("D1", "West only");
                    s.Col("A", 2,
                        "West", "East", "West", "North", "South", "West",
                        "East", "North", "West", "South", "East", "West");
                    s.Col("B", 2, 120, 90, 200, 60, 150, 175, 80, 110, 240, 95, 130, 210);
                    s.Select = "A2:B13";
                    break;

                case "lookup":
                    s.Title = "PExL Demo - Lookup";
                    s.Cell("A1", "Lookup id").Cell("A2", "P-107").Cell("B1", "Found name");
                    s.Cell("C1", "Product id").Cell("D1", "Product name").Cell("E1", "Unit price");
                    s.Col("C", 2,
                        "P-101", "P-102", "P-103", "P-104", "P-105", "P-106",
                        "P-107", "P-108", "P-109", "P-110", "P-111", "P-112");
                    s.Col("D", 2,
                        "Widget", "Gadget", "Gizmo", "Doohickey", "Sprocket", "Cog",
                        "Lever", "Piston", "Valve", "Gasket", "Bearing", "Flange");
                    s.Col("E", 2, 12.5, 8.0, 19.99, 4.25, 9.8, 13.2, 2.2, 7.6, 16.1, 5.3, 11.8, 6.9);
                    s.Select = "A2";
                    break;

                case "check":
                    s.Title = "PExL Demo - Check";
                    s.Cell("A1", "Score").Cell("B1", "Grade");
                    s.Col("A", 2, 95, 82, 71, 55, 88, 64, 77, 91, 49, 68, 83, 73);
                    s.Select = "A2";
                    break;

                case "sumwhere":
                case "quickstats":
                    s.Title = "PExL Demo - Summarize";
                    s.Cell("A1", "Region").Cell("B1", "Amount").Cell("D1", "Result");
                    s.Col("A", 2,
                        "West", "East", "West", "North", "West", "South",
                        "East", "West", "North", "East", "West", "South");
                    s.Col("B", 2, 120, 90, 200, 60, 175, 80, 110, 240, 95, 130, 210, 70);
                    s.Select = "B2:B13";
                    break;

                case "safedivide":
                    s.Title = "PExL Demo - Divide";
                    s.Cell("A1", "Total").Cell("B1", "Count").Cell("C1", "Average each");
                    s.Col("A", 2, 100, 50, 240, 0, 360, 90, 0, 180, 500, 75);
                    s.Col("B", 2, 4, 0, 6, 3, 0, 5, 0, 9, 10, 0);
                    s.Select = "C2";
                    break;

                case "fillblanks":
                    s.Title = "PExL Demo - Fill";
                    s.Cell("A1", "Region").Cell("B1", "Filled");
                    // Grouped report: the region is only written on its first row.
                    s.Col("A", 2,
                        "West", "", "", "East", "", "North",
                        "", "", "South", "", "", "West");
                    s.Select = "B3";
                    break;

                case "datediff":
                    s.Title = "PExL Demo - Dates";
                    s.Cell("A1", "Start date").Cell("B1", "Days to today");
                    s.Col("A", 2,
                        "=DATE(2024,1,15)", "=DATE(2023,6,30)", "=DATE(2022,11,5)", "=DATE(2024,3,22)",
                        "=DATE(2021,9,1)", "=DATE(2023,12,25)", "=DATE(2024,7,8)", "=DATE(2020,2,29)",
                        "=DATE(2023,4,14)", "=DATE(2022,8,19)");
                    s.Select = "B2";
                    break;

                case "today":
                    s.Title = "PExL Demo - Today";
                    s.Cell("A1", "Today goes here ->");
                    s.Select = "B1";
                    break;

                default:
                    s.Title = "PExL Demo";
                    s.Cell("A1", "Value")
                     .Col("A", 2, "hello world", 42, "WEST", "2024-01-15", "alice@contoso.com");
                    s.Select = "A2";
                    break;
            }
            return s;
        }
    }
}
