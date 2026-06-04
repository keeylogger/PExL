using System;
using System.Collections.Generic;

namespace PExL.Core.StdLib
{
    /// <summary>
    /// Knows the PExL verb vocabulary: canonical verb names, their curated
    /// synonyms, and prepositional-label aliases. The parser uses it to tell a
    /// verb from a bound name; the emitter switches on the canonical name.
    /// </summary>
    public static class VerbRegistry
    {
        // canonical verb -> synonyms (all compared case-insensitively)
        private static readonly Dictionary<string, string[]> Verbs = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            // text
            ["split"] = new[] { "divide", "separate" },
            ["fromLeft"] = new[] { "left", "before" },
            ["fromRight"] = new[] { "right", "after" },
            ["at"] = new[] { "piece", "nth" },
            ["spill"] = new[] { "all", "expand" },
            ["combine"] = new[] { "join", "concat", "concatenate", "merge" },
            ["trim"] = Array.Empty<string>(),
            ["clean"] = Array.Empty<string>(),
            ["upper"] = new[] { "uppercase" },
            ["lower"] = new[] { "lowercase" },
            ["proper"] = new[] { "titlecase" },
            ["replace"] = new[] { "substitute" },
            ["contains"] = new[] { "has", "includes" },
            ["startsWith"] = new[] { "beginsWith" },
            ["endsWith"] = Array.Empty<string>(),
            ["length"] = new[] { "len", "size" },

            // lookup
            ["find"] = new[] { "lookup", "search", "vlookup", "xlookup" },
            ["position"] = new[] { "indexOf", "match" },

            // logic
            ["ifError"] = new[] { "orElse", "onError" },
            ["ifs"] = Array.Empty<string>(), // internal (check blocks)

            // aggregation
            ["sum"] = new[] { "total" },
            ["avg"] = new[] { "average", "mean" },
            ["min"] = Array.Empty<string>(),
            ["max"] = Array.Empty<string>(),
            ["count"] = Array.Empty<string>(),
            ["countNum"] = new[] { "countNumbers" },
            ["sumWhere"] = new[] { "sumIf" },
            ["countWhere"] = new[] { "countIf" },
            ["avgWhere"] = new[] { "averageIf", "avgIf" },

            // dates
            ["today"] = Array.Empty<string>(),
            ["now"] = Array.Empty<string>(),
            ["addDays"] = Array.Empty<string>(),
            ["addMonths"] = Array.Empty<string>(),
            ["addYears"] = Array.Empty<string>(),
            ["yearOf"] = new[] { "year" },
            ["monthOf"] = new[] { "month" },
            ["dayOf"] = new[] { "day" },
            ["weekdayOf"] = new[] { "weekday" },
            ["dateDiff"] = new[] { "datedif" },

            // math
            ["round"] = Array.Empty<string>(),
            ["abs"] = Array.Empty<string>(),
            ["sqrt"] = Array.Empty<string>(),
            ["power"] = new[] { "pow" },
            ["mod"] = new[] { "modulo", "remainder" },

            // filter / shape
            ["filter"] = Array.Empty<string>(),
            ["sort"] = Array.Empty<string>(),
            ["unique"] = new[] { "distinct" },
            ["take"] = new[] { "first" },

            // references / literals
            ["col"] = new[] { "column" },
            ["row"] = Array.Empty<string>(),
            ["cell"] = Array.Empty<string>(),
            ["fixed"] = new[] { "lock", "absolute" },
            ["Date"] = Array.Empty<string>(),

            // escape hatches
            ["raw"] = new[] { "formula" },
            ["legacy"] = Array.Empty<string>(),

            // globals (document-level variables) + console commands
            ["makeGlobal"] = new[] { "global", "defineGlobal", "setGlobal" },
            ["showGlobals"] = new[] { "showGlobal", "globals", "listGlobals" },
        };

        // prepositional-label aliases -> canonical label
        private static readonly Dictionary<string, string> Prepositions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["within"] = "within",
            ["inRange"] = "within",
            ["in"] = "within",
            ["thenReturn"] = "thenReturn",
            ["returnFrom"] = "thenReturn",
            ["bringBack"] = "thenReturn",
            ["getFrom"] = "thenReturn",
            ["with"] = "with",
            ["by"] = "by",
            ["from"] = "from",
            ["where"] = "where",
            ["ifMissing"] = "ifMissing",
            ["orMissing"] = "ifMissing",
            ["default"] = "ifMissing",
            ["inTable"] = "inTable",
            ["returnColumn"] = "returnColumn",
        };

        private static readonly Dictionary<string, string> AliasToCanonical = BuildAliasMap();

        private static Dictionary<string, string> BuildAliasMap()
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in Verbs)
            {
                map[kv.Key] = kv.Key;
                foreach (var syn in kv.Value)
                    map[syn] = kv.Key;
            }
            return map;
        }

        public static bool IsVerb(string name) => name != null && AliasToCanonical.ContainsKey(name);

        public static string Canonical(string name)
            => name != null && AliasToCanonical.TryGetValue(name, out var c) ? c : (name ?? string.Empty);

        public static bool IsPreposition(string name) => name != null && Prepositions.ContainsKey(name);

        public static string CanonicalPreposition(string name)
            => name != null && Prepositions.TryGetValue(name, out var c) ? c : (name ?? string.Empty);
    }
}
