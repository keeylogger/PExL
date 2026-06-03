using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace PExL.Core.Decompile
{
    /// <summary>
    /// Turns a parsed Excel formula tree (<see cref="FNode"/>) back into readable
    /// PExL. Known Excel functions map to PExL verbs; anything unrecognized falls
    /// back to the <c>raw("FUNC", ...)</c> / <c>legacy.*</c> escape hatches so the
    /// result always recompiles to an equivalent formula.
    /// </summary>
    internal sealed class PExLWriter
    {
        public string Write(FNode n)
        {
            switch (n)
            {
                case FNum x: return x.Raw;
                case FStr s: return Quote(s.Value);
                case FBool b: return b.Value ? "true" : "false";
                case FRef r: return r.Text;
                case FName nm: return nm.Name;
                case FUnary u: return WriteUnary(u);
                case FPostfix p: return WritePostfix(p);
                case FBinary bin: return WriteBinary(bin);
                case FCall c: return WriteCall(c);
                case FArray a: return WriteArray(a);
                default: return "/* ? */";
            }
        }

        // ---------------- operators ----------------

        private string WriteUnary(FUnary u) => "-" + ParenAtom(u.Operand);

        private string WritePostfix(FPostfix p)
        {
            if (p.Operand is FNum num) return num.Raw + "%";
            return "(" + Write(p.Operand) + ")*0.01";
        }

        private static int Prec(string op)
        {
            switch (op)
            {
                case "or": return 1;
                case "and": return 2;
                case "=": case "<>": case ">": case "<": case ">=": case "<=": return 3;
                case "+": case "-": return 4;
                case "*": case "/": return 5;
                case "^": return 6;
                default: return 7;
            }
        }

        private string Paren(FNode child, int parentPrec)
        {
            string s = Write(child);
            if (child is FBinary cb && Prec(cb.Op) < parentPrec) return "(" + s + ")";
            return s;
        }

        private string ParenAtom(FNode child)
        {
            string s = Write(child);
            if (child is FBinary || child is FUnary) return "(" + s + ")";
            return s;
        }

        private string WriteBinary(FBinary b)
        {
            // string concatenation -> combine(...)
            if (b.Op == "&")
            {
                var parts = new List<FNode>();
                Flatten(b, "&", parts);
                var sb = new StringBuilder("combine(");
                for (int i = 0; i < parts.Count; i++)
                {
                    if (i > 0) sb.Append(", ");
                    sb.Append(Write(parts[i]));
                }
                return sb.Append(")").ToString();
            }

            // startsWith / endsWith: LEFT(t,k)=p  /  RIGHT(t,k)=p
            if (b.Op == "=")
            {
                var edge = TryEdge(b.Left, b.Right) ?? TryEdge(b.Right, b.Left);
                if (edge != null) return edge;
            }

            int p = Prec(b.Op);
            string l = Paren(b.Left, p);
            string r = Paren(b.Right, p + 1);
            return l + " " + b.Op + " " + r;
        }

        private string? TryEdge(FNode maybeCall, FNode probe)
        {
            if (maybeCall is FCall c && c.Args.Count == 2 && probe is FStr ps)
            {
                bool left = string.Equals(c.Name, "LEFT", StringComparison.OrdinalIgnoreCase);
                bool right = string.Equals(c.Name, "RIGHT", StringComparison.OrdinalIgnoreCase);
                if ((left || right) && IsIntLiteral(c.Args[1], out int k) && k == ps.Value.Length)
                {
                    string verb = left ? "startsWith" : "endsWith";
                    return verb + "(" + Write(c.Args[0]) + ", " + Quote(ps.Value) + ")";
                }
            }
            return null;
        }

        private static void Flatten(FNode n, string op, List<FNode> into)
        {
            if (n is FBinary b && b.Op == op) { Flatten(b.Left, op, into); Flatten(b.Right, op, into); }
            else into.Add(n);
        }

        // ---------------- function calls ----------------

        private string WriteCall(FCall c)
        {
            string u = c.Name.ToUpperInvariant();
            var a = c.Args;

            switch (u)
            {
                // text
                case "TRIM": return Fn("trim", a);
                case "CLEAN": return Fn("clean", a);
                case "UPPER": return Fn("upper", a);
                case "LOWER": return Fn("lower", a);
                case "PROPER": return Fn("proper", a);
                case "LEN": return Fn("length", a);

                case "SUBSTITUTE":
                    if (a.Count == 4)
                    {
                        if (IsIntLiteral(a[3], out int nth) && nth == 1)
                            return "replace.first(" + W(a[0]) + ", " + W(a[1]) + ", " + W(a[2]) + ")";
                        return "replace.nth(" + W(a[3]) + ", " + W(a[0]) + ", " + W(a[1]) + ", " + W(a[2]) + ")";
                    }
                    if (a.Count == 3) return "replace(" + W(a[0]) + ", " + W(a[1]) + ", " + W(a[2]) + ")";
                    break;

                case "ISNUMBER":
                    if (a.Count == 1 && a[0] is FCall inner && inner.Args.Count == 2)
                    {
                        if (string.Equals(inner.Name, "SEARCH", StringComparison.OrdinalIgnoreCase))
                            return "contains(" + W(inner.Args[1]) + ", " + W(inner.Args[0]) + ")";
                        if (string.Equals(inner.Name, "FIND", StringComparison.OrdinalIgnoreCase))
                            return "contains.caseSensitive(" + W(inner.Args[1]) + ", " + W(inner.Args[0]) + ")";
                    }
                    break;

                case "TEXTBEFORE": return TextSide(a, left: true);
                case "TEXTAFTER": return TextSide(a, left: false);
                case "TEXTSPLIT":
                    if (a.Count == 2) return "split(" + W(a[0]) + ", " + W(a[1]) + ") |> spill";
                    break;
                case "INDEX":
                    if (a.Count == 2 && a[0] is FCall ts && string.Equals(ts.Name, "TEXTSPLIT", StringComparison.OrdinalIgnoreCase) && ts.Args.Count == 2)
                        return "split(" + W(ts.Args[0]) + ", " + W(ts.Args[1]) + ") |> at(" + W(a[1]) + ")";
                    break;

                case "TEXTJOIN":
                    if (a.Count >= 3 && a[1] is FBool tj && tj.Value)
                    {
                        var args = a.GetRange(2, a.Count - 2);
                        string body = "combine(" + JoinArgs(args) + ")";
                        if (a[0] is FStr sep && sep.Value.Length == 0) return body;
                        return body + " with(" + W(a[0]) + ")";
                    }
                    break;

                // lookup / logic
                case "XLOOKUP": return XLookup(a);
                case "MATCH":
                    if (a.Count == 3 && IsIntLiteral(a[2], out int mm) && mm == 0)
                        return "position(" + W(a[0]) + ", " + W(a[1]) + ")";
                    break;
                case "IF": return WriteIf(a);
                case "IFS": return Ifs(a);
                case "IFERROR":
                    if (a.Count == 2) return "ifError(" + W(a[0]) + ", " + W(a[1]) + ")";
                    break;
                case "AND": return BoolJoin(a, "and");
                case "OR": return BoolJoin(a, "or");
                case "NOT":
                    if (a.Count == 1) return "not " + ParenAtom(a[0]);
                    break;

                // aggregation
                case "SUM": return Fn("sum", a);
                case "AVERAGE": return Fn("avg", a);
                case "MIN": return Fn("min", a);
                case "MAX": return Fn("max", a);
                case "COUNTA": return Fn("count", a);
                case "COUNT": return Fn("countNum", a);
                case "AGGREGATE": return Aggregate(a);
                case "SUMIFS": return Where("sumWhere", a, hasValueRange: true);
                case "COUNTIFS": return Where("countWhere", a, hasValueRange: false);
                case "AVERAGEIFS": return Where("avgWhere", a, hasValueRange: true);

                // dates / math
                case "EDATE":
                    if (a.Count == 2) return "addMonths(" + W(a[0]) + ", " + W(a[1]) + ")";
                    break;
                case "YEAR": return Fn("yearOf", a);
                case "MONTH": return Fn("monthOf", a);
                case "DAY": return Fn("dayOf", a);
                case "WEEKDAY": return Fn("weekdayOf", a);
                case "DATEDIF": return DateDiff(a);
                case "ROUND": return Round("round", a);
                case "ROUNDUP": return Round("round.up", a);
                case "ROUNDDOWN": return Round("round.down", a);
                case "ABS": return Fn("abs", a);
                case "SQRT": return Fn("sqrt", a);
                case "POWER": return Fn("power", a);
                case "MOD": return Fn("mod", a);
                case "TODAY": return "today()";
                case "NOW": return "now()";
                case "DATE": return DateLit(a);

                // filter / shape
                case "FILTER":
                    if (a.Count == 2) return "filter(" + W(a[0]) + ") where(" + W(a[1]) + ")";
                    break;
                case "SORT": return Sort(a);
                case "UNIQUE": return Fn("unique", a);
                case "TAKE": return Fn("take", a);

                // legacy lookups
                case "VLOOKUP": return Legacy("vlookup", a);
                case "HLOOKUP": return Legacy("hlookup", a);
            }

            // fallback: emit verbatim through the escape hatch.
            return Raw(c.Name, a);
        }

        private string W(FNode n) => Write(n);

        private string Fn(string verb, List<FNode> a) => verb + "(" + JoinArgs(a) + ")";

        private string JoinArgs(List<FNode> a)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < a.Count; i++) { if (i > 0) sb.Append(", "); sb.Append(Write(a[i])); }
            return sb.ToString();
        }

        private string TextSide(List<FNode> a, bool left)
        {
            // TEXTBEFORE(t,d) / TEXTAFTER(t,d) -> split.First; with instance -1 -> split.Last
            if (a.Count == 2)
                return "split.First(" + W(a[0]) + ", " + W(a[1]) + ") |> " + (left ? "fromLeft" : "fromRight");
            if (a.Count >= 3 && IsIntLiteral(a[2], out int inst) && inst == -1)
                return "split.Last(" + W(a[0]) + ", " + W(a[1]) + ") |> " + (left ? "fromLeft" : "fromRight");
            return Raw(left ? "TEXTBEFORE" : "TEXTAFTER", a);
        }

        private string XLookup(List<FNode> a)
        {
            if (a.Count < 3) return Raw("XLOOKUP", a);
            string mod = "";
            string ifMissing = "";
            // [0]k [1]la [2]ra [3]ifNotFound [4]matchMode [5]searchMode
            if (a.Count >= 6 && IsIntLiteral(a[5], out int sm) && sm == -1) mod = ".reverse";
            else if (a.Count >= 5 && IsIntLiteral(a[4], out int matchMode))
            {
                if (matchMode == 2) mod = ".wildcard";
                else if (matchMode == -1) mod = ".approx";
            }
            if (a.Count >= 4 && !(a[3] is FStr es && es.Value.Length == 0))
                ifMissing = " ifMissing " + W(a[3]);

            return "find" + mod + " " + W(a[0]) + " within " + W(a[1]) + " thenReturn " + W(a[2]) + ifMissing;
        }

        private string WriteIf(List<FNode> a)
        {
            // A nested IF ladder reads far better as a check block.
            if (a.Count == 3 && a[2] is FCall elseIf &&
                string.Equals(elseIf.Name, "IF", StringComparison.OrdinalIgnoreCase) && elseIf.Args.Count == 3)
            {
                var pairs = new List<KeyValuePair<FNode, FNode>> { new KeyValuePair<FNode, FNode>(a[0], a[1]) };
                FNode cur = a[2];
                while (cur is FCall c && string.Equals(c.Name, "IF", StringComparison.OrdinalIgnoreCase) && c.Args.Count == 3)
                {
                    pairs.Add(new KeyValuePair<FNode, FNode>(c.Args[0], c.Args[1]));
                    cur = c.Args[2];
                }
                return CheckBlock(pairs, cur);
            }
            if (a.Count == 2) return "if " + W(a[0]) + " then " + W(a[1]);
            if (a.Count == 3) return "if " + W(a[0]) + " then " + W(a[1]) + " else " + W(a[2]);
            return Raw("IF", a);
        }

        private string Ifs(List<FNode> a)
        {
            if (a.Count >= 2 && a.Count % 2 == 0)
            {
                var pairs = new List<KeyValuePair<FNode, FNode>>();
                FNode? elseResult = null;
                for (int i = 0; i + 1 < a.Count; i += 2)
                {
                    if (a[i] is FBool b && b.Value) { elseResult = a[i + 1]; break; }
                    pairs.Add(new KeyValuePair<FNode, FNode>(a[i], a[i + 1]));
                }
                if (pairs.Count > 0) return CheckBlock(pairs, elseResult);
            }
            return Raw("IFS", a);
        }

        /// <summary>Render a multi-branch check block (no subject, no leading 'if').</summary>
        private string CheckBlock(List<KeyValuePair<FNode, FNode>> pairs, FNode? elseResult)
        {
            var sb = new StringBuilder("check\n");
            foreach (var p in pairs)
                sb.Append("  ").Append(Write(p.Key)).Append(" then ").Append(Write(p.Value)).Append('\n');
            if (elseResult != null)
                sb.Append("  else ").Append(Write(elseResult)).Append('\n');
            return sb.ToString().TrimEnd('\n');
        }

        private string BoolJoin(List<FNode> a, string op)
        {
            if (a.Count == 0) return Raw(op.ToUpperInvariant(), a);
            var sb = new StringBuilder();
            for (int i = 0; i < a.Count; i++)
            {
                if (i > 0) sb.Append(" ").Append(op).Append(" ");
                sb.Append(BoolOperand(a[i], op));
            }
            return sb.ToString();
        }

        private string BoolOperand(FNode n, string parentOp)
        {
            // wrap a lower-precedence boolean child (e.g. an OR inside an AND chain)
            if (n is FBinary b && (b.Op == "and" || b.Op == "or") && Prec(b.Op) < Prec(parentOp))
                return "(" + Write(n) + ")";
            return Write(n);
        }

        private string Aggregate(List<FNode> a)
        {
            if (a.Count == 3 && IsIntLiteral(a[0], out int fn) && IsIntLiteral(a[1], out int opt) && opt == 6)
            {
                string? verb = fn switch
                {
                    9 => "sum",
                    1 => "avg",
                    4 => "max",
                    5 => "min",
                    2 => "countNum",
                    3 => "count",
                    _ => null
                };
                if (verb != null) return verb + ".ignoreErrors(" + W(a[2]) + ")";
            }
            return Raw("AGGREGATE", a);
        }

        private string Where(string verb, List<FNode> a, bool hasValueRange)
        {
            int start = hasValueRange ? 1 : 0;
            if (a.Count <= start || (a.Count - start) % 2 != 0) return Raw(verb.ToUpperInvariant(), a);

            var conds = new List<string>();
            for (int i = start; i + 1 < a.Count; i += 2)
                conds.Add(Criteria(a[i], a[i + 1]));

            string condText = string.Join(" and ", conds);
            return hasValueRange
                ? verb + "(" + W(a[0]) + ", " + condText + ")"
                : verb + "(" + condText + ")";
        }

        private string Criteria(FNode range, FNode crit)
        {
            string r = Write(range);

            if (crit is FStr s)
            {
                var m = Regex.Match(s.Value, @"^\s*(>=|<=|<>|>|<|=)?\s*(.*)$", RegexOptions.Singleline);
                string op = m.Groups[1].Success && m.Groups[1].Value.Length > 0 ? m.Groups[1].Value : "=";
                string rest = m.Groups[2].Value;
                string val = Regex.IsMatch(rest, @"^-?\d+(\.\d+)?$") ? rest : Quote(rest);
                return r + " " + op + " " + val;
            }
            if (crit is FNum n) return r + " = " + n.Raw;
            if (crit is FBool b) return r + " = " + (b.Value ? "true" : "false");
            if (crit is FBinary bin && bin.Op == "&" && bin.Left is FStr opStr &&
                Regex.IsMatch(opStr.Value, @"^(>=|<=|<>|>|<|=)$"))
                return r + " " + opStr.Value + " " + Write(bin.Right);

            return r + " = " + Write(crit);
        }

        private string DateDiff(List<FNode> a)
        {
            if (a.Count == 3 && a[2] is FStr unit)
            {
                string u = unit.Value.ToLowerInvariant();
                if (u == "d") return "dateDiff(" + W(a[0]) + ", " + W(a[1]) + ")";
                if (u == "m") return "dateDiff.months(" + W(a[0]) + ", " + W(a[1]) + ")";
                if (u == "y") return "dateDiff.years(" + W(a[0]) + ", " + W(a[1]) + ")";
            }
            return Raw("DATEDIF", a);
        }

        private string Round(string verb, List<FNode> a)
        {
            if (a.Count == 2) return verb + "(" + W(a[0]) + ", " + W(a[1]) + ")";
            return Raw("ROUND" + (verb.EndsWith("up") ? "UP" : verb.EndsWith("down") ? "DOWN" : ""), a);
        }

        private string Sort(List<FNode> a)
        {
            if (a.Count == 1) return "sort(" + W(a[0]) + ")";
            if (a.Count >= 2)
            {
                string by = " by(" + W(a[1]) + ")";
                bool desc = a.Count >= 3 && IsIntLiteral(a[2], out int dir) && dir == -1;
                return "sort(" + W(a[0]) + ")" + by + (desc ? " descending" : "");
            }
            return Raw("SORT", a);
        }

        private string DateLit(List<FNode> a)
        {
            if (a.Count == 3 && a[0] is FNum y && a[1] is FNum m && a[2] is FNum d &&
                int.TryParse(y.Raw, out int yi) && int.TryParse(m.Raw, out int mi) && int.TryParse(d.Raw, out int di))
                return "#" + yi.ToString("D4", CultureInfo.InvariantCulture) + "-" +
                       mi.ToString("D2", CultureInfo.InvariantCulture) + "-" +
                       di.ToString("D2", CultureInfo.InvariantCulture) + "#";
            return Raw("DATE", a);
        }

        private string Legacy(string fn, List<FNode> a)
        {
            // drop a trailing exact-match flag so it round-trips through the emitter
            var args = new List<FNode>(a);
            if (args.Count == 4 && (args[3] is FBool fb && !fb.Value || IsIntLiteral(args[3], out int zero) && zero == 0))
                args.RemoveAt(3);
            return "legacy." + fn + "(" + JoinArgs(args) + ")";
        }

        private string Raw(string name, List<FNode> a)
        {
            string n = name.ToUpperInvariant();
            if (a.Count == 0) return "raw(" + Quote(n) + ")";
            return "raw(" + Quote(n) + ", " + JoinArgs(a) + ")";
        }

        private string WriteArray(FArray a)
        {
            // PExL has no array literal; surface it via raw() so it at least round-trips visibly.
            var flat = new List<FNode>();
            foreach (var row in a.Rows) flat.AddRange(row);
            return "raw(\"ARRAY\", " + JoinArgs(flat) + ")";
        }

        // ---------------- helpers ----------------

        private static bool IsIntLiteral(FNode n, out int value)
        {
            value = 0;
            if (n is FNum num && int.TryParse(num.Raw, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value))
                return true;
            if (n is FUnary u && u.Op == "-" && u.Operand is FNum un && int.TryParse(un.Raw, out int v))
            {
                value = -v;
                return true;
            }
            return false;
        }

        private static string Quote(string value) => "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
