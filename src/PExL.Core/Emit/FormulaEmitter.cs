using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using PExL.Core.Diagnostics;
using PExL.Core.Parsing.Ast;

namespace PExL.Core.Emit
{
    /// <summary>
    /// Walks the AST and produces native Excel formula fragments. Binds (::) are
    /// inlined: a name resolves back to its expression and is re-emitted.
    /// </summary>
    public sealed class FormulaEmitter
    {
        private readonly Dictionary<string, Expr> _binds = new Dictionary<string, Expr>(StringComparer.OrdinalIgnoreCase);

        public void Bind(string name, Expr expr) => _binds[name] = expr;

        public EmitResult Emit(Expr e)
        {
            switch (e)
            {
                case NumberLit n: return new EmitResult(n.Raw);
                case StringLit s: return new EmitResult(QuoteString(s.Value));
                case BoolLit b: return new EmitResult(b.Value ? "TRUE" : "FALSE");
                case EmptyLit _: return new EmitResult("\"\"");
                case DateLit d: return new EmitResult(EmitDateFromText(d.Raw));
                case ReferenceExpr r: return new EmitResult(r.Text);
                case NameRef nr: return EmitName(nr);
                case Unary u: return EmitUnary(u);
                case Binary bin: return EmitBinary(bin);
                case IfExpr iff: return EmitIf(iff);
                case VerbCall vc: return EmitVerb(vc);
                default:
                    throw new PExLException("Cannot emit unknown expression node");
            }
        }

        private EmitResult EmitName(NameRef nr)
        {
            if (!_binds.TryGetValue(nr.Name, out var expr))
                throw new PExLException($"Unknown name '{nr.Name}'. Did you bind it with '::' first?", nr.Line, nr.Column);
            return Emit(expr);
        }

        private EmitResult EmitUnary(Unary u)
        {
            var inner = Emit(u.Operand).Formula;
            if (u.Op == "not") return new EmitResult($"NOT({inner})");
            return new EmitResult($"-{Paren(u.Operand, inner, 100)}");
        }

        private static int Prec(string op)
        {
            switch (op)
            {
                case "^": return 4;
                case "*": case "/": return 3;
                case "+": case "-": return 2;
                default: return 1; // comparisons
            }
        }

        private string Paren(Expr child, string childFormula, int parentPrec)
        {
            if (child is Binary cb && cb.Op != "and" && cb.Op != "or" && Prec(cb.Op) < parentPrec)
                return $"({childFormula})";
            return childFormula;
        }

        private EmitResult EmitBinary(Binary b)
        {
            if (b.Op == "and" || b.Op == "or")
            {
                var fn = b.Op == "and" ? "AND" : "OR";
                return new EmitResult($"{fn}({Emit(b.Left).Formula},{Emit(b.Right).Formula})");
            }
            int p = Prec(b.Op);
            var l = Paren(b.Left, Emit(b.Left).Formula, p);
            var r = Paren(b.Right, Emit(b.Right).Formula, p + 1); // right side needs tighter binding for - and /
            return new EmitResult($"{l}{b.Op}{r}");
        }

        private EmitResult EmitIf(IfExpr iff)
        {
            var cond = Emit(iff.Condition).Formula;
            var then = Emit(iff.Then).Formula;
            if (iff.Else == null) return new EmitResult($"IF({cond},{then})");
            return new EmitResult($"IF({cond},{then},{Emit(iff.Else).Formula})");
        }

        // ---------- verbs ----------

        private EmitResult EmitVerb(VerbCall c)
        {
            string v = c.Verb;
            string? mod = c.Modifier;

            switch (v)
            {
                case "split": return EmitSplit(c);
                case "fromLeft": return EmitFromSide(c, before: true);
                case "fromRight": return EmitFromSide(c, before: false);
                case "at": return EmitAt(c);
                case "spill": return EmitSpill(c);
                case "combine": return EmitCombine(c);
                case "trim": return Fn("TRIM", c);
                case "clean": return Fn("CLEAN", c);
                case "upper": return Fn("UPPER", c);
                case "lower": return Fn("LOWER", c);
                case "proper": return Fn("PROPER", c);
                case "replace": return EmitReplace(c);
                case "contains": return EmitContains(c);
                case "startsWith": return EmitEdge(c, start: true);
                case "endsWith": return EmitEdge(c, start: false);
                case "length": return Fn("LEN", c);

                case "find": return EmitFind(c);
                case "position": return EmitPosition(c);

                case "ifError": return Fn("IFERROR", c);
                case "ifs": return EmitIfs(c);

                case "sum": return EmitAggregate(c, "SUM", 9);
                case "avg": return EmitAggregate(c, "AVERAGE", 1);
                case "min": return EmitAggregate(c, "MIN", 5);
                case "max": return EmitAggregate(c, "MAX", 4);
                case "count": return EmitAggregate(c, "COUNTA", 3);
                case "countNum": return EmitAggregate(c, "COUNT", 2);
                case "sumWhere": return EmitWhere(c, "SUMIFS", hasValueRange: true);
                case "countWhere": return EmitWhere(c, "COUNTIFS", hasValueRange: false);
                case "avgWhere": return EmitWhere(c, "AVERAGEIFS", hasValueRange: true);

                case "today": return new EmitResult("TODAY()");
                case "now": return new EmitResult("NOW()");
                case "addDays": return new EmitResult($"{Arg(c, 0)}+{Arg(c, 1)}");
                case "addMonths": return new EmitResult($"EDATE({Arg(c, 0)},{Arg(c, 1)})");
                case "addYears": return EmitAddYears(c);
                case "yearOf": return Fn("YEAR", c);
                case "monthOf": return Fn("MONTH", c);
                case "dayOf": return Fn("DAY", c);
                case "weekdayOf": return Fn("WEEKDAY", c);
                case "dateDiff": return EmitDateDiff(c);

                case "round": return EmitRound(c);
                case "abs": return Fn("ABS", c);
                case "sqrt": return Fn("SQRT", c);
                case "power": return Fn("POWER", c);
                case "mod": return Fn("MOD", c);

                case "filter": return EmitFilter(c);
                case "sort": return EmitSort(c);
                case "unique": return Fn("UNIQUE", c);
                case "take": return Fn("TAKE", c);

                case "col": return EmitCol(c);
                case "row": return EmitRow(c);
                case "cell": return EmitCell(c);
                case "fixed": return EmitFixed(c);
                case "Date": return EmitDateVerb(c);

                case "raw": return EmitRaw(c);
                case "legacy": return EmitLegacy(c);

                default:
                    throw new PExLException($"Unknown verb '{c.Verb}'. Use raw(\"{c.Verb.ToUpperInvariant()}\", ...) to call an Excel function directly.", c.Line, c.Column);
            }
        }

        private EmitResult Fn(string excelName, VerbCall c)
        {
            var args = c.Positional.Select(a => Emit(a).Formula);
            return new EmitResult($"{excelName}({string.Join(",", args)})");
        }

        private string Arg(VerbCall c, int i)
        {
            if (i >= c.Positional.Count)
                throw new PExLException($"'{c.Verb}' is missing argument #{i + 1}", c.Line, c.Column);
            return Emit(c.Positional[i]).Formula;
        }

        // ---- text ----

        private EmitResult EmitSplit(VerbCall c)
        {
            var source = Arg(c, 0);
            var delim = Arg(c, 1);
            var strategy = SplitStrategy.All;
            if (string.Equals(c.Modifier, "First", StringComparison.OrdinalIgnoreCase)) strategy = SplitStrategy.First;
            else if (string.Equals(c.Modifier, "Last", StringComparison.OrdinalIgnoreCase)) strategy = SplitStrategy.Last;
            return new EmitResult(new SplitInfo(source, delim, strategy));
        }

        private EmitResult EmitFromSide(VerbCall c, bool before)
        {
            var input = Emit(c.Positional[0]);
            if (input.Split != null)
            {
                var s = input.Split;
                string fn = before ? "TEXTBEFORE" : "TEXTAFTER";
                bool last = s.Strategy == SplitStrategy.Last || (s.Strategy == SplitStrategy.All && !before);
                return new EmitResult(last
                    ? $"{fn}({s.Source},{s.Delimiter},-1)"
                    : $"{fn}({s.Source},{s.Delimiter})");
            }
            // plain text with an explicit delimiter argument
            var text = input.Formula;
            var delim = c.Positional.Count > 1 ? Emit(c.Positional[1]).Formula : "\" \"";
            string fn2 = before ? "TEXTBEFORE" : "TEXTAFTER";
            bool useLast = string.Equals(c.Modifier, "last", StringComparison.OrdinalIgnoreCase);
            return new EmitResult(useLast ? $"{fn2}({text},{delim},-1)" : $"{fn2}({text},{delim})");
        }

        private EmitResult EmitAt(VerbCall c)
        {
            var input = Emit(c.Positional[0]);
            var idx = Emit(c.Positional[1]).Formula;
            string arr = input.Split != null ? input.Split.AsArrayFormula() : input.Formula;
            return new EmitResult($"INDEX({arr},{idx})");
        }

        private EmitResult EmitSpill(VerbCall c)
        {
            var input = Emit(c.Positional[0]);
            return new EmitResult(input.Split != null ? input.Split.AsArrayFormula() : input.Formula);
        }

        private EmitResult EmitCombine(VerbCall c)
        {
            var sepExpr = c.FindNamed("with");
            var sep = sepExpr != null ? Emit(sepExpr).Formula : "\"\"";
            var args = c.Positional.Select(a => Emit(a).Formula);
            return new EmitResult($"TEXTJOIN({sep},TRUE,{string.Join(",", args)})");
        }

        private EmitResult EmitReplace(VerbCall c)
        {
            if (string.Equals(c.Modifier, "nth", StringComparison.OrdinalIgnoreCase))
                return new EmitResult($"SUBSTITUTE({Arg(c, 1)},{Arg(c, 2)},{Arg(c, 3)},{Arg(c, 0)})");
            if (string.Equals(c.Modifier, "first", StringComparison.OrdinalIgnoreCase))
                return new EmitResult($"SUBSTITUTE({Arg(c, 0)},{Arg(c, 1)},{Arg(c, 2)},1)");
            if (string.Equals(c.Modifier, "last", StringComparison.OrdinalIgnoreCase))
            {
                var t = Arg(c, 0); var f = Arg(c, 1); var w = Arg(c, 2);
                return new EmitResult($"SUBSTITUTE({t},{f},{w},(LEN({t})-LEN(SUBSTITUTE({t},{f},\"\")))/LEN({f}))");
            }
            return new EmitResult($"SUBSTITUTE({Arg(c, 0)},{Arg(c, 1)},{Arg(c, 2)})");
        }

        private EmitResult EmitContains(VerbCall c)
        {
            var text = Arg(c, 0);
            var sub = Arg(c, 1);
            string seek = string.Equals(c.Modifier, "caseSensitive", StringComparison.OrdinalIgnoreCase) ? "FIND" : "SEARCH";
            return new EmitResult($"ISNUMBER({seek}({sub},{text}))");
        }

        private EmitResult EmitEdge(VerbCall c, bool start)
        {
            var text = Arg(c, 0);
            var probe = c.Positional[1];
            string lenExpr = probe is StringLit sl
                ? sl.Value.Length.ToString(CultureInfo.InvariantCulture)
                : $"LEN({Emit(probe).Formula})";
            var probeF = Emit(probe).Formula;
            string fn = start ? "LEFT" : "RIGHT";
            return new EmitResult($"{fn}({text},{lenExpr})={probeF}");
        }

        // ---- lookup ----

        private EmitResult EmitFind(VerbCall c)
        {
            var lookup = Arg(c, 0);
            string la, ra;

            var table = c.FindNamed("inTable");
            if (table != null)
            {
                var t = Emit(table).Formula;
                var colExpr = c.FindNamed("returnColumn");
                var col = colExpr != null ? Emit(colExpr).Formula : "2";
                la = $"INDEX({t},0,1)";
                ra = $"INDEX({t},0,{col})";
            }
            else
            {
                la = (c.FindNamed("within") is Expr w) ? Emit(w).Formula : Arg(c, 1);
                ra = (c.FindNamed("thenReturn") is Expr rr) ? Emit(rr).Formula : Arg(c, 2);
            }

            var ifMissing = c.FindNamed("ifMissing");
            string? matchMode = null, searchMode = null;
            if (string.Equals(c.Modifier, "wildcard", StringComparison.OrdinalIgnoreCase)) matchMode = "2";
            else if (string.Equals(c.Modifier, "approx", StringComparison.OrdinalIgnoreCase)) matchMode = "-1";
            else if (string.Equals(c.Modifier, "reverse", StringComparison.OrdinalIgnoreCase)) searchMode = "-1";

            var parts = new List<string> { lookup, la, ra };
            if (searchMode != null)
            {
                parts.Add(ifMissing != null ? Emit(ifMissing).Formula : "\"\"");
                parts.Add(matchMode ?? "0");
                parts.Add(searchMode);
            }
            else if (matchMode != null)
            {
                parts.Add(ifMissing != null ? Emit(ifMissing).Formula : "\"\"");
                parts.Add(matchMode);
            }
            else if (ifMissing != null)
            {
                parts.Add(Emit(ifMissing).Formula);
            }
            return new EmitResult($"XLOOKUP({string.Join(",", parts)})");
        }

        private EmitResult EmitPosition(VerbCall c)
        {
            var value = Arg(c, 0);
            var range = (c.FindNamed("within") is Expr w) ? Emit(w).Formula : Arg(c, 1);
            return new EmitResult($"MATCH({value},{range},0)");
        }

        private EmitResult EmitIfs(VerbCall c)
        {
            var parts = new List<string>();
            for (int i = 0; i + 1 < c.Positional.Count; i += 2)
            {
                parts.Add(Emit(c.Positional[i]).Formula);
                parts.Add(Emit(c.Positional[i + 1]).Formula);
            }
            var elseExpr = c.FindNamed("else");
            if (elseExpr != null)
            {
                parts.Add("TRUE");
                parts.Add(Emit(elseExpr).Formula);
            }
            return new EmitResult($"IFS({string.Join(",", parts)})");
        }

        // ---- aggregation ----

        private EmitResult EmitAggregate(VerbCall c, string excelName, int aggregateFn)
        {
            if (string.Equals(c.Modifier, "ignoreErrors", StringComparison.OrdinalIgnoreCase))
                return new EmitResult($"AGGREGATE({aggregateFn},6,{Arg(c, 0)})");
            return Fn(excelName, c);
        }

        private EmitResult EmitWhere(VerbCall c, string excelName, bool hasValueRange)
        {
            var parts = new List<string>();
            int condIndex = 0;
            if (hasValueRange)
            {
                parts.Add(Arg(c, 0));
                condIndex = 1;
            }
            var cond = c.FindNamed("where") ?? (condIndex < c.Positional.Count ? c.Positional[condIndex] : null);
            if (cond == null)
                throw new PExLException($"'{c.Verb}' needs a condition", c.Line, c.Column);

            foreach (var cmp in FlattenAnd(cond))
            {
                if (cmp is Binary b && b.Op != "and" && b.Op != "or")
                {
                    parts.Add(Emit(b.Left).Formula);
                    parts.Add(BuildCriteria(b.Op, b.Right));
                }
                else
                {
                    throw new PExLException($"'{c.Verb}' conditions must be comparisons joined by 'and'", c.Line, c.Column);
                }
            }
            return new EmitResult($"{excelName}({string.Join(",", parts)})");
        }

        private IEnumerable<Expr> FlattenAnd(Expr e)
        {
            if (e is Binary b && b.Op == "and")
            {
                foreach (var x in FlattenAnd(b.Left)) yield return x;
                foreach (var x in FlattenAnd(b.Right)) yield return x;
            }
            else yield return e;
        }

        private string BuildCriteria(string op, Expr value)
        {
            bool isLiteral = value is StringLit || value is NumberLit || value is BoolLit;
            if (isLiteral)
            {
                string text = value is StringLit s ? s.Value
                    : value is NumberLit n ? n.Raw
                    : ((BoolLit)value).Value ? "TRUE" : "FALSE";
                if (op == "=")
                    return value is NumberLit ? text : QuoteString(text);
                return QuoteString(op + text);
            }
            // dynamic value -> concatenate operator with the cell/expression
            var f = Emit(value).Formula;
            return op == "=" ? f : $"{QuoteString(op)}&{f}";
        }

        // ---- dates / math ----

        private EmitResult EmitAddYears(VerbCall c)
        {
            var date = Arg(c, 0);
            if (c.Positional.Count > 1 && c.Positional[1] is NumberLit n && int.TryParse(n.Raw, out var yrs))
                return new EmitResult($"EDATE({date},{yrs * 12})");
            return new EmitResult($"EDATE({date},({Arg(c, 1)})*12)");
        }

        private EmitResult EmitDateDiff(VerbCall c)
        {
            string unit = "\"d\"";
            if (string.Equals(c.Modifier, "months", StringComparison.OrdinalIgnoreCase)) unit = "\"m\"";
            else if (string.Equals(c.Modifier, "years", StringComparison.OrdinalIgnoreCase)) unit = "\"y\"";
            return new EmitResult($"DATEDIF({Arg(c, 0)},{Arg(c, 1)},{unit})");
        }

        private EmitResult EmitRound(VerbCall c)
        {
            string fn = "ROUND";
            if (string.Equals(c.Modifier, "up", StringComparison.OrdinalIgnoreCase)) fn = "ROUNDUP";
            else if (string.Equals(c.Modifier, "down", StringComparison.OrdinalIgnoreCase)) fn = "ROUNDDOWN";
            return new EmitResult($"{fn}({Arg(c, 0)},{Arg(c, 1)})");
        }

        // ---- filter / shape ----

        private EmitResult EmitFilter(VerbCall c)
        {
            var range = Arg(c, 0);
            var cond = c.FindNamed("where") ?? (c.Positional.Count > 1 ? c.Positional[1] : null);
            if (cond == null) throw new PExLException("'filter' needs a where condition", c.Line, c.Column);
            return new EmitResult($"FILTER({range},{Emit(cond).Formula})");
        }

        private EmitResult EmitSort(VerbCall c)
        {
            var range = Arg(c, 0);
            var byExpr = c.FindNamed("by");
            bool desc = c.FindNamed("descending") != null;
            if (byExpr == null && c.Positional.Count > 1) byExpr = c.Positional[1];
            if (byExpr == null) return new EmitResult($"SORT({range})");
            var by = Emit(byExpr).Formula;
            return new EmitResult(desc ? $"SORT({range},{by},-1)" : $"SORT({range},{by})");
        }

        // ---- references / literals ----

        private EmitResult EmitCol(VerbCall c)
        {
            if (c.Positional[0] is StringLit s) return new EmitResult($"{s.Value}:{s.Value}");
            var f = Emit(c.Positional[0]).Formula;
            return new EmitResult($"{f}:{f}");
        }

        private EmitResult EmitRow(VerbCall c)
        {
            var f = Arg(c, 0);
            return new EmitResult($"{f}:{f}");
        }

        private EmitResult EmitCell(VerbCall c)
        {
            string colPart = c.Positional[0] is StringLit s ? s.Value : Emit(c.Positional[0]).Formula;
            string rowPart = c.Positional[1] is NumberLit n ? n.Raw : Emit(c.Positional[1]).Formula;
            return new EmitResult($"{colPart}{rowPart}");
        }

        private EmitResult EmitFixed(VerbCall c)
        {
            if (c.Positional[0] is ReferenceExpr r)
                return new EmitResult(Absolutize(r.Text));
            return new EmitResult(Emit(c.Positional[0]).Formula);
        }

        private EmitResult EmitDateVerb(VerbCall c)
        {
            if (c.Positional.Count > 0 && c.Positional[0] is StringLit s)
                return new EmitResult(EmitDateFromText(s.Value));
            throw new PExLException("Date(...) expects a date string like Date(\"2024-01-01\")", c.Line, c.Column);
        }

        private EmitResult EmitRaw(VerbCall c)
        {
            if (c.Positional.Count == 0 || !(c.Positional[0] is StringLit fn))
                throw new PExLException("raw(...) expects a function name first, e.g. raw(\"SUMPRODUCT\", ...)", c.Line, c.Column);
            var args = c.Positional.Skip(1).Select(a => Emit(a).Formula);
            return new EmitResult($"{fn.Value.ToUpperInvariant()}({string.Join(",", args)})");
        }

        private EmitResult EmitLegacy(VerbCall c)
        {
            if (string.IsNullOrEmpty(c.Modifier))
                throw new PExLException("legacy.* expects a function, e.g. legacy.vlookup(...)", c.Line, c.Column);
            var name = c.Modifier!.ToUpperInvariant();
            var args = c.Positional.Select(a => Emit(a).Formula).ToList();
            if (name == "VLOOKUP" && args.Count == 3) args.Add("FALSE");
            if (name == "HLOOKUP" && args.Count == 3) args.Add("FALSE");
            return new EmitResult($"{name}({string.Join(",", args)})");
        }

        // ---------- helpers ----------

        private static string QuoteString(string value)
            => "\"" + value.Replace("\"", "\"\"") + "\"";

        private static string EmitDateFromText(string raw)
        {
            var parts = raw.Trim().Split(new[] { '-', '/', '.' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 3 &&
                int.TryParse(parts[0], out var y) &&
                int.TryParse(parts[1], out var m) &&
                int.TryParse(parts[2], out var d))
            {
                return $"DATE({y},{m},{d})";
            }
            throw new PExLException($"Cannot parse date '{raw}'. Use ISO form like #2024-01-31#.");
        }

        private static string Absolutize(string a1)
        {
            // Sheet!A1:B2 -> Sheet!$A$1:$B$2 ; leaves any sheet qualifier untouched.
            int bang = a1.LastIndexOf('!');
            string sheet = bang >= 0 ? a1.Substring(0, bang + 1) : string.Empty;
            string body = bang >= 0 ? a1.Substring(bang + 1) : a1;
            var sb = new StringBuilder(sheet);
            foreach (var part in body.Split(':'))
            {
                if (sb.Length > sheet.Length) sb.Append(':');
                sb.Append(AbsolutizeCell(part));
            }
            return sb.ToString();
        }

        private static string AbsolutizeCell(string cell)
        {
            var sb = new StringBuilder();
            int i = 0;
            if (i < cell.Length && cell[i] == '$') i++;
            sb.Append('$');
            while (i < cell.Length && char.IsLetter(cell[i])) { sb.Append(cell[i]); i++; }
            if (i < cell.Length && cell[i] == '$') i++;
            sb.Append('$');
            while (i < cell.Length && char.IsDigit(cell[i])) { sb.Append(cell[i]); i++; }
            return sb.ToString();
        }
    }
}
