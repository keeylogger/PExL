// Full Monaco language support for PExL: tokenizer, theme, auto-closing,
// completions (verbs + keywords + snippets), hovers and signature help.
// Invoked by index.html once Monaco has loaded: window.registerPexl(monaco).
(function () {
  // ---- vocabulary (kept in sync with PExL.Core/StdLib/VerbRegistry) ----
  const KEYWORDS = [
    "if", "then", "else", "return", "true", "false", "empty",
    "within", "by", "with", "thenReturn", "from", "of", "where",
    "ifMissing", "inTable", "returnColumn", "and", "or", "not", "is",
    "check", "descending", "ascending"
  ];

  // verb -> { sig, doc } used for completion docs, hovers and signature help
  const VERBS = {
    split:       { sig: "split(delimiter)", doc: "Split text by a delimiter. Use `.First(n)` / `.Last(n)` to keep a side, or `|> spill` to spread across columns." },
    fromLeft:    { sig: "fromLeft", doc: "Take the part to the left of a split." },
    fromRight:   { sig: "fromRight", doc: "Take the part to the right of a split." },
    at:          { sig: "at(index)", doc: "Take the part at a 1-based index after a split." },
    spill:       { sig: "spill", doc: "Spread the pieces across adjacent cells (dynamic array)." },
    combine:     { sig: "combine(a, b, ...)", doc: "Join values together. Pair with `with(separator)`." },
    trim:        { sig: "trim", doc: "Remove leading/trailing/duplicate spaces (TRIM)." },
    clean:       { sig: "clean", doc: "Strip non-printable characters (CLEAN)." },
    upper:       { sig: "upper", doc: "Convert text to UPPERCASE." },
    lower:       { sig: "lower", doc: "Convert text to lowercase." },
    proper:      { sig: "proper", doc: "Capitalize Each Word (PROPER)." },
    replace:     { sig: "replace(find, with)", doc: "Replace text. `.first` / `.last` target one occurrence." },
    contains:    { sig: "contains(text)", doc: "True when the value contains the given text." },
    startsWith:  { sig: "startsWith(text)", doc: "True when the value starts with the given text." },
    endsWith:    { sig: "endsWith(text)", doc: "True when the value ends with the given text." },
    length:      { sig: "length", doc: "Number of characters (LEN)." },
    find:        { sig: "find x within range thenReturn range", doc: "Lookup. `ifMissing` sets a fallback; `.wildcard()` enables * and ?." },
    position:    { sig: "position(needle)", doc: "Position of a substring (SEARCH)." },
    ifError:     { sig: "ifError(value, fallback)", doc: "Return fallback when the value errors." },
    sum:         { sig: "sum(range)", doc: "Add numbers in a range (SUM)." },
    avg:         { sig: "avg(range)", doc: "Average of a range (AVERAGE)." },
    min:         { sig: "min(range)", doc: "Smallest value (MIN)." },
    max:         { sig: "max(range)", doc: "Largest value (MAX)." },
    count:       { sig: "count(range)", doc: "Count non-empty cells (COUNTA)." },
    countNum:    { sig: "countNum(range)", doc: "Count numeric cells (COUNT)." },
    sumWhere:    { sig: "sumWhere(sumRange, criteriaRange = value, ...)", doc: "Conditional sum (SUMIFS)." },
    countWhere:  { sig: "countWhere(criteriaRange = value, ...)", doc: "Conditional count (COUNTIFS)." },
    avgWhere:    { sig: "avgWhere(avgRange, criteriaRange = value, ...)", doc: "Conditional average (AVERAGEIFS)." },
    today:       { sig: "today()", doc: "Today's date (TODAY)." },
    now:         { sig: "now()", doc: "Current date and time (NOW)." },
    addDays:     { sig: "addDays(date, n)", doc: "Add days to a date." },
    addMonths:   { sig: "addMonths(date, n)", doc: "Add months (EDATE)." },
    addYears:    { sig: "addYears(date, n)", doc: "Add years." },
    yearOf:      { sig: "yearOf(date)", doc: "Year part (YEAR)." },
    monthOf:     { sig: "monthOf(date)", doc: "Month part (MONTH)." },
    dayOf:       { sig: "dayOf(date)", doc: "Day part (DAY)." },
    weekdayOf:   { sig: "weekdayOf(date)", doc: "Day of week (WEEKDAY)." },
    dateDiff:    { sig: "dateDiff(start, end, unit)", doc: "Difference between dates (DATEDIF)." },
    round:       { sig: "round(value, digits)", doc: "Round to digits (ROUND)." },
    abs:         { sig: "abs(value)", doc: "Absolute value (ABS)." },
    sqrt:        { sig: "sqrt(value)", doc: "Square root (SQRT)." },
    power:       { sig: "power(base, exp)", doc: "Raise to a power (POWER)." },
    mod:         { sig: "mod(n, divisor)", doc: "Remainder (MOD)." },
    filter:      { sig: "filter(range where condition)", doc: "Keep rows matching a condition (FILTER)." },
    sort:        { sig: "sort(range by key)", doc: "Sort a range (SORT). Add `descending`." },
    unique:      { sig: "unique(range)", doc: "Distinct values (UNIQUE)." },
    take:        { sig: "take(range, n)", doc: "First n rows (TAKE)." },
    col:         { sig: "col(\"A\")", doc: "Whole-column reference helper." },
    row:         { sig: "row(2)", doc: "Whole-row reference helper." },
    cell:        { sig: "cell(\"A\", 2)", doc: "Build an A1 reference from column and row." },
    fixed:       { sig: "fixed(ref)", doc: "Lock a reference with absolutes ($A$1)." },
    Date:        { sig: "Date(\"2024-01-31\", locale)", doc: "Parse a date from text." },
    raw:         { sig: "raw(\"EXCELFUNC\", args...)", doc: "Escape hatch: emit a native Excel function verbatim." },
    legacy:      { sig: "legacy.FUNC(args...)", doc: "Escape hatch: call any legacy Excel function by name." }
  };

  const VERB_NAMES = Object.keys(VERBS);

  window.registerPexl = function (monaco) {
    if (monaco.languages.getLanguages().some(l => l.id === "pexl")) return;
    monaco.languages.register({ id: "pexl" });

    // ---- syntax tokenizer ----
    monaco.languages.setMonarchTokensProvider("pexl", {
      keywords: KEYWORDS,
      verbs: VERB_NAMES,
      tokenizer: {
        root: [
          [/\/\/.*$/, "comment"],
          [/`[^`]*`/, "string"],
          [/"(?:[^"]|"")*"/, "string"],
          [/#[^#]*#/, "string.date"],
          // A1 references incl. sheet-qualified and absolute, ranges
          [/(?:[A-Za-z0-9_]+!)?\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?/, "variable.ref"],
          // whole-column / whole-row refs like A:A or 2:2
          [/(?:[A-Za-z0-9_]+!)?\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/, "variable.ref"],
          [/\b\d+(\.\d+)?%?\b/, "number"],
          [/\|>|::|->/, "operator.pexl"],
          [/>=|<=|<>|!=|[+\-*/^=<>]/, "operator"],
          [/[A-Za-z_]\w*/, {
            cases: {
              "@keywords": "keyword",
              "@verbs": "type",
              "@default": "identifier"
            }
          }],
          [/[()\[\],.]/, "delimiter"]
        ]
      }
    });

    // ---- auto-closing & surrounding pairs ----
    monaco.languages.setLanguageConfiguration("pexl", {
      comments: { lineComment: "//" },
      brackets: [["(", ")"], ["[", "]"]],
      autoClosingPairs: [
        { open: "(", close: ")" },
        { open: "[", close: "]" },
        { open: "\"", close: "\"" },
        { open: "`", close: "`" },
        { open: "#", close: "#" }
      ],
      surroundingPairs: [
        { open: "(", close: ")" },
        { open: "[", close: "]" },
        { open: "\"", close: "\"" },
        { open: "`", close: "`" }
      ]
    });

    // ---- dark theme tuned for PExL token types ----
    monaco.editor.defineTheme("pexl-dark", {
      base: "vs-dark",
      inherit: true,
      rules: [
        { token: "comment", foreground: "6a9955", fontStyle: "italic" },
        { token: "keyword", foreground: "c586c0" },
        { token: "type", foreground: "4ec9b0" },
        { token: "string", foreground: "ce9178" },
        { token: "string.date", foreground: "d7ba7d" },
        { token: "number", foreground: "b5cea8" },
        { token: "variable.ref", foreground: "569cd6", fontStyle: "bold" },
        { token: "operator.pexl", foreground: "d16969", fontStyle: "bold" },
        { token: "operator", foreground: "d4d4d4" },
        { token: "identifier", foreground: "9cdcfe" }
      ],
      colors: { "editor.background": "#1e1e1e" }
    });

    // ---- light theme (default) ----
    monaco.editor.defineTheme("pexl-light", {
      base: "vs",
      inherit: true,
      rules: [
        { token: "comment", foreground: "008000", fontStyle: "italic" },
        { token: "keyword", foreground: "8f1a9c" },
        { token: "type", foreground: "0a7d63" },
        { token: "string", foreground: "a31515" },
        { token: "string.date", foreground: "9a6a00" },
        { token: "number", foreground: "098658" },
        { token: "variable.ref", foreground: "0a5fb0", fontStyle: "bold" },
        { token: "operator.pexl", foreground: "b5320c", fontStyle: "bold" },
        { token: "operator", foreground: "1f1f1f" },
        { token: "identifier", foreground: "0a4a8c" }
      ],
      colors: { "editor.background": "#ffffff" }
    });

    // ---- completions: verbs, keywords, helpers + a few snippets ----
    monaco.languages.registerCompletionItemProvider("pexl", {
      triggerCharacters: [" ", "|", ".", "("],
      provideCompletionItems: function (model, position) {
        const word = model.getWordUntilPosition(position);
        const range = {
          startLineNumber: position.lineNumber, endLineNumber: position.lineNumber,
          startColumn: word.startColumn, endColumn: word.endColumn
        };
        const K = monaco.languages.CompletionItemKind;
        const suggestions = [];

        VERB_NAMES.forEach(function (v) {
          suggestions.push({
            label: v, kind: K.Function, insertText: v, range: range,
            detail: VERBS[v].sig,
            documentation: { value: "**" + VERBS[v].sig + "**\n\n" + VERBS[v].doc }
          });
        });
        KEYWORDS.forEach(function (k) {
          suggestions.push({ label: k, kind: K.Keyword, insertText: k, range: range });
        });

        const snip = monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet;
        [
          { label: "lookup", text: "find ${1:A2} within ${2:Sheet2!A:A} thenReturn ${3:Sheet2!B:B} ifMissing \"${4:N/A}\" -> ${5:C2}", doc: "Lookup with fallback" },
          { label: "split2cols", text: "${1:B2} |> split.First(\"${2:-}\") :: parts\nparts |> fromLeft  -> ${3:C2}\nparts |> fromRight -> ${4:D2}", doc: "Split into two columns" },
          { label: "check", text: "check\n  ${1:A2} > ${2:0} then \"${3:ok}\"\n  else \"${4:no}\"\n-> ${5:B2}", doc: "Multi-branch logic" },
          { label: "sumwhere", text: "sumWhere(${1:C2:C1000}, ${2:A2:A1000} = \"${3:West}\") -> ${4:F2}", doc: "Conditional sum" }
        ].forEach(function (s) {
          suggestions.push({
            label: s.label, kind: K.Snippet, insertText: s.text,
            insertTextRules: snip, range: range, detail: "snippet", documentation: s.doc
          });
        });

        return { suggestions: suggestions };
      }
    });

    // ---- hovers ----
    monaco.languages.registerHoverProvider("pexl", {
      provideHover: function (model, position) {
        const word = model.getWordAtPosition(position);
        if (!word) return null;
        const v = VERBS[word.word];
        if (!v) return null;
        return {
          range: new monaco.Range(position.lineNumber, word.startColumn, position.lineNumber, word.endColumn),
          contents: [{ value: "**" + v.sig + "**" }, { value: v.doc }]
        };
      }
    });

    // ---- signature help inside verb(...) ----
    monaco.languages.registerSignatureHelpProvider("pexl", {
      signatureHelpTriggerCharacters: ["(", ","],
      provideSignatureHelp: function (model, position) {
        const line = model.getValueInRange({
          startLineNumber: position.lineNumber, startColumn: 1,
          endLineNumber: position.lineNumber, endColumn: position.column
        });
        const m = line.match(/([A-Za-z_]\w*)\s*\([^()]*$/);
        if (!m) return null;
        const v = VERBS[m[1]];
        if (!v) return null;
        return {
          value: {
            signatures: [{ label: v.sig, documentation: v.doc, parameters: [] }],
            activeSignature: 0, activeParameter: 0
          },
          dispose: function () {}
        };
      }
    });
  };
})();
