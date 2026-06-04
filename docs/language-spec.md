# PExL Language Specification (v1)

PExL ("Plain Excel Language", styled _PExL_) is an English-like domain-specific
language that transpiles to native Microsoft Excel formulas. You write readable,
pipe-driven logic; PExL emits a standard `=FORMULA(...)` that works in any Excel,
on any machine, with or without the add-in installed.

This document is the authoritative specification for the v1 surface. It is the
contract the lexer, parser, and emitter implement.

---

## 1. Design principles

1. **Reads like thought.** `take this, then do that, then put it there`.
2. **One formula out.** Every statement compiles to a single, native Excel formula
   string. No macros, no runtime dependency, no performance hit.
3. **Structured but forgiving.** A fixed, predictable skeleton with a relaxed
   vocabulary: synonyms, optional filler words, and flexible phrasing are accepted,
   but ambiguity is never silently guessed - it is surfaced for confirmation.
4. **Modern by default, complete by escape hatch.** The ~35 core verbs map to the
   best modern functions (`XLOOKUP`, `TEXTSPLIT`, `FILTER`, ...). The other ~460
   Excel functions remain reachable through `raw(...)` and the `legacy.*` namespace.

---

## 2. Statement structure

Every PExL statement has the shape:

```
EXPRESSION  [:: name]  [-> target]
```

| Piece        | Operator | Meaning                                                       |
| ------------ | -------- | ------------------------------------------------------------- |
| `EXPRESSION` |          | Produces a value (cell, range, literal, or a chain of verbs). |
| bind         | `::`     | Names the value so later statements can reuse it.             |
| output       | `->`     | Writes the value to a destination cell or range.              |

- Statements are separated by **newlines**. Commas may also separate statements
  inline for one-liners.
- `//` starts a line comment. Comments are ignored by the transpiler but preserved
  in the editor.

```pexl
// split an order code and fan the parts into two cells
B2 |> split.First("-") :: parts
parts |> fromLeft  -> C2
parts |> fromRight -> D2
```

### 2.1 The three operators

| Operator | Name   | Behavior                                                                       |
| -------- | ------ | ------------------------------------------------------------------------------ |
| `\|>`    | pipe   | Feeds the value on the left in as the **first argument** of the verb on the right. |
| `::`     | bind   | Captures the current value under a name. The name is a **compile-time alias**.  |
| `->`     | output | Writes to a cell/range. **Output-only** - it never means anything else.        |

`a |> f(b)` is exactly equivalent to `f(a, b)`.

---

## 3. Lexical elements

### 3.1 References (Excel-native)

A1 notation passes through verbatim, including absolutes and sheet qualifiers:

```
A1            single cell
A1:B15        range
A:A   1:1     whole column / whole row
$A$1  $A1  A$1   absolute / mixed
Sheet2!A1     other sheet
'My Sheet'!A1:B5   quoted sheet name (spaces)
```

Friendly helpers (resolve to A1 at compile time when arguments are literal):

```
col("A")            -> A:A
row(2)              -> 2:2
cell("C", 2)        -> C2          (column letter, row number)
fixed(A1)           -> $A$1        (wrap any ref to make it absolute)
fixed(A1:B2)        -> $A$1:$B$2
```

### 3.2 Strings

```
"normal text"                  standard string
`raw text with "quotes"`       backtick raw string - inner double-quotes are literal
```

Use backticks whenever your text contains double quotes - no escaping required.
The emitter handles Excel-side quote doubling automatically.

### 3.3 Numbers and booleans

```
100      3.14      -5      1e6      10%
true     false
```

### 3.4 Dates

```
#2024-01-01#                   -> =DATE(2024,1,1)
Date("2024-01-01")             -> =DATE(2024,1,1)        (ISO assumed)
Date("01-02-2024", en-US)      -> =DATE(2024,1,2)        (locale parses "Jan 2")
Date("01-02-2024", pl-PL)      -> =DATE(2024,2,1)        (locale parses "1 Feb")
```

The locale tag only affects how PExL **reads** the string at compile time; the
emitted `DATE(y,m,d)` is always unambiguous.

### 3.5 Blank

```
empty        -> ""   (an empty string in the emitted formula)
```

---

## 4. The forgiving layer

PExL has two token classes with different rules.

### 4.1 Structural keywords - case-sensitive, lowercase

These form the language skeleton and must be written lowercase:

```
control flow:           if   then   else   return
literal keywords:       true   false   empty
prepositional labels:   within   by   with   thenReturn   from   of   where
```

### 4.2 Verbs - case-insensitive, with synonyms

Verb names match case-insensitively (`Split` = `split` = `SPLIT`), and each verb
carries a small **curated** synonym set:

```
find    = lookup  = search
combine = join    = concat = merge
avg     = average = mean
```

### 4.3 Filler words

These are ignored anywhere they appear, so natural phrasing parses:

```
the   of   value   in   then   please   a
```

### 4.4 Ambiguity policy

When a parse is ambiguous or low-confidence, PExL does **not** guess silently.
The add-in shows a **confirm-preview**: the normalized PExL plus the exact formula
it will produce, and requires an explicit click before writing to the cell.

---

## 5. Operator precedence

From highest to lowest; parentheses always override:

```
^                          power
* /                        multiply / divide
+ -                        add / subtract
= <> != > < >= <=          comparison   (is / is not / isNot are aliases)
and  or                    boolean       (&& / || / ! are aliases)
```

---

## 6. Core verb reference

Notation: each entry shows the canonical PExL form and the emitted formula. Synonyms
and the relaxed phrasings still parse to the same result.

### 6.1 Text

#### split

`split` chooses the split strategy; extractors pull pieces out of it.

| PExL                          | Emits                          | On `"$100-USD-x"` |
| ----------------------------- | ------------------------------ | ----------------- |
| `B2 \|> split.First("-")`     | binary split at **first** `-`  | `["$100","USD-x"]`|
| `B2 \|> split.Last("-")`      | binary split at **last** `-`   | `["$100-USD","x"]`|
| `B2 \|> split("-")`           | split on **all** `-` (array)   | `["$100","USD","x"]` |

Extractors (work on whatever `split` produced):

| PExL                                | Emits                          | Result   |
| ----------------------------------- | ------------------------------ | -------- |
| `B2 \|> split.First("-") \|> fromLeft`  | `=TEXTBEFORE(B2,"-")`      | `"$100"` |
| `B2 \|> split.First("-") \|> fromRight` | `=TEXTAFTER(B2,"-")`       | `"USD-x"`|
| `B2 \|> split.Last("-")  \|> fromLeft`  | `=TEXTBEFORE(B2,"-",-1)`   | `"$100-USD"` |
| `B2 \|> split.Last("-")  \|> fromRight` | `=TEXTAFTER(B2,"-",-1)`    | `"x"`    |
| `B2 \|> split("-") \|> at(2)`           | `=INDEX(TEXTSPLIT(B2,"-"),2)` | `"USD"` |
| `B2 \|> split("-") \|> spill`           | `=TEXTSPLIT(B2,"-")`      | spills   |

Defaults with no `.First`/`.Last`: `fromLeft` = first piece, `fromRight` = last piece.

#### Other text verbs

| PExL                         | Emits                                    |
| ---------------------------- | ---------------------------------------- |
| `combine(A1, B1, C1) with("-")` | `=TEXTJOIN("-",TRUE,A1,B1,C1)` (skips blanks) |
| `trim(B2)`                   | `=TRIM(B2)`                              |
| `clean(B2)`                  | `=CLEAN(B2)`                             |
| `upper(B2)` / `lower(B2)` / `proper(B2)` | `=UPPER/LOWER/PROPER(B2)`    |
| `replace(B2,"-","_")`        | `=SUBSTITUTE(B2,"-","_")` (all)          |
| `replace.first(B2,"-","_")`  | `=SUBSTITUTE(B2,"-","_",1)`              |
| `replace.last(B2,"-","_")`   | `=SUBSTITUTE(B2,"-","_",<count>)`        |
| `replace.nth(2, B2,"-","_")` | `=SUBSTITUTE(B2,"-","_",2)`              |
| `contains(B2,"x")` / `B2 contains "x"` | `=ISNUMBER(SEARCH("x",B2))` (case-insensitive) |
| `startsWith(B2,"USD")`       | `=LEFT(B2,3)="USD"`                      |
| `endsWith(B2,".csv")`        | `=RIGHT(B2,4)=".csv"`                    |
| `length(B2)`                 | `=LEN(B2)`                               |

### 6.2 Lookup

```
find D2 within A1:A100 thenReturn B1:B100               -> =XLOOKUP(D2,A1:A100,B1:B100)
find D2 within A1:A100 thenReturn B1:B100 ifMissing "N/A" -> =XLOOKUP(D2,A1:A100,B1:B100,"N/A")
D2 |> find(A1:A100, B1:B100)                            (concise pipe form, equivalent)
```

Match modifiers and table form:

| PExL                                      | Emits                                |
| ----------------------------------------- | ------------------------------------ |
| `find.wildcard D2 within ...`             | `=XLOOKUP(...,,2)`                    |
| `find.approx D2 within ...`               | `=XLOOKUP(...,,-1)`                   |
| `find.reverse D2 within ...`              | `=XLOOKUP(...,,,-1)`                  |
| `find D2 inTable A1:D100 returnColumn 3`  | `=XLOOKUP(D2,A1:A100,C1:C100)`        |
| `position D2 within A1:A100`              | `=MATCH(D2,A1:A100,0)`               |

### 6.3 Logic

```
if B2 > 10 then "High" else "Low"          -> =IF(B2>10,"High","Low")
```

Multi-branch with `check` (subject set once, compared implicitly):

```
check B2:
  if > 90 then "A"
  if > 80 then "B"
  else "C"
                                           -> =IFS(B2>90,"A",B2>80,"B",TRUE,"C")
```

```
find D2 within A1:A100 thenReturn B1:B100 |> ifError("not found")
                                           -> =IFERROR(XLOOKUP(D2,A1:A100,B1:B100),"not found")

A1 > 0 and B1 > 0                          -> =AND(A1>0,B1>0)
A1 > 0 or B1 < 5                           -> =OR(A1>0,B1<5)
not contains(B2,"x")                       -> =NOT(ISNUMBER(SEARCH("x",B2)))
```

### 6.4 Aggregation

| PExL                                          | Emits                                  |
| --------------------------------------------- | -------------------------------------- |
| `sum(A1:A10)`                                 | `=SUM(A1:A10)`                         |
| `avg(A1:A10)`                                 | `=AVERAGE(A1:A10)`                     |
| `min(A1:A10)` / `max(A1:A10)`                 | `=MIN/MAX(A1:A10)`                     |
| `count(A1:A10)`                               | `=COUNTA(A1:A10)` (non-empty)          |
| `countNum(A1:A10)`                            | `=COUNT(A1:A10)` (numbers only)        |
| `sumWhere(A1:A10, B1:B10 > 100)`              | `=SUMIFS(A1:A10,B1:B10,">100")`        |
| `sumWhere(A1:A10, B1:B10 > 100 and C1:C10 = "West")` | `=SUMIFS(A1:A10,B1:B10,">100",C1:C10,"West")` |
| `countWhere(B1:B10 = "West")`                 | `=COUNTIFS(B1:B10,"West")`             |
| `avgWhere(A1:A10, B1:B10 >= 50)`              | `=AVERAGEIFS(A1:A10,B1:B10,">=50")`    |
| `sum.ignoreErrors(A1:A10)`                    | `=AGGREGATE(9,6,A1:A10)`               |

### 6.5 Dates

| PExL                  | Emits                       |
| --------------------- | --------------------------- |
| `today()` / `now()`   | `=TODAY()` / `=NOW()`       |
| `addDays(A2, 7)`      | `=A2+7`                     |
| `addMonths(A2, 3)`    | `=EDATE(A2,3)`              |
| `addYears(A2, 1)`     | `=EDATE(A2,12)`             |
| `yearOf(A2)` / `monthOf(A2)` / `dayOf(A2)` | `=YEAR/MONTH/DAY(A2)` |
| `weekdayOf(A2)`       | `=WEEKDAY(A2)`              |
| `dateDiff(A2, B2)`    | `=DATEDIF(A2,B2,"d")`       |
| `dateDiff.months(A2,B2)` / `dateDiff.years(A2,B2)` | `=DATEDIF(...,"m"/"y")` |

### 6.6 Math

| PExL              | Emits              |
| ----------------- | ------------------ |
| `round(A2, 2)`    | `=ROUND(A2,2)`     |
| `round.up(A2,2)` / `round.down(A2,2)` | `=ROUNDUP/ROUNDDOWN(A2,2)` |
| `abs(A2)`         | `=ABS(A2)`         |
| `sqrt(A2)`        | `=SQRT(A2)`        |
| `power(A2, 3)`    | `=POWER(A2,3)`     |
| `mod(A2, 3)`      | `=MOD(A2,3)`       |
| `+ - * / ^`       | native arithmetic  |

### 6.7 Filter / shape (dynamic arrays)

| PExL                                   | Emits                       |
| -------------------------------------- | --------------------------- |
| `filter(A1:C100) where(B1:B100 > 100)` | `=FILTER(A1:C100,B1:B100>100)` |
| `sort(A1:C100) by(2)`                  | `=SORT(A1:C100,2)`          |
| `sort(A1:C100) by(2) descending`       | `=SORT(A1:C100,2,-1)`       |
| `unique(A1:A100)`                      | `=UNIQUE(A1:A100)`          |
| `take(A1:A100, 5)`                     | `=TAKE(A1:A100,5)`          |

### 6.8 Escape hatches

```
raw("SUMPRODUCT", A1:A10, B1:B10)   -> =SUMPRODUCT(A1:A10,B1:B10)   (any function, verbatim)
legacy.vlookup(D2, A1:B100, 2)      -> =VLOOKUP(D2,A1:B100,2,FALSE)  (deprecated, namespaced)
```

### 6.9 Globals (document variables)

`MakeGlobal(EXPRESSION) :: NAME` promotes a value, cell or range to a **workbook-level
global** — a native Excel **Defined Name** (workbook scope). It must be named with `::`;
an unnamed `MakeGlobal(...)` is an error. Synonyms: `global`, `defineGlobal`, `setGlobal`.

| PExL                              | Effect                                                          |
| --------------------------------- | -------------------------------------------------------------- |
| `MakeGlobal(0.2) :: TaxRate`      | Defined Name `TaxRate` refers to `=0.2`                         |
| `MakeGlobal(A2:A100) :: SalesQ1`  | Defined Name `SalesQ1` refers to `=$A$2:$A$100` (auto-locked)  |
| `MakeGlobal(B1 - C1) :: Margin`   | Defined Name `Margin` refers to `=B1-C1`                        |
| `MakeGlobal(0.2) :: TaxRate -> A1`| also drops `=TaxRate` into `A1`                                 |

- A **bare cell/range** inner expression is absolutized (as if wrapped in `fixed(...)`)
  so the name doesn't shift when referenced from different cells. Constants and formulas
  are emitted as written.
- Referencing a declared global elsewhere emits the **name verbatim** (e.g.
  `sum(SalesQ1) * TaxRate` → `=SUM(SalesQ1)*TaxRate`), unlike a `::` bind which is
  inlined. A global may reference earlier globals.
- Because the output is a standard Defined Name, globals persist in the workbook and keep
  working for recipients who don't have the add-in.

**Console commands.** `ShowGlobals()` is not a formula — it is a console command that the
add-in intercepts to open the **globals manager** (view / edit refers-to / rename /
delete). It produces no cell output. The manager lists **only PExL-created globals**
(tracked in a `CustomXMLPart` registry); the user's own Excel named ranges are never
shown or modified. Synonyms: `globals`, `listGlobals`. Using `MakeGlobal`/`ShowGlobals`
inside a formula value is an error.

---

## 7. Emission rules

- **Binds (`::`) are inlined by default.** Each use of a name re-emits its
  expression, so formulas work in every Excel version. A "use LET when supported"
  setting can instead emit `=LET(name, expr, ...)` when a bind is reused inside a
  single output cell (Excel 365/2021+).
- **Globals (`MakeGlobal`) are persisted, not inlined.** They become workbook Defined
  Names and are emitted verbatim by name wherever referenced (see §6.9).
- **Invariant formulas.** PExL always emits with `,` separators and injects via
  `Range.Formula2`, letting Excel localize the display. Behavior is tested under
  non-US locales (e.g. `pl-PL`, which displays `;`).
- **Output targets.**
  - `-> C2` writes the formula at a single anchor; dynamic arrays spill naturally.
  - `-> C2:C100` fills the formula across the range, adjusting relative references
    like Excel's fill handle.

---

## 8. Two-way code preservation

When PExL writes a formula, the add-in also stores the original PExL source keyed
by `Sheet!Address` in a `CustomXMLPart` inside the workbook. Selecting that cell
later repopulates the editor with your readable PExL instead of the raw formula -
round-tripping without needing the (later-phase) formula decompiler.

---

## 9. Grammar sketch (EBNF-ish)

```ebnf
program     = { statement } ;
statement   = expression [ "::" name ] [ "->" target ]
            | "MakeGlobal" "(" expression ")" "::" name [ "->" target ]
            | "ShowGlobals" "(" ")" ;
expression  = pipeline ;
pipeline    = term { "|>" verbcall } ;
term        = literal | reference | verbcall | "(" expression ")" | unary | binary ;
verbcall    = verb [ "." modifier ] [ "(" args ")" ] { preposition arg } ;
binary      = expression op expression ;
op          = "^" | "*" | "/" | "+" | "-" | comparison | "and" | "or" ;
comparison  = "=" | "<>" | "!=" | ">" | "<" | ">=" | "<=" | "is" | "is" "not" ;
preposition = "within" | "by" | "with" | "thenReturn" | "from" | "of" | "where" | "ifMissing" ;
literal     = number | string | rawstring | date | "true" | "false" | "empty" ;
reference   = a1ref | "col" "(" string ")" | "row" "(" number ")"
            | "cell" "(" string "," number ")" | "fixed" "(" reference ")" ;
target      = a1ref ;
```

This sketch is intentionally permissive on ordering; the forgiving layer (Section 4)
resolves relaxed phrasing, and the confirm-preview gate (Section 4.4) guards
ambiguity.
