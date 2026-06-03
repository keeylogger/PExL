# Contributing to PExL

Thanks for your interest in PExL — the Plain Excel Language! This guide covers how to
build, test, and contribute.

## Ground rules

- Be kind and constructive (see the [Code of Conduct](CODE_OF_CONDUCT.md)).
- Open an issue to discuss anything non-trivial before sending a large PR.
- Keep the language **forgiving but unambiguous**: when in doubt, surface a confirm-preview
  rather than silently guessing (see `docs/language-spec.md`, §4).

## Prerequisites

- **.NET SDK 9.0+** (builds the core + runs the tests).
- **.NET Framework 4.8 developer pack** (the Excel add-in targets `net48`).
- **Excel for Windows** + the **WebView2 runtime** (only needed to run the add-in, not to
  build/test the language).

## Project layout

```
src/
  PExL.Core/        the language: Lexing → Parsing → Emit (+ Decompile back to PExL).
                    Pure .NET (netstandard2.0), no Excel dependency.
  PExL.AddIn/       Excel-DNA add-in (net48): ribbon, task panes, COM injection, templates.
  PExL.Editor.Web/  Monaco editor assets, PExL syntax (pexl.lang.js), docs pane.
tests/
  PExL.Core.Tests/  xUnit golden tests (PExL → formula and formula → PExL).
docs/
  language-spec.md  the authoritative v1 spec — the contract the lexer/parser/emitter meet.
README.html         interactive site; embeds a JavaScript port of the transpiler ("PExLT").
brand/              logos and brand guidance.
```

## Build & test

```bash
# build everything
dotnet build PExL.sln -c Release -p:Platform=x64

# run the language test suite (no Excel required)
dotnet test tests/PExL.Core.Tests/PExL.Core.Tests.csproj

# build just the Excel add-in (.xll)
dotnet build src/PExL.AddIn/PExL.AddIn.csproj -c Release -p:Platform=x64
```

The add-in `.xll` and its sibling files land in `src/PExL.AddIn/bin/x64/Release/net48/`.

## Changing the language? Keep these in sync

A language change usually touches several places. Please update **all** that apply, in the
same PR:

1. **`src/PExL.Core`** — the lexer/parser/emitter (or `Decompile/` for the reverse direction).
2. **`tests/PExL.Core.Tests`** — add/adjust golden cases. New behavior needs a test.
3. **`src/PExL.Editor.Web/pexl.lang.js`** — vocabulary for highlighting, completions, hovers.
4. **`README.html`** — the in-browser transpiler `PExLT` is a faithful port of the C# engine;
   mirror grammar/emit changes here so the playground stays accurate.
5. **`docs/language-spec.md`** and the README cheat sheet / glossary — keep docs truthful.

> Rule of thumb: a verb isn't "done" until C#, the JS port, the spec, and a golden test agree.

## Coding style

- C# follows the repo `.editorconfig` (4-space indent, braces on new lines, `System` usings
  first). Prefer clear names over comments; comment **why**, not what.
- Keep `PExL.Core` free of any Excel/COM dependency so it stays unit-testable in isolation.

## Pull requests

1. Fork and branch from `main` (e.g. `feature/wildcard-lookup`).
2. Make focused commits; ensure `dotnet build` and `dotnet test` pass.
3. Fill in the PR template and link any related issue.
4. CI (build + tests on Windows) must be green before review.

## Reporting bugs / requesting features

Use the issue templates. For a transpiler bug, include the **PExL input**, the **formula you
got**, and the **formula you expected** — that makes it instantly reproducible as a test case.
