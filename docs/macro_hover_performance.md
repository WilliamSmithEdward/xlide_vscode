# Macro-string and empty-hover performance checks

`macroHoverSurfacePerformance.test.ts` checks token-read counts near the end of a
large module, macro string boundaries, project-free rejection of impossible hover
positions, and fallback to project metadata for unknown names and macro strings.

```powershell
npx vitest run tests/macroHoverSurfacePerformance.test.ts tests/macroNames217.test.ts tests/vbaHover.test.ts tests/vbaMemberCompletion.test.ts
```

The opt-in benchmark uses 21 warmed samples with a 1,200-procedure, 58,953-byte
module. Run on an otherwise idle machine:

```powershell
New-Item -ItemType Directory -Force out | Out-Null
$env:XLIDE_MACRO_HOVER_BENCHMARK_OUTPUT = 'out/macro-hover-latency.json'
npx vitest run tests/macroHoverSurfacePerformance.test.ts -t 'measures macro lookup and empty hover work'
Remove-Item Env:XLIDE_MACRO_HOVER_BENCHMARK_OUTPUT
Get-Content out/macro-hover-latency.json
```

The macro timing checks a dot position outside a string. The hover timing uses a
comment position and models the uncached provider path: local context builds a
real module index, while the asynchronous project-context method is stubbed.
The report includes how often that method was invoked. Source, lexer and parser
caches warm normally across samples. These are component CPU timings with
mocked VS Code; they exclude actual project I/O, mouse-hover delay and rendering.

The real-host editor suite verifies empty hovers over comments (including
continued comments), numbers, dates and indentation, while retaining ordinary
symbol hover and cross-module macro-name string hover:

```powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Editor surfaces'
```
