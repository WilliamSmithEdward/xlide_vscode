# Identifier symbol snapshot performance

Repeated bare-identifier completion and automatic casing rebuilt the entire
module symbol graph on every request. The existing resolver shared that graph
only between positions in one request, even though the lexer and parser already
retain bounded snapshots for unchanged source.

Completion now privately retains eight symbol projections, keyed by the complete
source value, module name and module kind, with least recently used eviction.
Projection remains lazy. An edit or module identity change rebuilds it; cursor
scope is selected on every request. Completion records are freshly allocated and
project/host context is merged anew. No other symbol-builder consumer receives
these cached graphs.

## Reproduce

PowerShell, from the repository root:

```powershell
$env:XLIDE_IDENTIFIER_CACHE_BENCH='1'
npx vitest run tests/identifierSymbolCachePerformance.test.ts -t 'measures repeated'
Remove-Item Env:XLIDE_IDENTIFIER_CACHE_BENCH
```

The fixture is 118,024 characters with 1,201 procedures and documentation on
1,200 procedures. Both requests resolve the same local variable; the casing
request produces one verified edit. Host globals and runtime completions are
disabled. One warm-up precedes 21 samples, whose median is reported. These are
analyzer component measurements, not end-to-end keyboard or UI timings.

Measured locally on Windows using the same fixture, before and after this change:

| Request | Before (ms) | After (ms) |
| --- | ---: | ---: |
| Repeated identifier completion | 1.1988 | 0.4432 |
| Repeated line casing | 0.9588 | 0.2788 |

New source revisions still pay the projection cost. This change removes repeated
projection for the same source, including completion and casing requests sharing
a document snapshot; it does not claim faster cold parsing after each keystroke.

Unit regressions verify single projection across completion/casing, independent
returned records, source invalidation, procedure scope, module identity, fresh
external context, lazy construction and bounded eviction.

The real VS Code 1.139.1 integration suite `Identifier snapshot surfaces` checks
repeated local completion in a 1,202-procedure module, then edits a declaration
and checks completion, hover and `ThisWorkbook.Sheets(1).` member completion.
