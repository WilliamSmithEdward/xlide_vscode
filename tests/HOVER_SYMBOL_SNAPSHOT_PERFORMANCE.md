# Hover symbol snapshot performance

Bare-symbol hover reconstructs module symbols before checking local declarations
and before falling back to host globals. Repeated mouse positions over one
unchanged module paid that projection repeatedly, despite parser/lexer caches
and the existing bounded identifier-completion symbol cache.

The completion cache now lives in `symbols/editorModuleSymbols.ts`, shared with
hover. Only these read-only editor resolvers borrow the graphs. Other symbol
builder consumers retain their own graphs. Eight snapshots are retained, keyed
by complete source, module name and module kind, with least recently used
eviction. Hover records, scope selection and external documentation are fresh
for every request.

## Reproduce

```powershell
$env:XLIDE_HOVER_SYMBOL_BENCH='1'
npx vitest run tests/hoverSymbolSnapshotPerformance.test.ts -t 'measures warm'
Remove-Item Env:XLIDE_HOVER_SYMBOL_BENCH
```

Windows, 55,495-character / 1,202-procedure source, 21-sample medians after a
verified warm-up hover. Requests target a local variable near the module end
and the bare host global `ThisWorkbook`.

| Analyzer hover | Before (ms) | After (ms) |
| --- | ---: | ---: |
| Local variable | 0.6178 | 0.0153 |
| Host global | 0.4540 | 0.0525 |

These are analyzer component timings, not mouse-to-tooltip latency. A new source
revision still pays symbol construction. Completion/casing and hover can reuse
one projection when their module identity and source agree.

Unit regressions cover one projection shared with completion, fresh returned
records, procedure scope, source/type changes, module identity, updated external
documentation and newly introduced host-global shadowing. Existing cache tests
retain coverage for bounded eviction and lazy construction.

The VS Code 1.139.1 suite `Hover snapshot surfaces` repeatedly requests hover
over locals in different procedures and a host global in a large module, then
edits declarations and verifies updated hover and scoped shadowing. Chained
`ThisWorkbook.Sheets(1).` completion in another procedure must still work.
