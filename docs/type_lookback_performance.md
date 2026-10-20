# Type-position lookback performance

Type-position detection filtered every significant token before the cursor to
remove newlines, then examined only a suffix. Completion and canonical casing
invoke this detector for ordinary identifiers as well as actual type names,
making a non-type check grow with the module prefix.

Read backwards until five non-newline tokens are collected. The longest suffix
the detector examines is `As New Library.Type`, which contains five tokens.
Newlines are skipped just as before, preserving existing continuation behavior.
The cursor context and candidate generation retain their existing behavior.

## Reproduce

```powershell
$env:XLIDE_TYPE_LOOKBACK_BENCH='1'
npx vitest run tests/typeLookbackPerformance.test.ts -t 'measures warm'
Remove-Item Env:XLIDE_TYPE_LOOKBACK_BENCH
```

Windows, 49,484-character / 1,201-procedure fixture, 21-sample medians after
warm-up. The casing request verifies three edits on one line at the module end.

| Analyzer request | Before (ms) | After (ms) |
| --- | ---: | ---: |
| Warm non-type position check | 0.2993 | 0.0013 |
| Warm line casing | 1.8767 | 0.8087 |

These are analyzer component timings, not keyboard-to-paint measurements.
The cached cursor context is warm; this change does not remove source lexing or
prefix-array construction after an edit. Long runs of newline tokens still
require skipping those tokens to preserve the existing detector's semantics.

The work-count regression observed 6,001 prefix token reads before the change
and seven afterwards for an ordinary expression. Behavioral coverage includes
`As New` with a qualified project class, continued qualified type names,
expression `New`, library qualifiers and ordinary expressions.

The real VS Code suite `Type lookback surfaces` checks qualified type completion
after edits to a continued declaration and actual typed/saved canonical casing
at the end of a large module.
