# Type-position lookback performance

This note records the lookback as first bounded. A later change made a line
break end the type position and removed the lookback across lines; see
"Later: a line break ends the type position" at the end.

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
npx vitest run tests/analyzer/typeLookbackPerformance.test.ts -t 'measures warm'
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

## Later: a line break ends the type position

Skipping newlines made the row below a declaration a type position: on an
empty row under `Dim x As Long` the detector read `As Long` and completion
offered only `Long`, `LongLong` and `LongPtr`, hiding keywords, snippets and
identifiers. A bare `As` or expression `New` followed by Enter offered the
whole type list on the next row. VBA ends a statement at a line break unless
` _` continues it, and the lexer keeps a continuation as trivia rather than a
newline token, so the detector now reads only the cursor's logical line. The
chunked lookback across earlier lines and its full-stream fallback are removed.

`typeLookbackPerformance.test.ts` and `cursorLexerWindowWork.test.ts` now check
that a type position stops at a line break and that a long tail of comment or
blank rows below an `As` lexes only the cursor's row.

Analyzer timings at the end of a large module (1,200 procedures), Linux, at
087a9c61. Warm: two alternating runs of 2,000 calls after a warm-up. Cold: the
first call in a fresh process.

| Caret position | Warm median, before / after | Warm p95, before / after | Cold first call, before / after |
| --- | ---: | ---: | ---: |
| Empty row below a bare `As` | 170–220 / 0.6 µs | 234–455 / 1.7–2.6 µs | 46 / 0.6 ms |
| Empty row below `New` | 163–168 / 0.6 µs | 242–302 / 1.7 µs | 59 / 0.5 ms |
| Empty row below `As Long` | 196–202 / 0.5 µs | 209–216 / 1.6–8.3 µs | 82 / 0.6 ms |
| Empty row after 40 comment rows | 209–214 / 0.4–0.6 µs | 250–334 / 1.5–2.0 µs | 54 / 0.6 ms |
| `Dim total As Lo` (type position) | 199–201 / 196–235 µs | 230–247 / 282–493 µs | 40 / 42 ms |
| Ordinary expression line | 2.8–3.3 / 2.7–3.5 µs | 4.1–8.5 / 3.3–8.3 µs | 2.5 / 1.4 ms |

The ordinary-line warm row comes from a separate run of that case alone
(20,000 calls after 20,000 warm-up calls, four runs), where the two trees
measure the same. In shorter mixed runs it measured a few microseconds slower
after the change; the cause of that gap was not established. The type-position
cost is building the candidate list, which this change leaves as it was.
