# Local declaration-order pass

Every reported forward reference previously spread all local declarations into an array and searched it for a containing Const line. That work ran even under Option Explicit, where the answer was unused. Last-declaration discovery also spread all declaration starts into Math.max, and the rule carried its own copy of the binary token-offset helper.

Track the last declaration during the existing collection pass. Resolve Const-line membership only when a non-explicit unresolved reference needs it; collect distinct Const physical-line spans once, then advance a cursor as uses arrive in source order. Reuse firstTokenAtOrAfter instead of the duplicate helper.

The original declaration selection, inactive-branch exclusions, shadowing, diagnostic order and physical-line Const semantics are preserved. Span state stays local to the procedure pass.

## Reproduction

Run node scripts/benchmark-local-declaration-order.mjs --rounds=15 against the target checkout and baseline 21c56500. The runner bundles local source with the existing esbuild dependency, parses/binds fixtures before timing and reports 15-sample median/p95 after three warmups. No test suite ran concurrently with these measurements.

Node 24.18.0, AMD Ryzen 7 9800X3D, Windows; milliseconds:

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 100-forward-uses/explicit/statement | 0.163 | 0.075 | 0.937 | 0.186 |
| 1000-forward-uses/explicit/statement | 2.001 | 0.383 | 2.48 | 0.631 |
| 5000-forward-uses/explicit/statement | 40.007 | 1.486 | 45.319 | 2.653 |
| 1000-forward-uses/implicit/const | 12.958 | 0.43 | 13.186 | 1.101 |

The 5,000-forward-use fixture is approximately 27 times faster; the implicit Const fixture is approximately 30 times faster. These are warm synthetic rule-pass measurements excluding parsing/binding, not whole-workbook latency guarantees. The runner requires the original diagnostic count on every sample.

## Validation

Both enumeration regressions fail on original code: 160,400 declaration visits in the explicit case and 320,800 in the implicit Const case, each above an 8,000-visit budget. They pass with the fix and require the correct diagnostic kinds/counts. Additional regressions preserve colon-separated Const physical-line coverage and exact spans under LF and CRLF. All nine existing local declaration-order tests pass.
Type checking passed. The full Vitest suite passed: 561 files, 12,000 tests passed and 13 skipped.
