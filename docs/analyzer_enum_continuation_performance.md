# Enum continuation validation

The continuation rule collected source-ordered continuation trivia, then compared every item with every Enum body. Seek directly to the first item after the Enum header and stop at the Enum end. The existing binary token helper now accepts the start-offset fields it actually needs, allowing reuse for trivia without another search implementation.

Diagnostic ordering, spans, continuation limits, empty-line checks and Enum-header exclusions are preserved. The source-ordered trivia array remains local to the rule pass.

## Reproduction

Run node scripts/benchmark-enum-continuations.mjs --rounds=15 against the target checkout and baseline f7aa0bb3. The runner bundles local source with the existing esbuild dependency, parses fixtures before timing, takes three warmups and reports median/p95 over 15 samples. Each Enum contains one invalid continuation and the runner checks the exact diagnostic count.

Node 24.18.0, AMD Ryzen 7 9800X3D, Windows. Times are milliseconds:

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 100-enums/one-continuation | 0.119 | 0.025 | 0.278 | 0.397 |
| 1000-enums/one-continuation | 1.238 | 0.143 | 1.531 | 0.452 |
| 5000-enums/one-continuation | 27.96 | 0.862 | 32.193 | 0.914 |
| 1000-enums/continuation-outside-enums | 0.43 | 0.444 | 0.784 | 0.629 |

The 5,000-Enum fixture is approximately 32 times faster. These are warm synthetic rule-pass measurements, excluding parsing. The unrelated-continuation control remains approximately unchanged; no improvement is claimed for it.

## Validation

The complexity test counts 94,125 trivia-start reads for 250 Enums against the original implementation, exceeding the 7,500-read budget. It passes after the fix. Semantic checks verify CR/LF/CRLF source boundaries, permitted header continuations, unrelated procedure continuations and the original diagnostic spans. Existing continuation-limit and binary-token-helper regressions also pass.
Type checking passed. The full Vitest suite passed: 558 files, 11,987 tests passed and 13 skipped.
