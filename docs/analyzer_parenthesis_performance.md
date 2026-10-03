# Parenthesis validation performance

The parenthesis rule matched every opening parenthesis by scanning forward, then copied and searched each grouping range. Nested expressions repeatedly visited and allocated overlapping ranges.

Statements with multiple groups now build matching-close, depth and first-named-token indexes in linear passes. Group validation uses range bounds rather than copied inner arrays; the first named token is still checked with the same depth semantics as before. A single group retains the direct matching path to avoid index allocation for ordinary statements.

## Reproduction

Run node scripts/benchmark-parenthesis-validation.mjs --rounds=15 against the target checkout and baseline 21c56500. The runner bundles local source using the existing esbuild dependency, parses/binds fixtures before timing and reports median/p95 over 15 samples after three warmups. No test suite ran concurrently with these measurements.

Node 24.18.0, AMD Ryzen 7 9800X3D, Windows; milliseconds:

| Fixture | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 20-nested-parentheses | 0.016 | 0.009 | 0.028 | 0.018 |
| 200-nested-parentheses | 0.264 | 0.025 | 1.149 | 0.045 |
| 2000-nested-parentheses | 6.801 | 0.185 | 8.641 | 1.845 |
| 1000-simple-groups | 0.327 | 0.354 | 0.386 | 0.397 |

The 2,000-level fixture is approximately 37 times faster. That fixture exercises typing/recovery beyond the expression parser nesting limit and physical-line limits; it is not a claim about valid workbook latency. The simple-group control remains sub-millisecond and no improvement is claimed for it.

## Validation

The complexity regression observes 323,206 token-text reads for 400 nested groups against the original rule, above a 16,000-read budget. It passes after the fix. Exact code/message/span outputs matched between original and updated implementations for 1,009 deterministic expression fixtures, including malformed nesting, unmatched delimiters, collection values, named arguments and member access after groups. All 39 existing parenthesis diagnostic tests pass.
Type checking passed. The full Vitest suite passed: 559 files, 11,987 tests passed and 13 skipped.
