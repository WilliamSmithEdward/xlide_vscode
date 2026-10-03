# Array element write equality depth

Array element write detection checks whether a statement is a call by looking for a top-level equals sign. Previously each equals token separately scanned the entire preceding token prefix to compute parenthesis depth. A call argument containing many comparisons repeated those scans.

The detector now uses the existing shared `topLevelEqualsIndex` helper. It scans once, retains the same raw parenthesis-depth semantics and removes the redundant private `topLevelAt` function. Tokens come from the lexer: a raw `=` token is an operator, while a string containing an equals sign retains its quoted raw text.

## Measurement

Run `node scripts/benchmark-array-write-depth.mjs --rounds=15`; compare with `--baseline=217f53c3`. The baseline substitutes only the old array rule. Parsing and lexing happen outside timing. Each fixture calls `Fill` with N comparisons joined by `And`, followed by `target(0)`; the early-assignment control prefixes the same call with `sink =`. Every run checks the element-write set.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups / 15 samples. Values are median milliseconds:

| Comparisons | Call before | Call after | Assignment before | Assignment after |
| --- | ---: | ---: | ---: | ---: |
| 100 | 0.177 | 0.035 | 0.055 | 0.051 |
| 1,000 | 17.408 | 0.253 | 0.146 | 0.171 |
| 3,000 | 164.232 | 0.659 | 0.488 | 0.462 |

Absolute timings varied: an earlier five-sample probe measured 74 ms before at 3,000 comparisons. The operation-count regression provides machine-independent evidence: at 1,000 comparisons, raw-token text reads drop from 5,528,512 to 28,012. The early-assignment path remains comparable because the old detector already stopped at its first top-level equals sign. This is an isolated helper stress measurement, not end-to-end editor latency.

## Validation

10,000 generated expressions produce exactly the same element-write sets in the same insertion order as the previous implementation. They include malformed parentheses, nested comparisons, assignments, calls, quoted equals signs and writing statements. New regressions cover the operation budget, plain and parenthesized calls, Let/ordinary assignments, read-only array elements, nested comparisons, string literals and unmatched parentheses.

Validation completed: 146 targeted tests and the type check passed; full suite 583 files / 12,287 tests passed (13 skipped).
