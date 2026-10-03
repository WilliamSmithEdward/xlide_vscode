# Loop zero-start declaration masking

Loop-counter inference can prove that an untouched numeric `Dim` local starts at zero. Its prior-text check blanks declarations from the procedure prefix. Previously it rebuilt that entire prefix for every declaration with two slices and concatenation.

The existing declaration walk now collects ranges. After the eligibility and jump checks pass, the prefix is assembled from original gaps and blank ranges once. Sorting and merging overlapping ranges preserve positions. Declarations, comments, prior mentions and jump detection retain their existing semantics; no procedure facts are cached beyond the existing loop-fact cache.

## Measurement

Run `node scripts/benchmark-loop-zero-declarations.mjs --rounds=15`; compare with `--baseline=0596025d`. The baseline substitutes only the old loop-counter implementation. Each fixture has N padding declarations plus a numeric counter declaration followed by a Do loop. The explicitly initialized control inserts `i = 0` immediately before the loop, bypassing zero-start inference.

Parsing/lexing happen outside timing. A fresh shallow body wrapper bypasses the loop-fact cache each run while retaining the same source and parsed nodes. Every run checks that one statement receives loop-counter facts.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups / 15 samples. Values are median milliseconds:

| Padding declarations | Implicit zero before | Implicit zero after | Explicit init before | Explicit init after |
| --- | ---: | ---: | ---: | ---: |
| 100 | 0.067 | 0.049 | 0.006 | 0.005 |
| 1,000 | 0.509 | 0.186 | 0.006 | 0.011 |
| 3,000 | 3.058 | 0.319 | 0.011 | 0.011 |

Absolute timings vary: the initial five-sample probe measured 8.9 ms before at 3,000 declarations. The copied-character guard provides direct evidence: for 1,000 declarations in a 16,959-character module, long-string slices produce 16,938,848 characters before vs 34,848 after. These are isolated loop-analysis stress measurements, not end-to-end editor latency.

## Validation

3,000 generated LF/CRLF fixtures preserve complete loop-counter facts and statement spans, including numeric/Variant/suffixed declarations, continued declarations, static/object/array/Const exclusions, prior reads/comments, explicit initialization and jumps. Regressions cover the copied-character budget, multiline and late declarations, conservative mention checks and excluded declaration types.

Validation completed: 60 targeted tests and the type check passed; full suite 585 files / 12,348 tests passed (13 skipped).
