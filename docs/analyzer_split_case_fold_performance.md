# Split case-fold preparation

Array-shape inference for `Split` with text comparison searched each delimiter through a helper that checked the whole input for ASCII and lowercased it again. With many delimiters this repeated whole-string work for every output part.

The helper now prepares an ASCII-validated, folded search once, reused by successive delimiter searches. Output slices still use the original text. Filter uses the same helper for its single search per element. Limit one retains its no-comparison behavior, including non-ASCII input; limit zero and empty text still bypass Split searches.

## Measurement

Run `node scripts/benchmark-split-case-fold.mjs --rounds=15`; compare with `--baseline=07b1f618`. The baseline substitutes only the previous array rule. Tokenization happens outside timing. Fixtures contain `AaX` repeated N times, with text comparison against `x`; binary comparison against `X` and text comparison with limit one are controls. Each call checks the output part count.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups, 15 samples. No own test runs overlapped the measurement. Values are median milliseconds:

| N | Text before | Text after | Binary before | Binary after | Limit one before | Limit one after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 1,000 | 3.138 | 0.060 | 0.035 | 0.061 | 0.002 | 0.003 |
| 10,000 | 378.630 | 0.504 | 0.266 | 0.373 | 0.003 | 0.004 |
| 30,000 | 3,528.432 | 1.248 | 0.903 | 1.256 | 0.004 | 0.005 |

This is an isolated helper stress test, not end-to-end editor latency. Absolute timings fluctuate; the operation-count regression proves the repeated folding is removed. For 1,000 delimiters / 3,000 characters, folded input characters fall from 3,003,000 to 3,000.

## Validation

6,000 complete Split/Filter shape outputs match the previous rule across ASCII/non-ASCII text and delimiters, case-sensitive/caseless comparison, limits, empty input, quoting, include/exclude Filter and Option Base. New regressions cover the folding budget, original-case parts, repeated/multi-character delimiters, unmatched delimiters, binary behavior, conservative non-ASCII handling and no-comparison limits.

Validation completed: 29 targeted tests and the type check passed. The complete suite passed 580 files / 12,255 tests (13 skipped).
