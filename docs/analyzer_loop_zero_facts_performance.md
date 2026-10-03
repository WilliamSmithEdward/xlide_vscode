# Procedure-wide zero-start loop facts

Zero-start Do/While counter inference previously repeated a procedure-wide declaration walk, jump check and masked prefix check for each counter. The declaration-mask improvement in #780 removed repeated copies within one query; many queries still repeated the same setup.

Loop analysis now lazily prepares declaration eligibility and jump checks once per invocation. A masked procedure text indexes first whole-word mentions using the existing identifier-boundary character set. Counter queries compare their loop offsets with those first mentions. First declaration within a group and last matching group precedence stay unchanged; comments and strings remain conservative prior mentions.

If case folding changes text length, or a counter name lies outside the legacy word-character set, the query uses the existing prefix check. This preserves Unicode semantics rather than interpreting folded offsets as source offsets. Those exceptional paths still scan their prefixes. Facts are local to the invocation and are not retained by the existing weak loop-result cache. Explicit initialization and procedures without zero-start queries do not prepare them.

## Measurement

Run `node scripts/benchmark-loop-zero-facts.mjs --rounds=15`; compare with `--baseline=e0628f30`. The baseline is the #780 mask implementation. Each fixture declares N distinct numeric counters then contains N successive Do loops. The explicit-initialization control sets each counter immediately before its loop. Every run verifies the number of statements with counter facts.

Parsing/lexing happen outside timing. A fresh shallow body wrapper bypasses the loop-result cache each run, retaining source and parsed node identities.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups / 15 samples. Values are median milliseconds:

| Loops | Implicit zero before | Implicit zero after | Explicit init before | Explicit init after |
| --- | ---: | ---: | ---: | ---: |
| 100 | 1.336 | 0.265 | 0.152 | 0.225 |
| 500 | 27.095 | 1.139 | 0.666 | 0.670 |
| 1,000 | 116.813 | 1.547 | 1.011 | 1.535 |

Absolute timings fluctuate; the explicit control remains below 2 ms in this measurement. The operation-count guard is direct evidence: 300 declarations / 300 loops read declaration names 90,000 times before vs 300 after. This is isolated loop analysis, not end-to-end editor latency.

## Validation

5,000 generated LF/CRLF fixtures preserve complete loop-counter facts and statement spans, covering repeated loops, declaration precedence, eligibility, prior reads/comments, explicit initialization, jumps and Unicode names including expanded case folds. New regressions cover declaration read cost, distinct versus reused counters, word-boundary collisions and Unicode prefix behavior. Existing #780 regressions cover continued/late declarations and conservative exclusions.

Validation completed: 56 targeted tests and the type check passed; full suite 589 files / 12,386 tests passed (13 skipped).
