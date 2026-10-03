# Module-array shadow setup

The fixed-array bounds rule excludes module arrays hidden by a procedure's local variables and parameters. Previously its module-array filter rebuilt the procedure's entire shadow set for each array. With A module arrays and L locals, that portion of setup did O(A * L) work.

The rule now builds the shadow set once per procedure, inside the existing filter only when an array needs checking. Empty module-array lists still do not build it. This is invocation-local setup; separate procedures retain separate scopes.

## Measurement

Run `node scripts/benchmark-array-shadow-setup.mjs --rounds=15`; compare with `--baseline=07b1f618`. The baseline substitutes only the previous array rule. Parsing and binding happen outside timing. Each fixture has N module fixed arrays and N local scalar declarations, followed by one out-of-bounds module-array assignment. The control omits module arrays and uses a scalar assignment. Every run checks the finding count.

Node 24.18.0, AMD Ryzen 7 9800X3D, three warmups and 15 samples. The final measurements below ran after the test suite finished. Values are median milliseconds:

| N | Module arrays before | Module arrays after | No arrays before | No arrays after |
| --- | ---: | ---: | ---: | ---: |
| 100 | 0.658 | 0.253 | 0.020 | 0.030 |
| 1,000 | 37.850 | 1.365 | 0.249 | 0.116 |
| 3,000 | 445.643 | 2.739 | 0.660 | 0.295 |

Absolute timings varied: a five-sample probe measured 339 ms before at N=3,000; an earlier 15-sample run measured 584 ms before and 9.7 ms after. The fixed operation-count regression provides machine-independent evidence: with 500 arrays and 500 locals, total local-symbol name reads drop from 254,500 to 5,000. This stress benchmark measures the isolated rule, not end-to-end editor latency.

## Validation

102 targeted tests and the type check passed. The full suite passed 580 files / 12,248 tests (13 skipped). Complete diagnostic output matched baseline for 1,000 generated fixtures with LF/CRLF, Option Base, multiple procedures, case-insensitive local shadows, parameter shadows, local fixed arrays and unshadowed module accesses.

The new regression checks the local-name read budget, which the old rule exceeds, and verifies that bounds stay separate across local, parameter and module scopes.
