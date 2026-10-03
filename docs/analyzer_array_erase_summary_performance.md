# Array-erasure procedure summary performance

The unallocated-array rule builds a summary of procedures that erase ByRef array parameters. Previously it scanned every procedure for early exits, even when there were no eligible parameters, then searched the body separately for each eligible parameter. Cached tokens avoided re-lexing but did not avoid repeated token traversal and name folding.

The summary now checks parameter eligibility first, stops inspecting statement spans after discovering an early exit, and collects final mentions and nested mentions for all eligible names in one body pass. Parameter indexes and duplicate-name behavior are preserved. Inactive non-leaf mentions still conservatively suppress erasure effects, matching the previous implementation; inactive leaf mentions remain excluded.

## Reproduction

Run from the checkout:

```powershell
node scripts/benchmark-array-erase-summary.mjs --baseline=9d091699 --rounds=15
node scripts/benchmark-array-erase-summary.mjs --rounds=15
```

The benchmark times the public unallocated-array rule with parsing and binding outside the timed region. It uses three warmups and 15 samples, with warmed statement-token caches. The ordinary fixture has no array parameters. The array fixture has 100 ByRef array parameters, a final erasure for each, and an allocated caller array whose later UBound produces one diagnostic. The harness verifies those finding counts.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds:

| Assignment statements | Ordinary before | Ordinary after | 100 array parameters before | After |
| ---: | ---: | ---: | ---: | ---: |
| 100 | 0.031 | 0.008 | 0.883 | 0.199 |
| 1,000 | 0.154 | 0.019 | 4.454 | 0.344 |
| 10,000 | 0.714 | 0.063 | 48.115 | 1.157 |

These are isolated-rule stress fixtures, not editor latency measurements. The 100-parameter fixture is below VBA's 255-parameter limit. Absolute timings vary with machine load; an earlier five-sample probe measured the largest array fixture at 54.884 ms before the change.

## Validation

- Type checking passed.
- Full suite: 590 files, 12,393 tests passed, 13 skipped.
- Targeted array and ByRef suites: 96 tests passed.
- 3,000 generated fixtures produced identical complete diagnostic callback outputs against baseline 9d091699, including conditional compilation, early exits, nested and single-line branches, shadowing, and parameter modes.
- With 200 assignments and 25 eligible parameters, folds of the assignment name fell from 5,200 to 400. A regression test bounds that work independently of wall-clock timing.
- A second operation-budget test ensures a procedure with no eligible parameters does not inspect each of its 1,000 statement spans. Behavioral cases cover final versus earlier erasures, early exits, ByVal and scalar parameters, and inactive nested mentions.
