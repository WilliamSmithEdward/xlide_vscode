# Fixed-array exclusion snapshot reuse

checkFixedArraySubscriptBounds previously rebuilt its ReDim exclusion set for every statement, spreading all target names and filtering them against the reshaped-array map. Many statements share one immutable map, so this repeated identical work.

Keep one exclusion set per reshaped-map identity in the procedure's rule invocation. Statements without a reshaped map still use the original ReDim target set. Distinct states after erasures, later ReDims and restored branches get their own exclusions. The cache has no global lifetime or cross-activity reuse.

## Reproduction

```powershell
node scripts/benchmark-array-exclusions.mjs --baseline=0f08df1d --rounds=15
node scripts/benchmark-array-exclusions.mjs --rounds=15
```

The benchmark times the public fixed-array rule after parsing and binding, using three warmups, 15 samples and warmed token caches. Each fixture has 100 arrays, repeated in-bounds accesses and one final out-of-bounds access. The dynamic fixture ReDims all arrays before the accesses; the control declares fixed arrays. The harness requires exactly one diagnostic.

Node 24.18.0, AMD Ryzen 7 9800X3D; median milliseconds:

| Repeated statements | ReDimmed before | After | Fixed-array control before | After |
| ---: | ---: | ---: | ---: | ---: |
| 100 | 1.012 | 0.896 | 0.161 | 0.158 |
| 1,000 | 5.835 | 5.309 | 0.530 | 0.540 |
| 3,000 | 15.883 | 13.820 | 1.484 | 1.535 |

These isolated-rule stress fixtures do not represent editor latency. The improvement is modest because other flow and lookup work dominates this rule. Absolute timings vary with load. Counting iterations independently demonstrates the eliminated repeated setup: reading a tracked name from the 100-name target set fell from 201 iterations to one for 200 ordinary accesses plus the final access.

The cache retains a set for each distinct reshaped map until the procedure invocation finishes. Fully known snapshots commonly produce empty sets; snapshots excluding many names retain larger sets. This is a temporary memory tradeoff for avoiding repeated scans, with no global retained state.

## Validation

- 85 targeted tests passed.
- 3,000 generated fixtures matched complete diagnostic callback outputs against baseline 0f08df1d, including Option Base, known/unknown ReDims, erasures, conditional activity, jump labels, GoSub, loops and branch restoration.
- Four regressions verify the wide-set operation budget, changed exclusions after erasure/ReDim, activity changes on reused parsed nodes and reachable branch restoration.

- Type checking passed; full suite: 596 files, 12,528 tests passed, 13 skipped.
