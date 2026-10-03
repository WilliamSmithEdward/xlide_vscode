# Project-class object-read facts performance

The object-default-value rule looked up a declared class with a linear project-surface search for each value read and For Each loop. It also repeatedly searched the same members for a default or enumerator and reinterpreted the default signature. Metadata is stable during one synchronous analysis invocation.

Cache results by queried normalized type name, including absent and incomplete classes. Preserve the first matching class, ignoring other surface kinds; an incomplete first class still prevents a later duplicate from proving absence. The shared completion index excludes ambiguous duplicate names and cannot replace these semantics. Cache each touched class's default-read facts (presence and read problem), and its first -4 enumerator or absence. Pass the enumerator lookup through the existing recursive walk. Each rule invocation starts fresh, so subsequent calls see changed members/signatures even on reused metadata objects and arrays.

This uses O(queried type names + classes whose default/enumerator is inspected) extra entries for the invocation lifetime. It does not retain a full index of unused project classes. Each distinct queried type still performs one existing linear surface scan, and each relevant member list still has its first scan; this is not a claim of constant cost for every project size or a general analyzer-wide class cache.

## Reproduce

```powershell
node scripts/benchmark-object-class-facts.mjs --baseline=1ff0c6be22e47e0a112ed8dbb58842464b813e8c
node scripts/benchmark-object-class-facts.mjs
```

Only objectValues.ts is substituted for the baseline bundle; surrounding code is identical. Parse once, build fresh bound symbols outside every timed sample (assert distinct root identity), then time rule factory setup and all procedure statement/header visitors. Three warmups, 15 samples; the script reports median and p95. Run without other tests or benchmarks in parallel.

Fixtures repeat 1,000 reads or loops. class-lookup puts the used class after the stated number of preceding classes. missing-class omits it and must produce no hits. default-member uses one class with the stated number of preceding ordinary members then a default requiring an argument. indexed-member has no default. enumerator puts the -4 member last, returning Collection. All positive cases must produce exactly 1,000 findings. scalar-control assigns n = 1 while the procedure still declares an unused c As Target and receives the class pool; it must produce no hits. The numbers in the first column are preceding class/member entries, not typical workbook sizes or total end-to-end analysis times.

## Results

2026-10-03, Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor. Rule medians in milliseconds:

| Preceding entries and mode | Before | After |
| --- | ---: | ---: |
| 10-class-lookup | 1.352 | 1.325 |
| 10-missing-class | 1.006 | 0.956 |
| 10-default-member | 1.156 | 1.121 |
| 10-indexed-member | 1.218 | 1.119 |
| 10-enumerator | 1.902 | 1.401 |
| 10-scalar-control | 0.292 | 0.294 |
| 100-class-lookup | 1.888 | 1.111 |
| 100-missing-class | 1.303 | 0.919 |
| 100-default-member | 0.893 | 0.844 |
| 100-indexed-member | 0.759 | 0.803 |
| 100-enumerator | 5.067 | 1.082 |
| 100-scalar-control | 0.291 | 0.293 |
| 1000-class-lookup | 9.898 | 1.103 |
| 1000-missing-class | 5.176 | 0.911 |
| 1000-default-member | 1.442 | 0.752 |
| 1000-indexed-member | 1.025 | 0.750 |
| 1000-enumerator | 50.548 | 1.126 |
| 1000-scalar-control | 0.320 | 0.291 |
| 3000-class-lookup | 29.782 | 1.144 |
| 3000-missing-class | 15.674 | 0.942 |
| 3000-default-member | 2.938 | 0.759 |
| 3000-indexed-member | 1.685 | 1.229 |
| 3000-enumerator | 118.592 | 1.251 |
| 3000-scalar-control | 0.284 | 0.293 |

At 3,000 entries, class lookup falls from 29.782 to 1.144 ms, missing-class lookup 15.674 to 0.942 ms, default reads 2.938 to 0.759 ms, indexed reads 1.685 to 1.229 ms and enumerators 118.592 to 1.251 ms. Small cases are mixed (100-member indexed: 0.759 to 0.803 ms). The 3,000-class scalar control is 0.284 to 0.293 ms. These synthetic rule-only cases establish repeated-work reductions, not portable latency or end-to-end speedups. Earlier probes varied notably, particularly enumerators (about 132/156 ms before versus 119 ms in the final run).

Stable getter counters independently verify work on 100 repeated uses with 100 preceding entries:

| Work | Before | After |
| --- | ---: | ---: |
| Class-name reads for default values | 10,200 | 300 |
| Default-member flag reads | 10,100 | 200 |
| Class-name reads for enumerators | 20,200 | 300 |
| Enumerator member-attribute reads | 10,000 | 100 |

The default flag still has a scan by the existing type-verdict helper in addition to this rule's first scan; the change does not claim all shared type inference is optimized.

## Validation

- Four new tests cover bounded actual getter reads, first-class/incomplete/other-kind semantics, updated defaults/enumerators across calls using the same metadata array/object, qualified/case-insensitive types and repeated enumerator report order.
- Five focused files: 47 tests passed; typecheck passed.
- Full suite on the final implementation: 605 files, 12,629 tests passed, 13 skipped.
- 2,000 complete analyzeModule outputs exactly match baseline 1ff0c6be, including messages, spans and order, without internal errors. Cases include required/optional/ParamArray/zero-argument defaults, Collection-returning defaults/enumerators, duplicate/incomplete and mixed-kind surfaces, qualified/array declarations, local shadowing, ten forms and conditional activity.
