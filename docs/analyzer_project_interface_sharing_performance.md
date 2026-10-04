# Project interface sharing in Set assignments

A Set between two project object types can succeed when a project surface implements both interfaces. After direct Implements/cast checks, the previous compatibility helper reread every project surface's Implements list for every eligible statement until finding a common implementer. Both incompatible pairs and a shared implementer near the end repeatedly scanned the same metadata.

The Set-rule invocation now owns a lazy incremental lookup of implemented names to numeric surface identities and a cache of queried unordered pairs. Existing memberships can answer a new query; otherwise the cursor resumes until a common surface is found or metadata is exhausted. Each examined Implements list is read once by this lookup, misses and hits are retained, and the first common surface stops the scan. Reversed pairs reuse the same sharing answer. No full interface-pair graph is generated for a surface with many interfaces.

Direct Implements/cast checks remain outside the symmetric lookup. All metadata kinds remain eligible in the sharing scan, unlike unique object-type resolution; duplicate surface names do not combine separate Implements lists. Matching preserves exact full lowercased interface names with no trimming or qualification removal. Both expected/actual and held-object compatibility checks share the lookup across procedures in one pass. Other callers keep their default uncached behavior. Metadata is stable within the pass and a fresh rule invocation observes later mutations. The separate raw object-type resolver and ActiveSheet membership predicate remain unchanged.

## Deterministic cost

Actual Implements getters outside timed runs, with 100 unrelated classes plus Box, optional Shared and 100 assignments:

| Mode | Before reads | After reads | Diagnostics |
| --- | ---: | ---: | ---: |
| Same class | 1 | 1 | 0 |
| Different classes, no common implementer | 10,102 | 103 | 100 |
| Common implementer first | 102 | 3 | 0 |
| Common implementer last | 10,202 | 104 | 0 |
| 100 distinct pairs | 10,201 | 202 | 100 |
| Direct cast | 2 | 2 | 0 |
| Object target | 0 | 0 | 0 |
| Nothing value | 1 | 1 | 0 |

Counts include unique object-type resolution's initial Implements reads, explaining the remaining one/two reads beyond sharing metadata. Project name/kind counts remain unchanged in the original four counting probes. Counters are not in production.

## Reproduction and timings

From the repository root with dependencies installed:

```powershell
node scripts/benchmark-project-interface-sharing.mjs --baseline=1f73825b
node scripts/benchmark-project-interface-sharing.mjs
```

The baseline replaces assignments.ts and typeInference.ts from 1f73825b, after object-type resolution indexing, in the current harness. Fixtures have the listed unrelated class count plus Box and 1,000 Set writes. Same assigns New Box to Box; different assigns New K0 to Box without a common implementer. Shared modes add a surface implementing Box and K0 first/last. Distinct declares 1,000 variables with K types cycling over the unrelated classes, reaching 1,000 unique pairs at 1,000/3,000 classes, and assigns New Box. Direct-cast makes Box implement K0. Object assigns New K0 to Object; Nothing assigns Nothing to Box. Every sample checks 1,000 diagnostics for different/distinct and zero for all other modes.

Parsing, tokenization and binding are outside timing. Parsed bodies and tokens are reused; samples have fresh bound roots and fresh completion caches, with a root identity assertion. Measurements cover the Set rule with reused parsed bodies, not cold dataflow or full analyzer/editor latency. Three warmups precede 15 samples. Before/after timing runs were sequential after tests and differential probes completed, with no concurrent audit workloads. Separate getter probes do not affect timings.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           ; measured locally on 2026-10-03. Times in milliseconds.

| Surfaces / mode | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: |
| 10-same | 1.149 | 1.189 | 1.426 | 2.159 |
| 10-different | 1.553 | 1.444 | 2.663 | 1.891 |
| 10-shared-first | 1.286 | 1.167 | 1.381 | 1.370 |
| 10-shared-last | 1.541 | 1.199 | 2.080 | 1.326 |
| 10-distinct | 1.650 | 1.272 | 2.572 | 2.099 |
| 10-direct-cast | 1.194 | 1.228 | 1.432 | 1.745 |
| 10-object | 0.784 | 0.806 | 1.235 | 1.199 |
| 10-nothing | 1.379 | 1.328 | 1.774 | 1.845 |
| 100-same | 0.850 | 0.851 | 1.903 | 1.102 |
| 100-different | 3.544 | 1.023 | 4.007 | 1.347 |
| 100-shared-first | 1.217 | 1.132 | 1.361 | 1.865 |
| 100-shared-last | 3.712 | 1.126 | 4.297 | 1.273 |
| 100-distinct | 4.029 | 1.333 | 4.865 | 2.541 |
| 100-direct-cast | 1.116 | 1.180 | 1.285 | 2.186 |
| 100-object | 0.773 | 0.784 | 1.021 | 1.267 |
| 100-nothing | 1.061 | 1.076 | 1.418 | 1.739 |
| 1000-same | 0.848 | 0.861 | 1.221 | 0.902 |
| 1000-different | 24.764 | 1.092 | 25.897 | 1.406 |
| 1000-shared-first | 1.265 | 1.197 | 1.461 | 1.400 |
| 1000-shared-last | 24.801 | 1.201 | 28.759 | 1.398 |
| 1000-distinct | 25.998 | 1.955 | 26.626 | 3.087 |
| 1000-direct-cast | 1.143 | 1.195 | 1.327 | 1.750 |
| 1000-object | 0.770 | 0.792 | 1.071 | 1.185 |
| 1000-nothing | 1.086 | 1.105 | 1.115 | 1.942 |
| 3000-same | 0.952 | 0.955 | 1.037 | 1.173 |
| 3000-different | 69.974 | 1.190 | 78.577 | 1.514 |
| 3000-shared-first | 1.359 | 1.296 | 1.654 | 1.619 |
| 3000-shared-last | 70.580 | 1.307 | 73.831 | 1.535 |
| 3000-distinct | 73.199 | 2.013 | 76.591 | 2.552 |
| 3000-direct-cast | 1.246 | 1.285 | 1.563 | 1.828 |
| 3000-object | 0.772 | 0.786 | 0.886 | 1.039 |
| 3000-nothing | 1.320 | 1.189 | 2.242 | 1.613 |

At 3,000 classes, incompatible repeated pairs improve about 58.8 times, shared-last about 54.0 times and distinct pairs about 36.4 times. Early shared hits stay close to baseline. Bypass controls are mixed: direct-cast medians/p95 increase at all sizes, Object medians rise slightly, and some same/Nothing p95 values increase. The factory adds closures; its maps remain lazy until a shared-interface query. Retained memory is proportional to examined interface memberships and queried pairs, rather than all possible pairs. New queries may intersect existing surface sets; a first query still examines metadata until a hit or exhaustion. These synthetic results do not establish a universal latency improvement.

## Validation

- Type check passed.
- 266 focused tests passed, including nine new interface-sharing regressions.
- Full suite: 625 files, 12,843 tests passed, 13 skipped.
- New regressions cover actual read bounds across procedures for first/last/missing and distinct pairs, reversed queries, resumed lookups and earlier retained hits/misses, all metadata kinds, exact full names/no trimming, duplicate names on separate surfaces, many-interface metadata, direct casts and generic/host bypasses, mutable Implements lists, appended metadata and changing actual held types. Cached compatibility reasons match the default helper for varied types and pairs.
- 3,000 differential cases match complete analyzer and direct Set outputs with no internal errors, across four host models, metadata casing/kinds/duplicates/misses/interfaces, generic/host/library/project/scalar/Variant/qualified/array targets, held-object changes, scripting values, conditional declarations, branches/loops and multiple procedures.

Counting/differential probes remain audit scratch material. Timing benchmark and regressions are committed.
