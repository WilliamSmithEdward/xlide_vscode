# Empty project-member lookup performance

Issue #897: projectTypeAt and privateMemberOwnerAt resolved receiver prefixes and bindings even when their existing project surface index had no entries. projectClassMemberAt inherited that work. An empty index cannot produce a project result, including when all supplied names are ambiguous.

Read the existing index first and return undefined if it is empty; otherwise keep the receiver resolution and member matching. Reuse the index for the final lookup. This adds no cache and keeps new metadata lists, first case-insensitive member matching, private/public behavior and non-class project surfaces intact. The benefit applies to empty indexes; populated workbook/class/form metadata follows the existing resolution path.

Validation: nine work-counter regressions failed on baseline b8e44324 and three behavior controls passed. All 12 pass after the change. For each public API, 1,000 absent/empty/ambiguous-index requests perform zero receiver-token-prefix reads instead of 1,000. 209 focused tests and type checking pass. Full suite: 646 files, 13,138 tests passed, 13 skipped.

Complete diagnostic parity: 8,256 oracle corpus sources × four hosts × five metadata contexts = 165,120 analyses, all outputs identical and no internal errors. Cached statement token arrays are frozen. Contexts cover absent, explicit empty, ambiguous, class/public/private, and standard module metadata.

Benchmark: Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           , 15 rounds after three warmups, baseline b8e44324. Direct rows call all three public APIs per reference with pre-parsed AST/shared tokens/pre-indexed metadata. Full rows run analyzeModule with a fresh metadata list per round; full-fresh also changes the source key. Negative fixtures use a host Range receiver, class controls use C. Complete result/diagnostic hashes match in all 36 main rows. Two invalid-offset controls also preserve undefined results. Assertions run outside the clock. These are analyzer/provider timings, not editor or Office latency.

The direct negative lookup gain is consistent in isolated checks. Whole-analyzer timings vary and do not establish a consistent speedup. Populated-index controls keep their behavior; fresh positive class controls were modestly slower in isolated runs. A first invalid-offset lookup with 1,001 class entries now builds the index earlier (about 0.069 ms in this run), versus near-zero before. These tradeoffs are retained in the tables rather than hidden by an aggregate.

| References | Metadata | Scope | Before median ms | After median ms | Before p95 ms | After p95 ms |
| ---: | --- | --- | ---: | ---: | ---: | ---: |
| 1 | absent | direct | 0.007 | 0.001 | 0.029 | 0.005 |
| 1 | absent | full | 1.161 | 1.150 | 1.977 | 1.829 |
| 1 | absent | full-fresh | 1.034 | 1.144 | 1.244 | 1.282 |
| 1 | empty | direct | 0.004 | 0.000 | 0.006 | 0.002 |
| 1 | empty | full | 0.831 | 0.850 | 1.135 | 1.050 |
| 1 | empty | full-fresh | 0.899 | 0.869 | 1.718 | 1.795 |
| 1 | ambiguous | direct | 0.003 | 0.000 | 0.003 | 0.011 |
| 1 | ambiguous | full | 0.838 | 0.810 | 1.555 | 1.483 |
| 1 | ambiguous | full-fresh | 0.839 | 0.827 | 1.083 | 1.061 |
| 1 | class | direct | 0.004 | 0.004 | 0.012 | 0.014 |
| 1 | class | full | 0.717 | 0.710 | 0.835 | 0.795 |
| 1 | class | full-fresh | 0.785 | 0.775 | 1.461 | 1.468 |
| 100 | absent | direct | 0.161 | 0.012 | 0.253 | 0.033 |
| 100 | absent | full | 5.311 | 4.952 | 9.156 | 7.503 |
| 100 | absent | full-fresh | 5.595 | 5.445 | 7.478 | 6.886 |
| 100 | empty | direct | 0.141 | 0.006 | 0.226 | 0.067 |
| 100 | empty | full | 3.984 | 3.863 | 4.822 | 4.269 |
| 100 | empty | full-fresh | 4.353 | 4.402 | 5.180 | 6.523 |
| 100 | ambiguous | direct | 0.109 | 0.005 | 0.177 | 0.011 |
| 100 | ambiguous | full | 3.970 | 3.602 | 5.136 | 4.140 |
| 100 | ambiguous | full-fresh | 4.641 | 4.336 | 5.235 | 6.074 |
| 100 | class | direct | 0.138 | 0.142 | 0.333 | 0.358 |
| 100 | class | full | 3.582 | 3.464 | 4.788 | 4.853 |
| 100 | class | full-fresh | 4.433 | 3.915 | 5.628 | 5.130 |
| 1000 | absent | direct | 1.259 | 0.051 | 1.930 | 0.066 |
| 1000 | absent | full | 35.733 | 34.541 | 37.599 | 36.256 |
| 1000 | absent | full-fresh | 43.430 | 39.459 | 64.208 | 54.655 |
| 1000 | empty | direct | 1.082 | 0.069 | 2.027 | 0.083 |
| 1000 | empty | full | 54.824 | 38.634 | 63.768 | 60.710 |
| 1000 | empty | full-fresh | 53.689 | 55.561 | 68.608 | 67.707 |
| 1000 | ambiguous | direct | 1.747 | 0.089 | 1.848 | 0.095 |
| 1000 | ambiguous | full | 58.341 | 56.877 | 62.562 | 60.490 |
| 1000 | ambiguous | full-fresh | 53.193 | 57.646 | 71.419 | 67.211 |
| 1000 | class | direct | 1.363 | 2.124 | 5.308 | 3.859 |
| 1000 | class | full | 34.476 | 42.166 | 52.177 | 65.174 |
| 1000 | class | full-fresh | 46.531 | 54.074 | 59.372 | 61.533 |
| 1 | 1001classes-invalid-offset | cold | 0.001 | 0.069 | 0.011 | 0.092 |
| 1 | 1001classes-invalid-offset | warm | 0.001 | 0.002 | 0.010 | 0.002 |

Isolated repeats use baseline (A), fixed (B), fixed (B), baseline (A) in separate sequential processes. All complete result hashes match across each four-run set. Each uses 1,000 references and the same 15 rounds/three warmups.

| Metadata | Run | Direct median ms | Full median ms | Fresh median ms | Direct p95 ms | Full p95 ms | Fresh p95 ms |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| class | A1 | 1.605 | 33.751 | 35.216 | 2.342 | 39.397 | 42.259 |
| class | B1 | 1.560 | 34.019 | 36.382 | 2.244 | 38.221 | 41.519 |
| class | B2 | 1.523 | 33.955 | 35.579 | 2.225 | 42.647 | 42.310 |
| class | A2 | 1.509 | 34.752 | 34.567 | 2.241 | 43.697 | 42.052 |
| empty | A1 | 1.143 | 38.677 | 41.139 | 1.883 | 46.521 | 49.211 |
| empty | B1 | 0.083 | 37.724 | 39.288 | 0.907 | 44.153 | 44.059 |
| empty | B2 | 0.084 | 38.883 | 41.548 | 1.113 | 45.255 | 48.973 |
| empty | A2 | 1.088 | 38.243 | 39.293 | 2.078 | 45.135 | 40.931 |

Reproduce full runs sequentially:

```powershell
node scripts/benchmark-project-member-empty-index.mjs --baseline=b8e44324
node scripts/benchmark-project-member-empty-index.mjs
```

For isolated repeats, add `--metadata=class --references=1000` or `--metadata=empty --references=1000`, and alternate baseline/fixed/fixed/baseline.
