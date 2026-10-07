# Cached folder expansion lookup

ProjectExplorer caches each project's folder layout, but every folder child request walked its dotted path and linearly searched the sibling folders at each level. Reopening a late folder in a wide layout therefore repeated many comparisons despite having all data in memory.

Store canonical-path lookup entries in the same cached layout object as its tree. The existing folder-node registration walk fills the map; no second full traversal is added. Each drawn folder is then found directly. Both tree and lookup disappear together through existing refresh, edited-annotation and closed-editor invalidation. Stable folder/module nodes, parent links, render-version updates and sorting are unchanged. Exact canonical case and the old dotted-path normalization are preserved, including the no-segment path returning the root layout.

The index adds one map entry per folder referencing its existing metadata. It trades O(folder count) extra index storage and map insertion during registration for direct cached lookups. It does not cache module rows separately or change backend calls.

## Deterministic lookup cost

A scratch real-provider probe installs transparent getters on cached folder metadata after initial construction, then requests the last leaf folder's children 200 times and checks the returned module. This measures actual path comparisons without timing instrumentation:

| Fixture | Sibling width | Depth | Path reads before | Path reads after | listModules calls in each version |
| --- | ---: | ---: | ---: | ---: | ---: |
| small | 10 | 1 | 2,000 | 0 | 1 |
| wide | 1000 | 1 | 200,000 | 0 | 1 |
| wider | 3000 | 1 | 600,000 | 0 | 1 |
| deep-wide | 1000 | 8 | 201,400 | 0 | 1 |
| deep | 1 | 64 | 12,800 | 0 | 1 |

Zero means no cached metadata path reads for the measured leaf lookup, not zero provider work. Initial index construction and necessary mapping of child folders are outside those counters.

## Reproduction and timings

```powershell
node scripts/benchmark-folder-expansion.mjs --baseline=8c2cf8ed
node scripts/benchmark-folder-expansion.mjs
```

The benchmark bundles the actual provider and dependencies, reuses the repository's VS Code host stub with lightweight function wrappers, and supplies deterministic bridge responses for a VB6 project. No Office/backend filesystem I/O, real renderer, editor activation or tree reveal is measured. Each module has one leaf folder. Width is the number of sibling leaf folders; depth includes common ancestors. Layouts have 10/1,000/3,000 siblings, depth eight with 1,000 terminal siblings, or one leaf at depth 64.

Warm mode primes root/project children outside timing, then times 200 complete getChildren requests for the last leaf and verifies module answers afterward. Wide-first and wide-missing measure a first sibling and absent folder. Build mode times root discovery, module listing/sorting, layout construction and one leaf request for a new provider. It includes index allocation/registration and uses mocked immediate backend responses. Bridge/provider creation, result checks and disposal are outside timing. There are no counters in timing bundles. Three warmups precede 15 samples. Before/after runs were sequential before the full test suite, without concurrent audit workloads.

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor; measured locally 2026-10-03. Times are milliseconds per row's request batch.

| Fixture / mode | Width | Depth | Requests | Before median | After median | Before p95 | After p95 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| small-warm | 10 | 1 | 200 | 0.516 | 0.461 | 1.022 | 0.549 |
| small-build | 10 | 1 | 1 | 0.059 | 0.063 | 0.124 | 0.089 |
| wide-warm | 1000 | 1 | 200 | 1.419 | 0.497 | 2.407 | 1.169 |
| wide-build | 1000 | 1 | 1 | 3.904 | 3.862 | 4.552 | 5.481 |
| wider-warm | 3000 | 1 | 200 | 3.329 | 0.500 | 3.845 | 0.593 |
| wider-build | 3000 | 1 | 1 | 11.101 | 10.888 | 13.915 | 13.007 |
| deep-wide-warm | 1000 | 8 | 200 | 2.187 | 0.496 | 2.431 | 0.524 |
| deep-wide-build | 1000 | 8 | 1 | 3.781 | 3.774 | 6.499 | 5.703 |
| deep-warm | 1 | 64 | 200 | 2.011 | 0.491 | 3.962 | 0.514 |
| deep-build | 1 | 64 | 1 | 0.069 | 0.079 | 0.073 | 0.080 |
| wide-first-warm | 3000 | 1 | 200 | 0.524 | 0.494 | 0.730 | 0.526 |
| wide-missing-warm | 3000 | 1 | 200 | 3.263 | 0.514 | 3.341 | 0.559 |

Warm late/missing and deep-folder requests improve in these fixtures. Initial builds remain dominated by other work and show mixed changes: small median 0.059 to 0.063 ms, wide build p95 4.552 to 5.481 ms, and deep build median 0.069 to 0.079 ms. These synthetic provider measurements do not establish a universal UI latency improvement.

## Validation

Type check passed. Focused folder/provider/follow/layout/VB6 tests: 107 across five files; the separate counter probe also passed. Full suite: 629 files, 12,915 passed, 13 skipped.

Two added tests cover first-spelling paths, prior dotted-path normalization and exact-case misses, two-project isolation, annotation moves, removed folders and refreshes. Existing move/close/refresh/stable-node/follow cases remain. Native structured comparison across 200 generated layouts checks folder-layout row kind/label/path/module/count, render ID/state and parent identity through seven phases: initial, active module, annotation move, close/forget, refresh, flat view and folder view. All 1,400 snapshots and 1,000 noncanonical/missing/root path queries match the baseline. The comparison does not include Office shape rendering or every TreeItem property.

No dependencies or production counters added. Real VS Code integration was not run. The broader tree interaction and repository audit remains incomplete.
