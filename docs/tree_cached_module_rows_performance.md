# Cached module-row expansion performance

A cached project expansion rebuilt the module-row array from all backend metadata on every render, even though individual node identities were already kept. The module-list cache now holds those rows. Overlapping cold expansion/follow callers also reuse the first caller's constructed rows. This replaces retained metadata with references to the existing nodes, rather than adding another cache. Existing refresh, retry and editor-close invalidations remain in place. Live folder edits update the node objects the cache already holds.

Because old loads no longer get a fresh metadata pass on every cache hit, an obsolete load must not mutate a node owned by the new render. It now builds unregistered temporary nodes when its generation has changed. Generation checks still prevent stale cache writes; the project render joins the current generation before returning derived rows.

## Reproduction

```powershell
node scripts/benchmark-tree-module-overlap.mjs --baseline=3378fd2c --warm-expansions=200
node scripts/benchmark-tree-module-overlap.mjs --warm-expansions=200
```

The existing benchmark bundles the actual provider and the shared VS Code mock; only projectExplorer.ts is substituted for the baseline. Default warm-expansions remains one, preserving the previous overlap benchmark command. Root discovery and fixture creation are outside timing. Cold rows have one project expansion and zero, one or seven follow calls. Warm rows preload outside timing, then measure 200 sequential project expansions and one cached follow. Sorting, output-order validation, frozen-input verification and shared-node checks run outside the clock; one backend call is asserted for every sample. No comparator/getter instrumentation runs in timing samples. Eight overlapping callers are a stress case.

v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor; 15 rounds after three warm-ups, sequential before/after runs, no other audit tests running. Times are milliseconds per sample. Cold and warm rows have different numbers of expansions.

| Modules | Cache | Callers | Expansions | Before median | After median | Before p95 | After p95 |
| ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 50 | cold | 1 | 1 | 0.083 | 0.076 | 0.142 | 0.141 |
| 50 | cold | 2 | 1 | 0.100 | 0.061 | 0.145 | 0.068 |
| 50 | cold | 8 | 1 | 0.292 | 0.069 | 0.381 | 0.104 |
| 50 | warm | 2 | 200 | 4.613 | 0.919 | 6.136 | 1.092 |
| 1000 | cold | 1 | 1 | 0.650 | 0.773 | 1.018 | 1.273 |
| 1000 | cold | 2 | 1 | 1.173 | 0.676 | 1.818 | 0.828 |
| 1000 | cold | 8 | 1 | 3.532 | 0.752 | 4.242 | 0.925 |
| 1000 | warm | 2 | 200 | 77.434 | 1.813 | 82.493 | 2.892 |
| 3000 | cold | 1 | 1 | 2.134 | 2.162 | 2.800 | 2.473 |
| 3000 | cold | 2 | 1 | 3.498 | 2.178 | 4.108 | 2.467 |
| 3000 | cold | 8 | 1 | 10.741 | 2.284 | 11.914 | 2.983 |
| 3000 | warm | 2 | 200 | 230.814 | 3.715 | 239.820 | 7.184 |

A separate actual-provider getter probe counts 200,000 backend name reads for 200 cached expansions of 1,000 modules, reduced to zero. Cold two/eight-caller probes also show only one metadata-to-row mapping after the shared sort. Five regressions check these three reductions plus live folder edit/close/refresh and an obsolete same-module result arriving after the current result. On the baseline, the three work-count regressions fail and the two compatibility controls pass.

Single-caller controls are mixed: the 1,000-module cold median is 0.650 -> 0.773 ms. This is provider CPU measurement, not VS Code renderer, animation or Office bridge latency. Remaining shape/sheet layout work and output array copies are still performed per render.

A generated folder parity probe compared 200 layouts across seven phases (1,400 snapshots) plus 1,000 unusual folder-path queries. Row kind/label/folder/module/count, render ID/state and parent identity matched the baseline through active-module changes, editor folder edits, editor close, refresh and both view layouts. This scope excludes procedure rows and real Office/VS Code rendering.
