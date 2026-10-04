# Tree module-list overlap performance

Expanding a project while tab-follow resolves a module coalesces the backend listModules call, but each awaiting caller used to copy and sort the same list separately. The pending load now owns the sorted copy. Existing per-caller generation guards still govern cache writes and node registration; bridge input remains untouched.

## Reproduction

Run sequentially, without other benchmarks/tests running:

```powershell
node scripts/benchmark-tree-module-overlap.mjs --baseline=921eb8d37f6e0716123f55feb6e11da09542789c
node scripts/benchmark-tree-module-overlap.mjs
```

The script bundles the actual provider with esbuild and the existing shared VS Code mock. Only baseline projectExplorer.ts is substituted. Project discovery and fixture construction happen outside the clock. Each cold sample starts one expansion and zero, one or seven tab-follow requests, lets follow traverse the root, then releases the shared deferred RPC. Eight callers are a stress case. Warm controls preload the tree outside the clock. Ordering, frozen input, shared module-node identity and exactly one listModules call are asserted outside timing. Sorting/reference output preparation also occurs outside timing; comparator counters run only in separate tests. No additional dependencies.

Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor; 15 measured rounds after three warm-ups per row. Times are milliseconds for one sample, not per module. Before and after ran sequentially.

| Modules | Cache | Callers | Before median | After median | Before p95 | After p95 |
| ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 50 | cold | 1 | 0.078 | 0.080 | 0.143 | 0.132 |
| 50 | cold | 2 | 0.111 | 0.099 | 0.171 | 0.124 |
| 50 | cold | 8 | 0.361 | 0.293 | 0.773 | 0.590 |
| 50 | warm | 2 | 0.035 | 0.035 | 0.064 | 0.057 |
| 1000 | cold | 1 | 0.686 | 0.691 | 0.964 | 0.866 |
| 1000 | cold | 2 | 1.545 | 1.138 | 2.156 | 1.583 |
| 1000 | cold | 8 | 5.665 | 3.577 | 7.506 | 4.209 |
| 1000 | warm | 2 | 0.519 | 0.442 | 0.600 | 0.776 |
| 3000 | cold | 1 | 2.130 | 2.513 | 2.513 | 3.501 |
| 3000 | cold | 2 | 4.589 | 3.485 | 5.745 | 3.909 |
| 3000 | cold | 8 | 17.995 | 10.580 | 20.109 | 12.164 |
| 3000 | warm | 2 | 1.331 | 1.388 | 1.765 | 1.661 |

The deterministic provider regression counts 8,741 comparisons for one caller and 17,482 for expansion plus one follow on the baseline, despite one backend request. Both now share the single sort. Larger overlap rows improve, while controls are mixed: notably the 3,000-module single-caller median rises from 2.130 to 2.513 ms. These samples do not establish a universal latency improvement. Node mapping and each caller's derived child layout remain per caller, explaining residual overlap work. No additional persistent cache or index is introduced.

Real VS Code rendering, tree animations, Office bridge latency and end-to-end interaction were not measured.
