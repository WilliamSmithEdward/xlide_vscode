# Procedure row lookup performance

Tree follow resolved a cached procedure with a case-insensitive array scan on every pass. The row cache now owns a first-label index alongside its stable nodes, so module and whole-tree refresh discard both together. Designer rows remain excluded and the first case-insensitive duplicate still wins.

A regression fixture recorded 200,000 label reads for 200 last-procedure lookups in a 1,000-procedure module before the change, and zero afterward. Three behavior controls passed before and after; focused tree, tab, and refresh tests total 123 passing. Type checking passed. Full suite: 641 files, 12,999 tests passed, 13 skipped.

Measured with Node v24.18.0 on AMD Ryzen 7 9800X3D 8-Core Processor           , baseline 1c34e8f1. Each row uses 15 rounds after three warmups. The benchmark calls the actual ProjectExplorer with mocked bridge and VS Code; assertions compare every lookup to the original case-insensitive first-match rule outside the timer. Frozen backend input is reused, one module listing and one procedure listing are required per provider. Cold is one procedure expansion; other rows are 200 awaited follow lookups. This measures provider work, not rendered UI, Office, or debounce latency.

The index adds linear construction and memory per cached module. Large last/missing/jumping lookups improve; early hits and small modules change little. Cold expansion is slower for larger lists because the index is built once. No persistent project state or extra bridge calls are introduced.

| Procedures | Mode | Before median ms | After median ms | Before p95 ms | After p95 ms |
| --- | --- | ---: | ---: | ---: | ---: |
| 5 | cold | 0.009 | 0.008 | 0.035 | 0.035 |
| 5 | first | 0.414 | 0.407 | 0.549 | 0.632 |
| 5 | last | 0.300 | 0.293 | 0.494 | 0.381 |
| 5 | missing | 0.309 | 0.296 | 0.607 | 0.374 |
| 5 | jumps | 0.289 | 0.283 | 0.653 | 0.388 |
| 100 | cold | 0.008 | 0.015 | 0.019 | 0.028 |
| 100 | first | 0.274 | 0.270 | 0.387 | 0.342 |
| 100 | last | 0.414 | 0.270 | 0.598 | 0.326 |
| 100 | missing | 0.403 | 0.266 | 0.509 | 0.359 |
| 100 | jumps | 0.345 | 0.277 | 0.447 | 0.359 |
| 1000 | cold | 0.053 | 0.101 | 0.188 | 0.237 |
| 1000 | first | 0.312 | 0.272 | 0.431 | 0.346 |
| 1000 | last | 1.918 | 0.270 | 2.462 | 0.462 |
| 1000 | missing | 1.663 | 0.268 | 2.415 | 0.411 |
| 1000 | jumps | 1.032 | 0.312 | 1.457 | 0.414 |
| 3000 | cold | 0.166 | 0.333 | 0.988 | 0.502 |
| 3000 | first | 0.320 | 0.270 | 0.508 | 0.512 |
| 3000 | last | 4.988 | 0.277 | 6.076 | 0.325 |
| 3000 | missing | 4.488 | 0.270 | 5.956 | 0.456 |
| 3000 | jumps | 2.450 | 0.279 | 3.946 | 0.447 |

Reproduce sequentially from the repository root:

```powershell
node scripts/benchmark-tree-procedure-lookup.mjs --baseline=1c34e8f1
node scripts/benchmark-tree-procedure-lookup.mjs
```
