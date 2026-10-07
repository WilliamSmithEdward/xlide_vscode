# Unrelated tab closure performance

Closing non-module tabs previously scanned all open tab groups and URI strings. Resolve closing module candidates first and read the open-tab provider only when needed. Array callers remain supported. Module closures preserve URI matching and first module-identity deduplication.

Three input-read regressions failed before the repair; ten existing controls passed. The fixed suite has 82 focused tests passing, type checking passing, and 640 files / 13,000 tests passing / 13 skipped in the full suite. A real-decoder native differential compared 1,024 complete array and lazy-provider outputs to the baseline across text, custom, diff, mixed and malformed fixtures.

Benchmark: actual ExplorerFollow tab callback, extracted production extension callback and tab helper; deterministic module ownership stub; mocked VS Code host. This excludes real ownership decoding, Office and rendered UI latency. Counts are measured separately from timing; assertions are outside the clock. Node v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor           , baseline 1c34e8f1, 15 rounds after three warmups, 200 events per row. Genuine module closures retain all URI reads and fold outcomes; their timings are mixed.

| Event | Before URI reads | After URI reads | Before median ms | After median ms | Before p95 ms | After p95 ms |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| changed-1 | 0 | 0 | 0.006 | 0.004 | 0.035 | 0.026 |
| changed-50 | 0 | 0 | 0.002 | 0.002 | 0.002 | 0.005 |
| changed-1000 | 0 | 0 | 0.002 | 0.002 | 0.034 | 0.023 |
| opened-1 | 0 | 0 | 0.002 | 0.002 | 0.002 | 0.002 |
| opened-50 | 0 | 0 | 0.002 | 0.002 | 0.008 | 0.008 |
| opened-1000 | 0 | 0 | 0 | 0 | 0.002 | 0.002 |
| empty-1 | 0 | 0 | 0 | 0 | 0 | 0.001 |
| empty-50 | 0 | 0 | 0 | 0 | 0.001 | 0.001 |
| empty-1000 | 0 | 0 | 0 | 0.001 | 0.001 | 0.001 |
| nonmodule-close-1 | 200 | 0 | 0.109 | 0.028 | 0.131 | 0.091 |
| nonmodule-close-50 | 10000 | 0 | 1.1 | 0.028 | 1.729 | 0.034 |
| nonmodule-close-1000 | 200000 | 0 | 19.479 | 0.047 | 20.274 | 0.065 |
| module-close-1 | 200 | 200 | 0.194 | 0.249 | 0.264 | 0.31 |
| module-close-50 | 10000 | 10000 | 1.165 | 1.113 | 2.009 | 1.484 |
| module-close-1000 | 200000 | 200000 | 19.477 | 19.53 | 21.418 | 21.703 |

Reproduce sequentially:

```powershell
node scripts/benchmark-tree-tab-events.mjs --baseline=1c34e8f1
node scripts/benchmark-tree-tab-events.mjs
```
