# Folder pending-review badge performance

When any module in a project awaits agent review, every folder tree item scanned all loaded module nodes, normalized each project path and checked pending state. The tree already has a normalized project/module identity map, and pending reviews already enumerate the project modules awaiting review. Folder badges now look up only those pending modules. No additional cache, invalidation rule or dependency is introduced. The pending list is still read on every render, so Keep/Revert and editor folder moves show immediately. Enumerating pending identities still costs time proportional to global pending reviews.

The single-use _findModuleNode fallback was also redundant: it scanned loaded rows after a normalized-key miss, but compared exactly the same normalized project path and lowercase module name used by the key. Its caller now uses getModuleNode directly. An unknown-module refresh drops from 1,000 loaded module path reads to zero; existing differently-cased refresh/name tests remain passing.

## Reproduction

```powershell
node scripts/benchmark-tree-folder-review.mjs --baseline=0d769761dc7dd11a0be7383752d680425e9a2088
node scripts/benchmark-tree-folder-review.mjs
```

Actual ProjectExplorer bundled with esbuild and the existing shared VS Code mock; only projectExplorer.ts is substituted for the baseline. Real presentAgentModuleWrite/keepAgentChange register/clear reviews outside the clock; the mock diff command does not open Office or UI. Project discovery, module/folder loading, review registration and cleanup are outside timing. Each sample measures 200 getTreeItem calls, cycling over the first folders; single pending review is either the first or last module, or all/none await review. Case-normalized review names are used. Expected description/icon colour/no decoration URI are asserted after timing. No getter/comparator instrumentation is used in timed samples.

v24.18.0, AMD Ryzen 7 9800X3D 8-Core Processor; 15 measured rounds after three warm-ups per row. Baseline and changed runs are sequential, without other audit tests/benchmarks running. Milliseconds per 200-render sample.

| Modules | Pending | Before median | After median | Before p95 | After p95 |
| ---: | --- | ---: | ---: | ---: | ---: |
| 50 | none | 0.292 | 0.279 | 0.453 | 0.500 |
| 50 | first | 5.546 | 0.462 | 6.563 | 0.494 |
| 50 | last | 5.644 | 0.445 | 6.215 | 0.828 |
| 50 | all | 7.815 | 6.317 | 8.643 | 8.457 |
| 1000 | none | 0.162 | 0.155 | 0.322 | 0.161 |
| 1000 | first | 99.201 | 0.419 | 107.191 | 1.431 |
| 1000 | last | 101.905 | 0.461 | 112.852 | 0.737 |
| 1000 | all | 61.957 | 52.870 | 63.666 | 57.343 |
| 3000 | none | 0.151 | 0.149 | 4.050 | 0.197 |
| 3000 | last | 301.816 | 0.411 | 313.950 | 5.955 |

Separate actual-provider probes count loaded module path reads: with 1,000 modules, one pending review and 200 unrelated folder renders, 400,200 -> 200. Unknown-module agent refresh: 1,000 -> zero. The no-review path reads no module paths in either version. Five regressions cover these work counts plus project isolation/unknown reviews and nested folder changes/Keep with stable render IDs. Two work-count tests fail on the baseline; three compatibility controls pass there.

No-review controls are essentially unchanged, with mixed p95 noise (50-module p95 0.453 -> 0.500 ms). All-pending cases still enumerate many review identities and candidate folder chains; this change does not make that work constant time. These measurements are provider CPU work, not VS Code rendering or Office latency.
