# Sheets folder surface lookup

Expanding a workbook's Sheets folder used Array.find on the shape-surface list for every sheet. A 1,000-sheet public ShapeRows.children regression measured 500,500 surface-name reads per expansion, even with cached bridge results. The temporary index reduces that to 1,000, with no additional persistent cache or invalidation policy. The first exact-name match wins, including a match with no shapes.

The compatibility controls preserve tab order, duplicate and exact-case matching, missing/unlisted surfaces, empty module identity and parents, hidden/chart sheets, and shape-refresh movement between the main and bare Sheets folders. Existing real-file tree tests also run.

## Reproduce

Node v24.18.0; AMD Ryzen 7 9800X3D 8-Core Processor           . Baseline 693a4d2de415f72863bb60b314ef1cc5f8b9d445. Existing pinned esbuild builds the actual ShapeRows implementation, with a lightweight VS Code mock and an in-memory bridge. Only src/shapeRows.ts is replaced for the baseline. Both runs use frozen input arrays/entries, 15 measured rounds after three warm-ups, one cold expansion or 20 sequential cached expansions. Catalog loading occurs before the clock; cold rows include the first shape listing request. Assertions and row comparisons are outside the clock. Timing includes row construction and async orchestration, not Office I/O, VS Code rendering, or tab switching.

```powershell
node scripts/benchmark-tree-sheet-surface-lookup.mjs --baseline=693a4d2de415f72863bb60b314ef1cc5f8b9d445 > before.json
node scripts/benchmark-tree-sheet-surface-lookup.mjs > after.json
```

Every run verifies one catalog and one shape RPC, row labels/order and parents. The complete kind/label/item-count signatures for all 12 workloads matched before and after. Work-count tests use getters separately from timing. An additional baseline-versus-fixed probe compared 480 exact layout, rendered-item and parent snapshots across 120 generated workbooks and four phases (initial, cached, shape refresh, and clear). It includes mixed module/code states, missing and duplicate surfaces, exact-case mismatches, hidden and chart sheets. All snapshots matched; this covers the Sheets and bare Sheets rows, not full procedure trees or real VS Code UI.

## Results

All values are milliseconds for the stated expansion count. Small controls are mixed (including the small cold p95); these measurements do not establish a universal or UI latency improvement.

| Layout | Sheets | Expansions | Median before | Median after | p95 before | p95 after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| small (cold) | 5 | 1 | 0.014 | 0.013 | 0.029 | 0.557 |
| small (warm) | 5 | 20 | 0.171 | 0.167 | 0.466 | 0.297 |
| medium (cold) | 100 | 1 | 0.214 | 0.095 | 0.355 | 0.130 |
| medium (warm) | 100 | 20 | 3.702 | 1.298 | 4.093 | 1.888 |
| large (cold) | 1000 | 1 | 11.831 | 0.664 | 12.530 | 1.026 |
| large (warm) | 1000 | 20 | 241.060 | 13.110 | 362.038 | 15.004 |
| larger (cold) | 3000 | 1 | 102.854 | 2.094 | 120.004 | 4.701 |
| larger (warm) | 3000 | 20 | 2444.753 | 41.415 | 2887.494 | 48.486 |
| missing (cold) | 1000 | 1 | 32.655 | 0.806 | 37.967 | 1.187 |
| missing (warm) | 1000 | 20 | 570.728 | 12.921 | 650.862 | 18.567 |
| reversed (cold) | 1000 | 1 | 12.582 | 0.626 | 18.439 | 0.744 |
| reversed (warm) | 1000 | 20 | 314.040 | 12.241 | 387.503 | 12.896 |

The index is rebuilt on each Sheets folder expansion, requiring O(surface count) temporary references. Existing per-module or single-surface lookups are unchanged. No production dependency was added.

Validation passed: 68 focused tests including real workbook fixtures, type checks, and the full suite (636 files, 12,966 passed, 13 skipped). The new regression fails baseline with 500,500 name reads; its two semantic controls pass baseline and fixed.
