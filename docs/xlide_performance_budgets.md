# XLIDE Performance Budgets

Purpose: keep XLIDE responsive while editing real VBA workbooks, and make slow
paths measurable before v2 launch.

These budgets are launch targets, not proof that every path already meets them.
When a path exceeds the budget, XLIDE should either make the work incremental,
cancelable, cached, or visibly reported with progress.

## Interaction Budgets

| Path | Target | Hard ceiling before progress/cancellation |
|---|---:|---:|
| Keystroke-triggered editor diagnostics for the active module | 150 ms p95 | 500 ms |
| Keyword/type/member completion resolution after a trigger character | 100 ms p95 | 250 ms |
| Hover/signature-help resolution in the active module | 100 ms p95 | 250 ms |
| Current-module analysis command | 500 ms p95 | 2 s |
| Workbook-wide analysis command | 5 s p95 for ordinary workbooks | 10 s |
| Project index rebuild after one module save | 500 ms p95 | 2 s |
| Workbook tree refresh for ordinary workspace folders | 1 s p95 | 5 s |
| Import/export diff preview generation | 2 s p95 | 10 s |
| Test GUI discovery before Excel execution | 1 s p95 | 5 s |

## Keystroke-Path Rules

Measured on a real 24,283-line class module (867 KB of source):

| In-host cost per event | Path |
|---:|---|
| ~0.01 ms | loop-iterator sync early-out (per keystroke) |
| 4.4 ms | project-index fold (`buildModuleSymbols`, per provider request) |
| 29 ms | structural analysis alone |
| **~700-945 ms** | `analyzeVbaModuleSource` - a full module analysis |

That last number is why BOTH diagnostic passes ride the analysis worker when it
is healthy: an in-host "local" pass is a full module analysis, and firing it
90 ms into every typing pause froze typing on large modules. Rules that keep
this fixed:

- Never run `analyzeVbaModuleSource` on the extension host on a keystroke-paced
  timer. When the worker is up, both the local and full passes go to it; the
  worker keeps per-document incremental state so the follow-up pass re-analyzes
  only what changed.
- When the worker is down, large workbook modules drop the local pass entirely
  (the paced full pass covers them) and loose `.bas` files pace their only pass
  to the same backoff. `tests/vbaDiagnosticScheduling.test.ts` pins all of this.
- Analyze Workbook rides the same worker: each module's analysis runs
  off-thread (own seed namespace, content-fingerprint generation so an
  unchanged re-run skips the seed transfer and reuses per-module incremental
  state), with the identical in-host pass as fallback.

## Workbook Engine Budgets

Every workbook operation runs in-process against the file on disk, so these are
whole-call costs with no IPC or process startup in them. Measured on a 2.4 MB
workbook with 42 modules (621 KB of VBA source), which is past the "medium"
fixture below:

| Operation | Target | Measured p50 |
|---|---:|---:|
| `listModules`, `listSubs`, `getProtectionInfo`, `validateWorkbook` | 25 ms | ~4 ms |
| `readModule` (one module) | 25 ms | ~4 ms |
| `readModules` (every module, the analysis path) | 100 ms | ~12 ms |
| `listSheets`, `readCells`, `readFormulas` | 25 ms | 1-3 ms |
| `writeModule` on an existing module (the Ctrl+S path) | 150 ms | ~17 ms |
| `writeModule` creating a module | 150 ms | ~37 ms |
| `writeCells` | 100 ms | ~4 ms |
| `createWorkbook` | 100 ms | <1 ms |

Warm calls are two orders of magnitude cheaper again: reads against a workbook
whose parse is already cached cost ~0.02 ms, so a burst like an explorer
expansion (listModules + a protection probe + one listSubs per module) pays for
exactly one parse. On a 42-module workbook that is 44 calls for ~22 ms total.

Three properties keep those numbers where they are, and all are easy to undo by
accident:

- **Module sources decompress lazily.** Parsing a project inflates only the
  module bodies a caller actually reads; classification reads a header prefix
  (`VbaModule.sourceHeader`). Touching `module.source` in a path that only needs
  names or types silently doubles every read in the extension.
- **Reads share one parse per workbook.** `workbookService` caches the parsed
  package/project per path, validated against (mtimeMs, size) on every call and
  dropped by `atomicWrite`. Mutating operations must keep using the fresh-parse
  path (`openWorkbookForWrite`) - handing a writer the shared parse poisons it
  for every reader if the save fails halfway.
- **Rewritten ZIP entries deflate at level 4, not 6.** Re-deflating
  `vbaProject.bin` dominates a save. Level 6 costs roughly 80% more time for an
  entry ~2.5% smaller.

## Scale Assumptions

Use these as routine development fixtures and manual smoke targets:

- Small workbook: 1-10 modules, under 5,000 total VBA lines.
- Medium workbook: 10-50 modules, under 50,000 total VBA lines.
- Large workbook: 50-250 modules, under 250,000 total VBA lines.
- Large single module: 10,000 lines.
- Diagnostic-heavy module: at least 500 findings without UI lockup.

If a workbook exceeds these assumptions, XLIDE should still avoid blocking the
UI thread and should report long-running work clearly.

## Measurement Rules

- Measure analyzer paths outside VS Code first with deterministic unit or
  integration fixtures.
- Measure VS Code provider paths separately when the UI layer is the suspected
  bottleneck.
- Use `XLIDE: Copy Performance Snapshot` after reproducing latency in the
  editor. The snapshot includes recent timings for completion, hover, signature
  help, semantic tokens, live diagnostics, workbook context indexing, tree
  expansion, commands, workbook engine calls, virtual file reads/writes, workbook
  analysis, module sync/export, documentation metadata, and VBA test stages.
- Set `xlide.performance.trace` to `true` only while debugging latency. It writes
  slow trace events to the XLIDE output channel; the snapshot command remains
  available without enabling verbose output logging.
- Keep Excel COM execution out of routine performance tests. Excel-hosted test
  and macro runs have their own host trace timings.
- Use p95 targets to avoid optimizing only the smallest examples.
- Preserve correctness before speed. Performance work must not silently skip
  diagnostics, project symbols, or workbook modules.

## Required Hardening Before Broad Release

- Add at least one large-workbook fixture or synthetic project generator that
  exercises parser, symbol graph, diagnostics, and completion setup.
- Add budget assertions for pure analyzer paths where results are stable across
  machines.
- Add cancellation or explicit progress for workbook-wide analysis, import/export
  previews, and test discovery if they exceed the hard ceiling.
- Keep live diagnostics debounced and cancel stale work after newer edits.
- Invalidate caches by source text, workbook identity, module identity, settings,
  and metadata version.

## Reporting

When a command takes longer than its hard ceiling, prefer visible progress or a
clear status row over writing only to the Output channel. Support bundles should
include recent command timing and host/test trace summaries where available.

Important trace names:

- `completion`, `hover`, `signatureHelp`
- `semanticTokens`
- `liveDiagnostics.local`, `liveDiagnostics.full`
- `tree.getChildren`
- `workbookContext.getAllModules`
- `analyzeWorkbook.readModules`, `analyzeWorkbook.listModules`,
  `analyzeWorkbook.readModule`
- `analyzeWorkbook.buildProjectContext`, `analyzeWorkbook.settings`,
  `analyzeWorkbook.analyzeModule`, `analyzeWorkbook.total`
- `command`
- `filesystem.readFile`, `filesystem.writeFile`
- `docs.reload`
- `sidebar.render`, `sidebar.workbookFiles`
- `moduleExport.single`, `moduleExport.workbook`
- `moduleSync.buildExportPlan`, `moduleSync.buildImportPlan`,
  `moduleSync.renderHtml`, `moduleSync.refreshFromDisk`,
  `moduleSync.refreshWorkbookSettings`, `moduleSync.refreshSettings`,
  `moduleSync.chooseFolder`, `moduleSync.apply`
- `workbookAnalysis.renderPanel`
- `vbaTests.discoverWorkbook`, `vbaTests.runWorkbook`,
  `vbaTests.ownedExcelHost`, `vbaTests.renderPanel`
