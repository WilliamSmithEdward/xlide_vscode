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
| Sidebar setup-health refresh without side effects | 500 ms p95 | 2 s |
| Import/export diff preview generation | 2 s p95 | 10 s |
| Test GUI discovery before Excel execution | 1 s p95 | 5 s |

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
  expansion, commands, backend RPCs, virtual file reads/writes, workbook
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
