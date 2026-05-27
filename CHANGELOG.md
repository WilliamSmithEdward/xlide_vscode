# Changelog

All notable changes to **XLIDE: VBA for VS Code** are documented here.

## [1.0.8] - 2026-05-26

### Fixed
- **Tree view sync after rename/edit** — Renaming a procedure via Rename Symbol (F2) or by editing the source manually now refreshes the affected module's sub list in the XLIDE Explorer on save, instead of showing the old name until a full refresh.

## [1.0.7] - 2026-05-26

### Changed
- Unified the XLIDE Explorer welcome message: a single entry shown in all sessions (host and Live Share guest) with the workbook hint, Refresh link, and brief Live Share notes.

## [Unreleased]

### Added
- **Module tree accordion** — clicking a module tab or tree node auto-expands that module's procedure list and collapses all others. Closes last tab to collapse everything.
- **Accordion debounce** — rapid Ctrl+W cycling coalesces into a single tree update (60 ms), eliminating race-condition stragglers.
- **Add Class Module** command and context menu item with `Class_Initialize` / `Class_Terminate` stub.
- **Module type detection overhaul** — UserForms identified by two-GUID `VB_Base` pattern (works for both live workbook and exported `.cls` files). Document CLSIDs for Workbook / Worksheet / Chart correctly classified.
- **COM window focus** — after opening a workbook or running a macro on Windows, Excel is restored and brought to the foreground via P/Invoke (`ShowWindow` + `SetForegroundWindow`).
- **Symbol kinds** — Outline / breadcrumbs now show Const → Constant, Enum → Enum, Type → Struct icons in addition to Sub / Function / Property.
- **listModules cache** — module list cached per workbook during a session; cleared on tree refresh to avoid stale data.
- **Filesystem watcher debounce** — rapid save-storm events coalesce into a single explorer refresh (200 ms).
- **Cancellation tokens** — `bridge.call()` accepts an optional `CancellationToken`; pending RPC requests are rejected on cancellation.
- **Smoke test command** (`XLIDE: Run Smoke Test`) — verifies listModules and readModule against a workspace workbook from the command palette.
- **TS unit tests** — vitest suite covering `parseVbaModule` (Sub/Function/Property/Const/Enum/Type, line spans, visibility) and `decodeModuleUri` (module name decode, URL encoding, extension variants, error cases).
- **Python unit tests** — pytest suite covering `_split_vba_source` (round-trip, VERSION/BEGIN/END stripping) and `_module_type` (userform two-GUID, document CLSIDs, name heuristics, PredeclaredId).
- **CI workflow** — GitHub Actions runs `npm run compile` + `npm test` (TypeScript) and `pytest` (Python) on push and pull requests.
- **Export path fix** — class modules now export as `.cls`; document modules as `.cls`; userforms as `.frm`; standard modules as `.bas`.
- **Import UX** — `.frm` / document / userform files that don't exist in the live workbook are shown with an explanatory detail in the QuickPick and cannot be selected.
- **Live Share** — module type surfaced for remote modules; userform icon and sort order applied consistently.
- **Status bar** — shows active workbook / module name; Live Share guest count.
- **VBA snippets** — 21 snippet entries (`sub`, `func`, `for`, `forEach`, `with`, `select`, `class`, `prop`, …).
- **onEnterRules** — auto-insert `End Sub` / `End Function` / `End If` / `Next` / `Loop` / `Wend` on Enter, matching VBE behaviour.
- **Bridge auto-restart** — unexpected Python child-process exit marks the bridge stopped; next call shows a clear actionable error.
- **Workbook-locked UX** — WinError 32 / sharing violation detected and surfaced as a warning with a Retry action.
- **Marketplace display name** updated to `XLIDE: VBA for VS Code`.

### Changed
- Context menu reorganised into logical groups: create, edit, workbook, transfer, settings.
- `xlide.newClassModule` replaces the generic new-module path for class modules.

---

*XLIDE follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) conventions.*
