# Changelog

All notable changes to **XLIDE: VBA for VS Code** are documented here.

## [2.1.0] - 2026-06-09

### Added

- Added module-qualified procedure, function, constant, enum, and exported
  global resolution across diagnostics, completion, hover, signature help,
  rename, and Option Explicit checks.
- Added module-level XML documentation comment support so module summaries and
  member comments flow through IntelliSense surfaces.
- Added `XLIDE: Copy Performance Snapshot` and optional
  `xlide.performance.trace` output logging so editor, analysis, backend,
  virtual filesystem, sidebar, sync, documentation metadata, and VBA test
  latency can be diagnosed from one recent timing buffer.
- Expanded performance trace coverage across Python bridge startup/RPC calls,
  `xlide-vba://` reads/writes, workbook analysis stages, sidebar workbook
  discovery/rendering, module sync/export previews, analysis/test result
  webviews, documentation metadata reloads, and owned Excel test execution.
- Promoted a large wave of Excel host metadata for completion, hover, signature
  help, and receiver-chain inference, including WorksheetFunction, Pivot
  objects, QueryTables, chart internals, ShapeRange, comments, sort/filter
  helpers, form controls/OLEObjects, workbook connections, slicers/timelines,
  shape and chart formatting internals, conditional formatting subtypes,
  legacy drawing objects, sparklines, XML maps, publish objects, and web options.

### Changed

- Refocused and closed the Version 2.1.0 roadmap as the completed
  red-squiggle completeness sprint, with parser, binder, diagnostics, host
  metadata, rename, IntelliSense, and performance work prioritized by
  developer-experience impact.
- Moved all remaining open 2.1.0 backlog into `docs/roadmap_version_2.2.0.md`,
  with object member/event authoring continuation as the new 2.2 Priority 1.
- Aligned normal module rename behavior with the shared module-rename strategy
  so module-qualified references update consistently outside the class-module
  tree path.
- Moved user-facing documentation comment guidance into the user guides area and
  recorded internal oracle discipline for sequential, non-parallel runs.
- Improved workbook context and analysis responsiveness with bounded parallel
  module analysis, more explicit trace stages, cancellation/supersession
  handling, and cache pruning for editor project context.
- Changed XLIDE VBA editor defaults to use 4-space indentation and a quieter
  minimap/overview-ruler profile for large generated or workbook-backed modules.

### Fixed

- Hardened shared expression binding and name resolution for bare and
  module-qualified identifiers, procedure return variables, exported globals,
  constants, enum members, arrays, `With` receivers, source-shadowed runtime
  names, and host globals.
- Reduced false positives from live workbook testing around Option Explicit,
  module-qualified calls/reads, Attribute metadata placement, function return
  checks, array-return assignments, inactive `#If` branches, `VBA7`/platform
  defaults, parser-recovered procedure headers, and late-bound object/Variant
  receiver behavior.
- Expanded high-confidence red diagnostics for declaration order, parameter and
  property shapes, duplicate labels, enum ambiguity, conditional branch order,
  `For`/`Next` variable mismatches, `For Each` control/source typing,
  `ReDim`/`Erase`/`LBound`/`UBound` array misuse, deterministic runtime
  conversions, `Null`, `CVErr`, scalar/object `Set` usage, ByRef binding, and
  straight-line object-variable-not-set cases.
- Fixed source-shadowed runtime names, intrinsic/host-global resolution pressure,
  `With` receiver lookup, enum member ambiguity, and module-qualified exported
  member lookup in red-squiggle diagnostics and IntelliSense.

## [2.0.2] - 2026-06-04

### Fixed

- Fixed workbook analysis false positives in common line-numbered VBA, including
  `On Error GoTo 0`, `On Error GoTo -1`, `Erl`, and line-numbered
  `Select Case`/`Case` blocks.
- Hardened diagnostics and call analysis around numeric line labels for
  assignments, `Set` assignments, `Const` writes, procedure calls, argument
  counts, and mismatched `Exit` statements.
- Improved workbook analysis result navigation by queueing rapid row clicks and
  centering the selected finding in the opened module.

## [2.0.1] - 2026-06-04

### Changed

- Removed contribution links from the in-extension XLIDE sidebar so the product
  shell ends with Support actions.
- Moved open-source support links to the bottom of the repository README.
- Tightened VSIX packaging excludes so workbook files and test workbook folders
  stay out of Marketplace packages.

## [2.0.0] - 2026-06-04

Version 2 turns XLIDE from a workbook/module bridge into a fuller VBA
development environment for VS Code: project-aware editing, deterministic
analysis, previewable sync, workbook tests, safer workbook mutation, and
agent-verifiable workflows.

### Added

- **XLIDE Activity Bar/sidebar** - dedicated setup health, selected-workbook
  actions, global settings, support commands, and a stable workbook-action
  surface while keeping workbook/module navigation in the VS Code Explorer.
- **Workbook-wide VBA analysis** - `XLIDE: Analyze Workbook` opens a dedicated
  results UI with module grouping, severity filters, counts, copy/export
  actions, workbook/global rule settings, tracking controls, and click-through
  navigation.
- **Agent analysis verification** - `xlide_analyzeWorkbook` returns the same
  workbook analysis shape used by the UI:
  `{filePath, moduleCount, errorCount, warningCount, problems: [...]}`.
- **Deterministic diagnostic metadata** - rules now carry category,
  VBE-compile-equivalence, diagnostic kind, stable codes, and source labels so
  red diagnostics stay reserved for proven compile/runtime errors.
- **Expanded VBA diagnostics** - v2 adds and hardens high-confidence checks for
  block balance, unterminated strings, duplicate declarations, `Const`
  assignment, invalid procedure headers, unbalanced parentheses, declaration and
  call hygiene, argument counts and type mismatches, unknown calls, invalid
  declaration names, scalar member access, object/scalar `Set` usage, return
  assignment warnings, test marker syntax, and source-backed `member-not-found`.
- **Project-aware language service** - completions, hover, signature help,
  go-to-definition, find references, rename, semantic type coloring, and code
  actions now use the shared analyzer/project model where XLIDE can bind the
  target deterministically.
- **Syntax and editor hardening** - Smart Enter, block snippets, close-keyword
  completions, keyword casing, comment continuation, paired `For`/`Next`
  iterator edits, parser recovery, and MS-VBAL-backed keyword/token handling
  were expanded for common VBA editing flows.
- **Source-backed object/member understanding** - project classes, public
  fields, properties, return-name assignments, UDT fields, document/UserForm
  code names, known runtime signatures, and the first generated Excel host
  surfaces participate in completion, hover, navigation, diagnostics, and
  assignment validation where the surface is proven.
- **Documentation metadata** - inline `'''` XML doc comments and external
  `.vbref.xml` metadata enrich hover, completion, and signature help for
  source-backed and externally documented symbols.
- **Generated host/reference metadata slice** - generated Excel reference
  metadata is now used for the proven Excel `Workbook` surface, with coverage
  and provenance tracked separately from runtime extension code.
- **Workbook VBA test runner** - marked `@xlide-test` procedures run through an
  XLIDE-owned read-only Excel host, with `XlideAssert.bas` support, skip/xfail
  metadata, tag filters, current-module/current-test commands, rerun flows,
  output capture, artifacts, and `status_for_ci.json`.
- **Agent test execution** - `xlide_runVbaTests` runs the same workbook tests
  headlessly and returns `{ok, summary, artifacts, report}` for AI-agent and CI
  verification.
- **Previewable import/export sync** - bulk export/import and current-module
  export use diff previews, explicit apply steps, changed/skipped/removed/failed
  summaries, safe true-up behavior, and one workbook settings owner.
- **Workbook and global settings surfaces** - `xlide.openGlobalSettings`,
  workbook-facing analysis/sync settings, guarded severity overrides,
  untracked-rule controls, strict sidecar parsing, and provenance labels make
  settings explicit without weakening diagnostic determinism.
- **Code actions** - deterministic quick fixes/source actions cover the shipped
  safe edit set, including adding `Option Explicit`, fixing known call syntax,
  adding/removing `Set`, inserting block closers, moving misplaced `Option`
  statements, splitting local `Dim` initializers, adding suppression comments,
  creating safe private stubs, analyzing the current module, and exporting the
  current module.
- **Safety, trust, and recovery** - dirty `xlide-vba` backup restoration,
  explicit mutation/run commands, write audit summaries, workbook lock/open-state
  checks, COM timeout/cleanup handling, support bundles, copy diagnostics, and
  redacted support output make workbook operations easier to trust.
- **Performance budgets** - v2 records launch-facing latency budgets for
  keystroke diagnostics, module analysis, workbook analysis, project index
  rebuilds, and sidebar health refreshes; deeper performance hardening moved to
  the v2.1.0 roadmap.
- **User-facing documentation** - README and user guides now cover setup,
  workbook workflows, analysis/ignores, import/export sync, testing,
  automation/CI, safety/support, and the v2 feature surface.

### Changed

- Replaced the older dedicated `xlide.diagnostics.optionExplicit` model with
  guarded global/workbook analysis settings, including `visibleSeverities`,
  `untrackedRules`, and `ruleSeverityOverrides`.
- Consolidated workbook-specific configuration under
  `<workbook>.xlide_settings.json`; older workbook sidecar names are not part of
  the supported v2 settings contract.
- Moved remaining binder, designer/UserForm metadata, external metadata,
  host-metadata completeness, workbook-to-workbook transfer, and performance
  scale work to `docs/roadmap_version_2.1.0.md`.
- Kept hard object-member absence diagnostics limited to exhaustive receiver
  surfaces; incomplete host/designer/external surfaces can power completion and
  hover without inventing red `member-not-found` errors.
- Updated macro/test workflows so execution remains explicit, Windows Excel COM
  remains execution-only, and normal read/edit/analyze/sync workflows continue
  to use the Python backend.

### Fixed

- Malformed workbook settings sidecars now report explicit settings errors
  instead of being silently treated as empty defaults.
- Tightened Live Share guest messaging around the current supported behavior:
  guests can edit host-opened VBA buffers, but only the host can browse/open new
  workbook modules through XLIDE.
- Refined release docs and roadmap references so v2 is closed and forward
  scope points to the v2.1.0 roadmap and related sub-roadmaps.

## [1.0.9] - 2026-05-26

### Added

- **`xlide_listWorkbooks`** agent tool - discovers all `.xlsm`, `.xlsb`, and
  `.xlam` files in the workspace so an agent can find the target workbook.
- **`xlide_getWorkbookInfo`** agent tool - returns sheets, VBA modules, and
  named ranges in one round trip.
- **`xlide_listSheets`** agent tool - lists worksheet names and used dimensions
  for cell-range discovery.
- **`xlide_readFormulas`** agent tool - reads raw formula strings instead of
  computed values.
- **`xlide_runOpenpyxl`** agent tool - executes arbitrary openpyxl code against
  a workbook for worksheet-level automation.
- **`xlide_renameModule`** and **`xlide_deleteModule`** agent tools - expose
  existing module mutation support to AI agents.
- **`.github/copilot-instructions.md`** - canonical XLIDE agent workflow loaded
  automatically by Copilot in the repository.

### Fixed

- Clarified the `xlide_writeModule` tool description: writing a missing module
  name creates the module.
- Updated `xlide_readCells` and `xlide_writeCells` descriptions to direct
  agents through `xlide_listSheets` when the sheet name is unknown.

## [1.0.8] - 2026-05-26

### Fixed

- Refreshed the affected module's procedure list after Rename Symbol or manual
  source edits so the XLIDE Explorer no longer shows stale procedure names until
  a full refresh.

## [1.0.7] - 2026-05-26

### Changed

- Unified the XLIDE Explorer welcome message for host and Live Share guest
  sessions, including workbook discovery, Refresh, and Live Share notes.

## Earlier Development Notes

These notes predate the current v2 changelog structure and are kept for
historical context.

### Added

- Module tree accordion behavior and debounce.
- Add Class Module command and context-menu entry.
- Improved module type detection for UserForms and document modules.
- COM window restore/focus after opening a workbook or running a macro.
- Outline/breadcrumb symbol-kind polish for VBA declarations.
- Per-workbook module-list caching, filesystem watcher debounce, and RPC
  cancellation token support.
- Smoke test command, TypeScript/Python unit tests, and CI workflow.
- Protected-workbook editing, signature-invalidation notices, protection and
  signature badges, Validate VBA Project, and New Macro-Enabled Workbook.
- Early VBA snippets, enter-time block closing, status bar items, workbook-locked
  error UX, and marketplace display-name polish.

### Changed

- Reorganized context menus into create, edit, workbook, transfer, and settings
  groups.
- Split class-module creation into `xlide.newClassModule` instead of routing it
  through the generic new-module path.

---

*XLIDE follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) conventions.*
