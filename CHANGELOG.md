# Changelog

All notable changes to **XLIDE: VBA for VS Code** are documented here.

## [2.5.2] - 2026-06-20

A diagnostics accuracy patch driven by three false-positive-shaped findings on
real stdVBA modules, each adjudicated against the Excel/VBE oracle.

### Fixed

- **`required-param-after-optional` no longer flags a Property Let/Set value
  parameter.** The mandatory value parameter of a `Property Let`/`Property Set`
  may legally follow an `Optional` index parameter (e.g.
  `Property Let X(Optional ByVal i As Long, ByVal v As Long)`); only non-value
  parameters obey the required-after-optional ordering. An interior required
  parameter after an `Optional` one is still flagged. (VBE-oracle verified.)
- **`exit-wrong-proc` no longer flags `Exit Function`/`Exit Sub` inside a
  `Property Get`.** A `Property Get` is value-returning, and VBE accepts every
  `Exit` kind there; `Property Let`/`Set`, `Sub`, and `Function` still require
  the matching keyword. (VBE-oracle verified across the full truth table.)

### Changed

- **`set-required` is reclassified from a compile error to a deterministic
  runtime error (Run-time error 91).** A bare assignment to an object target
  (`obj = …` without `Set`) compiles cleanly and fails only when it executes —
  VBE-oracle verified across `= Null`, `= New`, and Function-return-name
  assignment. The finding itself is unchanged (it is still a real bug, e.g.
  `protGetNextDescendent = Null` should be `Set … = Nothing`); only its evidence
  label is corrected, matching `object-variable-not-set`. The rule may now be
  downgraded to a warning via settings.

### Internal

- Added eight VBE-oracle fixtures covering the property value parameter, the
  `Exit`-in-Property truth table, and the missing-`Set` runtime-91 cases, and
  promoted the three rules to `vbe-oracle-verified` in the diagnostic influence
  audit.

## [2.5.1] - 2026-06-20

A correctness patch for the diagnostics engine.

### Fixed

- **`module-declaration-in-procedure` no longer suppresses a module's entire
  diagnostic output.** When a comment-only line was the first statement after a
  conditional-compilation directive (`#If` / `#Else` / `#End If`) inside a
  procedure body — a common 32/64-bit pattern — the rule's alternative-header
  probe tokenized the line to an empty list and threw a `TypeError`. Because
  `analyzeModule` converts any rule exception into an empty result, that single
  throw silently discarded **every** diagnostic for the affected module, not just
  the offending line. `tokenName` is now guarded against an empty token list, and
  its two near-identical copies are consolidated onto one implementation so the
  guard cannot drift again. (Behaviorally, the unguarded copy also mis-stripped an
  unterminated `[name` bracketed identifier; the unified version is correct.)

## [2.5.0] - 2026-06-17

Version 2.5.0 builds on the v2.4.0 static-analysis baseline with two goals, both
reaching their definition of done: completing the MS-VBAL §5.6 expression binder
and cashing it in for binder-dependent diagnostics, and finishing the
syntax-corpus mining so no source file is left un-dispositioned. The same
no-false-positive discipline applies — every shipped red has positive, negative,
and no-diagnostic controls plus a named evidence source (MS-VBAL, the Excel/VBE
oracle, or deterministic XLIDE metadata), and anything not provable stays quiet
and is deferred with a documented reason. The TypeScript suite grew to 2,071
tests; the Excel/VBE oracle now backs the diagnostics with 397 verified cases.
See `docs/static_analysis_completeness_2.5.0.md` for the auditable record.

### Added

- **`argument-shape-mismatch`** (compile-error): a bare array variable or
  same-module user-defined `Type` value passed where a parameter is a scalar — or
  a scalar (including `Variant`) passed where a parameter is declared an array —
  is a VBE compile error. The rule decides on declared shape only, never
  element-type coercion; it is oracle-verified across 9 cases and is disjoint from
  `byref-argument-type-mismatch`.
- **Operator-shape diagnostics**: `non-scalar-binary-operand` (an array or
  same-module `Type` used as the operand of a scalar-requiring operator),
  `is-operator-non-object` (`Is` on a provably scalar operand), and
  `typeof-is-always-false` (a `TypeOf x Is Y` that can never hold), all off the
  §5.6 expression AST.
- **Expression-AST structuring** for named arguments (`name:=expr`), omitted
  arguments, and bang (`!`) member access.

### Changed

- **Flow precision**: the shared dataflow now merges `If`/`ElseIf`/`Else` branch
  arms, so `object-variable-not-set` (Run-time error 91) and
  `unallocated-dynamic-array-access` (Run-time error 9) check accesses inside
  balanced `If` arms — conservatively falling back on any label, `GoTo`,
  `On Error`, or loop. Default-member (`VB_UserMemId`) matching is now a
  deterministic numeric parse against a named DISPID constant.
- **Syntax corpus fully mined**: every remaining `mining` source file is
  dispositioned and promoted to `reference`; zero files remain in `mining`.

### Performance

- A behavior-preserving pass on the per-edit analysis hot path: the operand rules
  share **one** expression-tree traversal per procedure body via an
  expression-visitor registry (rather than one walk each);
  `procedureHasUnstructuredFlow`, `moduleNonCallableSymbols`, and the same-module
  `Type`-name set are memoized per parse; `me-outside-object-module` rides the
  shared statement walk; and the branch-merge join drops a per-name allocation.
  The full 2,071-test suite stays byte-identical.

### Removed

- The one-time popup recommending users disable AI inline (ghost-text) completions
  for XLIDE VBA modules. XLIDE now coexists with inline suggestions — the `smartTab`
  keybinding yields to a visible suggestion so `Tab` accepts the ghost text — so
  the recommendation was obsolete.

### Deferred (documented)

- The comparison / Boolean / string-concatenation scalar-coercion matrix, Date
  coercion, and default-member-aware diagnostics — VBA coerces these at runtime,
  so no no-false-positive compile red is provable. Numeric/host boundary overflow
  and flow phase 2 (definite assignment) remain oracle- / binder-gated. See the
  completeness report.

## [2.4.0] - 2026-06-14

Version 2.4.0 is the static-analysis completeness release. It closes the
evidence-led completeness sprint: every shipped diagnostic now has positive,
negative, and no-diagnostic (no-false-positive) controls plus a named evidence
source — MS-VBAL, the Excel/VBE oracle, or deterministic XLIDE metadata — and
everything that cannot be proven without the expression binder or further oracle
mapping is explicitly deferred with a documented reason. The release ships with
an auditable completeness record (`docs/static_analysis_completeness_2.4.0.md`).
The TypeScript test suite grew to 1,954 tests; 342 Excel/VBE oracle cases now
back the diagnostics.

### Added

- **Numeric overflow diagnostics**, oracle-verified against the live Excel VBE:
  out-of-range `Long` and `Currency` literal assignments/arguments (Run-time
  error 6), and a new compile-time `suffixed-literal-overflow` for over-range
  `%` (Integer) type-suffixed literals. The `&` (Long) suffix is deliberately
  excluded because it is ambiguous with the concatenation operator.
- **Declaration and identifier diagnostics**: invalid identifier start/character,
  empty `Type`, duplicate `Type` fields, too-many-parameters, identifier-too-long,
  `Optional`/`ByVal` UDT-parameter constraints, and non-constant `Const`/`Enum`/
  parameter-default values.
- **Control-flow and statement diagnostics**: stray `Else`/`ElseIf` outside an
  `If`, duplicate `Case Else`, `Me` outside an object module, invalid assignment
  targets, `Open` missing `For`, `TypeOf` missing its operand, and impossible /
  too-many array-declaration bounds.

### Fixed

- The structural block-balance pass no longer emits phantom
  `unmatched-block-closer` / `missing-block-closer` errors when a `: Rem ...`
  comment trails a statement (`stripVba` now blanks `Rem` at any statement start).

### Changed

- Per-rule evidence audit across all 112 diagnostic codes; the syntax corpus and
  diagnostic-influence audit are now provenance-tracked and test-enforced.
- Minor live-diagnostics performance: two whole-source rules now use the shared
  cached tokenizer, removing redundant per-keystroke lexing.

## [2.3.0] - 2026-06-11

Version 2.3.0 is the audit-remediation release: roughly 205 commits of
performance, correctness, and packaging hardening across the analyzer, editor
providers, commands, webviews, Python bridge, and VSIX layout. The headline is
the VBA analysis engine, whose quadratic hot paths were rebuilt to scale
near-linearly with module size.

Internally, the release also restructured the code it hardened: the
diagnostics engine moved to a rule registry with per-family rule modules,
commands split into per-domain modules, webview HTML/CSS/JS moved to template
assets under `assets/webview/`, workbook module operations and the
per-workbook project index became shared services, and the C# and PowerShell
test-host sources were externalized to `assets/testhost/`. The TypeScript
test suite grew from 1,604 to 1,735 tests over the effort, joined by new
real-workbook Python backend tests.

### Changed

- **Analysis engine performance** - each analysis pass now lexes and parses a
  module once and shares statement tokens, signature tables, type
  environments, and member resolution across all rules, instead of re-lexing
  per statement and re-parsing the module for every dotted reference.
  Measured on a ~4,000-line module, full analysis dropped from ~115 s to
  ~0.95 s (~122x faster), and warm member completion resolves ~350x faster.
- **One shared project index per workbook** - diagnostics, completion, hover,
  signature help, navigation, and semantic tokens now read a single
  incremental per-workbook project index that folds in module changes,
  replacing four divergent caches that each rebuilt project context while
  typing.
- **Faster startup** - the Python backend now starts lazily on first use
  (expanding a workbook, an XLIDE command, or the sidebar becoming visible)
  instead of spawning in every window at activation, the Excel host object
  model builds lazily, and the extension bundle is smaller; measured window
  startup improved by about 12%. Workspaces without Excel workbooks never
  start Python.
- Rebuilt Smart Enter on the analyzer lexer's stripped-line substrate, so
  statements hidden in trailing `: Rem ...` comments no longer fool block
  auto-close, `With` continuation, or loop-iterator sync.
- Shrank the VSIX download from about 2.1 MB to about 536 KB: the marketplace
  icon went from 1.69 MB to a 28 KB 256x256 PNG, dev-only commands are hidden
  from the Command Palette outside development mode, and local development
  configuration and coverage output no longer leak into the package.

### Fixed

- Fixed silent data loss when saving a module that had also been edited in
  the Excel VBE: module file stats now derive from the real workbook file
  mtime, so VS Code's save-conflict detection prompts instead of silently
  overwriting the newer workbook copy.
- Fixed `getWorkbookInfo` and `listSheets` crashing on every workbook with
  openpyxl 3.1 and newer, where `ReadOnlyWorksheet` lost its `dimensions`
  property and `defined_names` became a dict.
- Fixed the lexer reading a file-number `#` (for example `Write #1, x`) as a
  date-literal opener, which swallowed the rest of the statement's tokens.
- Space-triggered completion no longer fires the full resolver cascade for
  spaces typed in ordinary code; `As `-type completion, `End`/`Exit`
  keywords, event stubs, and labels still work.
- Module-name input boxes now share one validator, closing the creation path
  that accepted digit-leading module names VBA rejects.
- The `xlide_createWorkbook` agent tool now refuses to overwrite an existing
  workbook instead of silently replacing it.
- A crashed Python backend now offers a Restart Backend action instead of
  asking for a whole window reload.
- Live Share guest writes are now recorded in the host's write audit and
  refresh the host's open editors instead of clobbering the guest's change on
  the next save.
- Stale dirty-module backups are now pruned on activation instead of
  accumulating in extension global storage.

## [2.1.2] - 2026-06-09

### Changed

- Refined the marketplace README positioning for new VBA programmers, students,
  experienced VBA developers, and agentic AI workflows over real workbook
  context.
- Clarified the XLIDE workflow around auto-detected workbooks, tree navigation,
  direct workbook read/write editing, local disk push/pull sync, detailed diffs,
  and committing exported modules to version control.
- Documented that import/export is workbook-scoped when multiple workbooks are
  present, including selected-workbook previews, sidecar settings, and import
  apply behavior.

## [2.1.1] - 2026-06-09

### Fixed

- Fixed module export/import preview coloring so create rows read as additions,
  delete rows read as removals, and overwrite/update rows keep modified
  treatment.
- Improved sync preview clarity with operation-specific diff titles, semantic
  status badges, a `Select Pending` action, and clearer copy-button tooltips for
  missing workbook or repo sides.
- Rewrote the README for a broader Excel-user audience, putting the main value
  proposition, getting-started path, and full cross-platform links first.

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
