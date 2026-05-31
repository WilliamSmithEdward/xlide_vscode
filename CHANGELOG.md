# Changelog

All notable changes to **XLIDE: VBA for VS Code** are documented here.

## [Unreleased]

### Added
- **Built-in VBA runtime IntelliSense** — hovering a built-in VBA function or statement (`MsgBox`, `Left`, `CLng`, `Now`, `Array`, `RGB`, ...) now shows its verified signature and return type, and those built-ins also appear in the bare-identifier completion menu — just like resting on an intrinsic function in Visual Studio. The metadata is a curated, verified subset (~85 entries) transcribed from the Microsoft VBA language reference and MS-VBAL, never LLM-invented; names that collide with intrinsic data types (`Date`, `Time`, `String`, `Error`) are deliberately excluded so a type in an `As` position is never read as a function. A user declaration of the same name correctly shadows the built-in. The metadata lives in the pure analyzer layer (`src/analyzer/runtime/vbaRuntime.ts`) and is covered by `tests/vbaRuntime.test.ts`.
- **Hover (IntelliSense)** — hovering an identifier now shows a VBA signature and description, like resting the pointer over a symbol in Visual Studio. It describes `receiver.member` host members (e.g. `Workbook.Worksheets`), host-injected globals (`ThisWorkbook As Workbook`), worksheet/document code names (`Sheet1 As Worksheet`), and your own declarations resolved live from the module symbol graph (no save required): procedure signatures with parameters and return type, variables/parameters/constants with their `As` type, enums and members, and user types and fields — annotated with the declaring module and visibility. Unknown members are never guessed. The resolver is pure (`src/analyzer/hover/resolveHover.ts`) and covered by `tests/vbaHover.test.ts`.
- **Identifier completion (IntelliSense)** — typing a bare identifier (at a statement or expression position, not after `.` or `As`) now shows a menu of the existing objects you can reference: host-injected globals (`ThisWorkbook`, `ActiveWorkbook`, `ActiveSheet`, `ActiveCell`, `Selection`, `Application`), the workbook's worksheet/document code names (`Sheet1`, `Sheet3`, ...), and your own in-scope declarations — parameters and locals of the enclosing procedure plus module-level variables, constants, procedures, enums (and their members), and user-defined types. Candidates are filtered by what you have typed and suppressed where they would be wrong (after a member-access dot, in an `As` type position, or when naming a new declaration after `Dim`/`Const`/`Public`/`Sub`/...). The resolver is pure (`src/analyzer/completion/identifierCompletion.ts`) and covered by `tests/vbaIdentifierCompletion.test.ts`.
- **Type-name completion (IntelliSense)** — in a declaration type position (after `As` or `As New`, including parameters, function/property return types, and subsequent items in a `Dim` list), XLIDE now offers a menu of known types: the VBA built-in data types (`Long`, `String`, `Boolean`, `Object`, `Variant`, ...), the Excel host types (`Workbook`, `Worksheet`, `Range`, `Application`), and the project's own types — user `Type`s and `Enum`s declared in the current module, public (non-`Private`) `Type`s and `Enum`s declared in the workbook's other modules (read on demand and cached per workbook), plus class and UserForm module names. Candidates are filtered by what you have typed, project-defined types are listed first and can shadow a built-in of the same name, and `Decimal` is intentionally excluded because it is not directly declarable in VBA. The resolver is pure (`src/analyzer/completion/typeCompletion.ts`) and covered by `tests/vbaTypeCompletion.test.ts`.
- **Project-wide VBA symbol graph (analyzer)** — a new pure, `vscode`-free symbol layer under `src/analyzer/symbols/` builds on the AST to model every named declaration in a workbook's VBA project: modules, `Sub`/`Function`/`Property` procedures (with parameter and block-nested local children), module variables and constants, `Type` fields, `Enum` members, and `Declare` statements. The `ProjectIndex` answers hierarchical document symbols, filtered workspace symbols, conservative go-to-definition name resolution (locals/parameters, then same-module declarations, then exported `Public`/`Global` declarations in other modules), and duplicate-procedure detection — with cross-module visibility following MS-VBAL (default-`Public` procedures are exported; `Private`/`Dim`/`Friend` module members stay private). Identifier spans are located with the real lexer so a go-to-definition target never lands inside a comment or string. Covered by `tests/vbaSymbolGraph.test.ts`.
- **Host-context member completion (IntelliSense)** — typing `.` after a host object now shows a completion popup of verified Excel object-model members. It resolves host-injected globals (`ThisWorkbook`, `ActiveWorkbook`, `Application`, `ActiveSheet`, `ActiveCell`, `Selection`), worksheet/document code names (`Sheet1`, by VBA-project code name rather than tab name), `Me` by module kind (worksheet vs workbook document module), and user-declared variables/parameters/module variables typed `As Workbook`/`Worksheet`/`Range`/`Application`. Member-access chains are followed through return types, including across call parentheses (e.g. `ws.Range("A1").Offset(1, 0).`). Member metadata is transcribed from the official Microsoft Office VBA object-model reference (verified 2026-05-30) and is never LLM-generated; the resolver lives in the pure analyzer layer (`src/analyzer/host/`, `src/analyzer/completion/memberAccess.ts`) and is covered by `tests/vbaMemberCompletion.test.ts`.
- **VBA language analyzer (foundation)** — a new pure, `vscode`-free analyzer layer under `src/analyzer/`, verified against the official Microsoft VBA Language Specification (`[MS-VBAL]`, v20250520). Phase 1 ships a loss-aware, round-trippable tokenizer (`analyzer/lexer/tokenize.ts`) covering identifiers, keywords, numeric/string/date literals, comments, line continuations, separators, operators, bracketed names, and conditional-compilation markers; Phase 2 ships the spec-verified canonical keyword table (`analyzer/lexer/keywordTable.ts`) with correct VBE capitalization; Phase 3 ships an error-tolerant parser (`analyzer/parser/parseModule.ts`) that builds a `ModuleNode` AST — attributes, `Option` directives, module/`Const` declarations, `Type`/`Enum` blocks, `Sub`/`Function`/`Property` procedures with full parameter lists, and nested `If`/`For`/`Do`/`While`/`With`/`Select` block statements — never throws on malformed input, and emits block-mismatch diagnostics with source spans. Spec coverage and deviations are tracked in `docs/spec/MS-VBAL.verification-map.md`.
- **Live VBA linting** — a structural block-balance analyzer (`src/vbaLinter.ts`) reports unbalanced constructs as diagnostics while you type: missing `End Sub`/`End Function`/`End Property`/`End If`/`End With`/`End Select`/`End Type`/`End Enum`/`Next`/`Loop`/`Wend`, stray closers with no opener, and inner blocks left unclosed. Strings, comments, and `_` line continuations are handled.
- **Smart Enter (auto-block)** — typing a `Sub`/`Function`/`Property` header and pressing Enter auto-inserts the matching `End ...` below, leaving the caret on the indented body line, just like the VBA IDE.
- **Protected-workbook editing** — VBA module writes/renames/deletes now save with `allow_protected=True`, so password-locked VBA projects can be edited without errors.
- **Signature-invalidation notice** — when an edit drops a workbook's VBA digital signature, XLIDE surfaces a one-time warning (per workbook, per session) prompting the user to re-sign externally. The signature state is captured via `signatureDropped` from the Python layer rather than silently discarded.
- **Protection/signature badges** in the XLIDE Explorer — workbook nodes now show `[locked]` and/or `[signed]` tags (lazily probed) so the project's protection state is visible at a glance.
- **Validate VBA Project** command + `xlide_validateWorkbook` agent tool — audits a workbook's VBA project for cross-structure inconsistencies and reports any issues to the XLIDE Output channel.
- **New Macro-Enabled Workbook** command (view title + Command Palette) + `xlide_createWorkbook` agent tool — creates a fresh `.xlsm`/`.xlsb` with an empty VBA project from pyOpenVBA's baked-in template.
- **`getProtectionInfo` / `validateWorkbook` / `createWorkbook`** JSON-RPC handlers in the Python server; `get_workbook_info` now also returns `isPasswordProtected` and `isSigned`.

### Changed
- **F5 (Run Macro at Cursor)** now saves the active module first when it has unsaved changes, so the macro that runs reflects the current editor source rather than the last-saved version.
- **Find All References** now excludes the procedure declaration token itself, returning only call sites; added a **Find All References** entry to the Explorer tree-view context menu for Sub/Function nodes.

## [1.0.9] - 2026-05-26

### Added
- **`xlide_listWorkbooks`** agent tool — discovers all `.xlsm`/`.xlsb`/`.xlam` files in the workspace so the agent never needs to be told a file path.
- **`xlide_getWorkbookInfo`** agent tool — single round-trip returning sheets, VBA modules, and named ranges together.
- **`xlide_listSheets`** agent tool — sheet names and used dimensions for cell-range discovery.
- **`xlide_readFormulas`** agent tool — reads raw formula strings (`=SUM(A1:A10)`) instead of computed values.
- **`xlide_runOpenpyxl`** agent tool — executes arbitrary openpyxl Python code against a workbook, exposing the full openpyxl API (styling, charts, number formats, conditional formatting, etc.).
- **`xlide_renameModule`** agent tool — renames a VBA module (Python layer already supported this; now exposed to AI agents).
- **`xlide_deleteModule`** agent tool — deletes a VBA module (same).
- **`.github/copilot-instructions.md`** — canonical XLIDE agent workflow loaded automatically by Copilot in every session.

### Fixed
- **`xlide_writeModule` description** clarified that passing a non-existent module name creates the module automatically.
- **`xlide_readCells` / `xlide_writeCells` descriptions** updated to reference `xlide_listSheets` for sheet discovery.

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
