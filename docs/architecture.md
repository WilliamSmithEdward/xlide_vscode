# XLIDE – Architecture

## Overview

XLIDE is a VS Code extension that turns Excel macro files (`.xlsm`, `.xlsb`, `.xlam`) into first-class editable documents. VBA modules open in the editor like normal source files, Ctrl+S writes them back into the workbook, and 15 Language Model tools expose every operation to Copilot and other VS Code AI agents.

The extension is split into two layers connected by a long-lived child process:

```
VS Code Extension (TypeScript)
        |
        |  child_process.spawn — JSON-RPC 2.0, newline-delimited over stdio
        |
Python Backend  (pure Python, no COM, no Office install required)
        |-- pyOpenVBA   VBA module read / write
        +-- openpyxl    Excel cell data read / write
```

---

## Repository layout

```
xlide_vscode/
  src/
    extension.ts        Activation entry point — registers all providers and commands
    pythonBridge.ts     PythonBridge class — spawns server.py, JSON-RPC 2.0 client
    xlsmExplorer.ts     XlsmExplorer — TreeDataProvider for the XLIDE sidebar
    xlideFileSystem.ts  XlideFileSystemProvider — virtual xlide-vba:// filesystem
    commands.ts         Command handlers: open/new/rename/delete, workbook open/run, export modules
    agentTools.ts       LanguageModelTool registrations for AI agent use
    moduleExport.ts     Shared export/config logic for UI commands and AI tools
    liveShare.ts        LiveShareIntegration — host/guest Live Share bridge over the VSLS service API
    statusBar.ts        XlideStatusBar — two status bar items (active module, Live Share guest indicator)
    vsls.d.ts           Ambient type declarations for the VS Code Live Share extension API
    vbaSymbolIndex.ts   VbaSymbolIndex — workbook-scoped cache of parsed VBA symbols
    vbaLanguageProviders.ts  Document/definition/reference/rename providers, diagnostics, and smart-enter for the vba language
    vbaLinter.ts        Pure structural block-balance analysis (lintVbaSource) and smart-enter helpers (no vscode dependency)
    vbaWorkbookLint.ts  Shared workbook-wide lint core (lintWorkbook) reused by the Lint-All command and the xlide_lintWorkbook agent tool; flattens lintVbaSource + analyzeModule into 1-based {moduleName,line,column,severity,code,message} problems
    analyzer/
      lexer/
        keywordTable.ts MS-VBAL 3.3.5.2 reserved-identifier + contextual keyword tables with canonical casing
        tokenKinds.ts   TokenKind/Trivia/VbaToken types and WSC helpers (MS-VBAL 3.3)
        trivia.ts       Leading whitespace / line-continuation scanner (MS-VBAL 3.2.2)
        tokenize.ts     Loss-aware, round-trippable VBA tokenizer (MS-VBAL 3.3.1-3.3.5, 3.4)
      parser/
        nodes.ts        AST node types + spans + ParseDiagnostic (MS-VBAL 4.2/5.x)
        parserState.ts  Logical-statement splitter + statement cursor (MS-VBAL 3.3.1 EOS)
        parseModule.ts  Error-tolerant module parser -> ModuleNode AST (MS-VBAL 5.x)
      semantic/
        typeSemanticTokens.ts  Pure resolver for type-name semantic tokens and hover in declaration/New positions
      index.ts          Public, vscode-free analyzer surface (lexer + parser)

  python/
    server.py           JSON-RPC 2.0 dispatcher (stdin -> stdout, newline-delimited)
    xlide/
      __init__.py
      vba_io.py         list/read/write/rename/delete modules via pyOpenVBA
      excel_io.py       read/write cell ranges via openpyxl

  docs/
    architecture.md     This file
    roadmap.md          Feature roadmap

  package.json          Extension manifest, contributes, LM tool declarations
  tsconfig.json         Strict TypeScript config (module: Node16)
  esbuild.js            Bundle script — produces out/extension.js
  .vscode/
    launch.json         F5 Extension Development Host config
    tasks.json          Default build task (npm run watch)
    settings.json       Workspace Python interpreter path
  python/requirements.txt   pyOpenVBA, openpyxl
```

---

## Virtual filesystem — `xlide-vba://`

Clicking a module in the sidebar opens it under the custom scheme:

```
xlide-vba:///C:/path/to/workbook.xlsm/Module1.bas
```

`XlideFileSystemProvider` implements `vscode.FileSystemProvider`:

| Method | Action |
|---|---|
| `readFile(uri)` | Calls `readModule` on the Python bridge; returns UTF-8 bytes |
| `writeFile(uri, content)` | Calls `writeModule`; saves the .xlsm in place |
| `stat()` | Returns a synthetic, stable `FileStat`; `mtime` changes only after XLIDE saves or receives an explicit module-change notification |
| All others | Throw `FileSystemError.NoPermissions` |

VS Code treats the file as fully editable — Ctrl+S triggers `writeFile` with no extra command needed.

`src/xlideDirtyModuleBackups.ts` adds an XLIDE-owned safety layer for dirty
module editors. Because VS Code Hot Exit is not reliable enough for virtual
workbook modules, every dirty local `xlide-vba://` document is synchronously
mirrored into extension global storage. When the same module reopens after a
VS Code restart, XLIDE reapplies the stored text as an editor edit, making the
tab dirty again and offering Save/Revert actions. A successful save removes the
backup.

### URI encoding / decoding

- Encode: replace `\` with `/`; prepend `/` for Windows drive letters; append `/<moduleName>.bas`
- Decode: regex matches everything up to `.xlsm`/`.xlsb`/`.xlam` as the workbook path; the final segment (minus `.bas`) is the module name; on Windows strip the leading `/` before the drive letter

---

## Sidebar tree — `XlsmExplorer`

`TreeDataProvider<XlideNode>` with three levels:

| Level | Node kind | Children source |
|---|---|---|
| 0 | `xlsm` — one per file found by `findFiles('**/*.{xlsm,xlsb,xlam}')` | modules |
| 1 | `module` — name + type (standard / class / document) | subs |
| 2 | `sub` — procedure name, kind, 1-based line number | none |

Clicking a `module` node opens the module via `xlide.openModule`. Clicking a `sub` node opens the module and moves the cursor to that line.

Module type is inferred from the VBA source:
- Starts with `VERSION 1.0 CLASS` → `class`
- Contains `Attribute VB_PredeclaredId = True` → `document`
- Name matches `ThisWorkbook`, `Sheet\d*`, or `Chart\d*` → `document`
- Anything else → `standard`

When the hidden `VB_Base` attribute is available, `listModules` also returns an
optional `documentType` for document modules (`workbook`, `worksheet`, or
`chart`). The broad `type` field remains `document` so existing import/export
behavior has one module-kind contract, while language-service features can use
the subtype when they need workbook-vs-worksheet-vs-chart semantics.

---

## Python bridge — `PythonBridge`

Spawned once at activation with `cwd` set to `python/` so the `xlide` package is importable without installation. Communicates via newline-delimited JSON-RPC 2.0 over stdin/stdout.

**Python resolution order:**
1. `xlide.pythonPath` VS Code setting (if set)
2. `.venv/Scripts/python.exe` (Windows) or `.venv/bin/python3` (Mac/Linux) inside the first workspace folder, if it exists
3. `python` (Windows) / `python3` (Mac/Linux) from `PATH`

All calls are queued if the process has not yet started; in-flight calls are rejected if the process exits.

---

## Windows Excel COM behavior

The commands `xlide.openWorkbook` and `xlide.runMacroAtCursor` use PowerShell COM automation on Windows.

Setting:

- `xlide.attachToRunningExcel` (default `true`)
  - `true`: tries to attach to a running `Excel.Application` and reuse an already-open workbook (matched by full path or workbook name) before opening.
  - `false`: always opens through a new COM-created Excel application path.

---

## JSON-RPC methods

| Method | Required params | Optional params | Returns |
|---|---|---|---|
| `listModules` | `path` | — | `[{name, type, documentType?}]` |
| `listSubs` | `path`, `module` | — | `[{name, kind, line}]` |
| `readModule` | `path`, `module` | — | `{source}` |
| `writeModule` | `path`, `module`, `source` | — | `{ok, signatureDropped}` |
| `renameModule` | `path`, `module`, `newName` | — | `{ok, signatureDropped}` |
| `deleteModule` | `path`, `module` | — | `{ok, signatureDropped}` |
| `listSheets` | `path` | — | `{sheets: [{name, dimensions}]}` |
| `getWorkbookInfo` | `path` | — | `{modules, sheets, namedRanges, isPasswordProtected, isSigned}` |
| `getProtectionInfo` | `path` | — | `{isPasswordProtected, isSigned}` |
| `validateWorkbook` | `path` | — | `{issues: [string]}` |
| `createWorkbook` | `path` | — | `{ok, path}` |
| `readCells` | `path`, `sheet`, `range` | — | `{data: [[…]]}` |
| `readFormulas` | `path`, `sheet`, `range` | — | `{data: [[…]]}` (raw formula strings) |
| `writeCells` | `path`, `sheet`, `startCell`, `data` | — | `{ok}` |
| `runOpenpyxl` | `path`, `code` | `save` (bool, default `true`) | `{result, stdout}` |

Errors are returned as `{"error": {"code": -32000, "message": "…"}}`.

---

## Protected & signed workbooks

All mutating saves in `vba_io.py` (`writeModule`, `renameModule`, `deleteModule`) call `ExcelFile.save(allow_protected=True)`, so password-locked VBA projects can be edited in place. The save is wrapped in `warnings.catch_warnings(record=True)`: pyOpenVBA emits a `UserWarning` when it drops a now-stale digital-signature stream, and that is surfaced to the caller as `signatureDropped: true` rather than being silenced.

On the TypeScript side, `notifySignatureDropped(filePath, signatureDropped)` in `xlideFileSystem.ts` shows a one-time-per-workbook warning when a signature is invalidated. `writeFile` and the three write agent tools/commands all forward the flag.

`getProtectionInfo` reports `{isPasswordProtected, isSigned}` using public pyOpenVBA APIs (`vba_project().protection` + `detect_signature(CFB(vba_project_bytes()))`). `XlsmExplorer` lazily probes this when a workbook is expanded and renders `[locked]`/`[signed]` badges on the workbook node. `getWorkbookInfo` folds the same two flags into its summary.

`validateWorkbook` wraps `ExcelFile.validate()` (cross-structure consistency check); `createWorkbook` wraps `ExcelFile.create_new(path)` to scaffold a fresh macro-enabled workbook from pyOpenVBA's baked-in template.

---

## Module export / import

`moduleExport.ts` is the single source of truth for export/config behavior.

Both lanes call into this shared implementation:

- UI commands (`xlide.exportModulesToFolder`, `xlide.configureExportMode`)
- AI tools (`xlide_exportModules`, `xlide_configureExportMode`)

**Export** reads all modules live over JSON-RPC (`listModules` then `readModule` per module) and writes them to a folder.

- Output file extension is `.bas` for standard modules and `.cls` for class/document modules.
- Export mode is per-workbook and persisted in the workbook-local JSON config:
  - `trueUp` (default): replace existing, add new, remove no-longer-existing modules
  - `replaceExistingOnly`: replace files that already exist; do not add missing files; do not remove stale files

**Import** (`xlide.importModulesFromFolder`) reads `.bas`/`.cls`/`.frm` files from the configured (or user-chosen) folder and writes each back into the workbook via `writeModule`. A QuickPick lets the user select which files to import. Document modules and UserForms cannot be created from scratch — they are only importable if the module already exists in the workbook.

**Change export folder** (`xlide.changeRepoFolder`) updates `exportFolder` in the workbook config without running an export.
- A workbook-local config file is written beside the workbook:

```
<workbook-filename>.extension.repo.json
```

Config schema:

```json
{
  "exportFolder": "C:/absolute/path/to/export/folder",
  "exportMode": "trueUp",
  "managedFiles": ["Module1.bas", "Sheet1.cls"]
}
```

On later runs, `exportFolder` is used as the default folder in the picker.

The command `xlide.configureExportMode` updates `exportMode` for a workbook.

---

## AI agent tools

Declared in `package.json` under `contributes.languageModelTools` and registered at activation via `vscode.lm.registerTool`. Copilot can invoke them inline or via `#` references in chat.

| Tool name | Chat reference | Side effects | Confirmation |
|---|---|---|---|
| `xlide_listWorkbooks` | `#xlideListWorkbooks` | none | No |
| `xlide_listModules` | `#xlideListModules` | none | No |
| `xlide_listSubs` | `#xlideListSubs` | none | No |
| `xlide_readModule` | `#xlideReadModule` | none | No |
| `xlide_writeModule` | `#xlideWriteModule` | saves .xlsm | Yes |
| `xlide_renameModule` | `#xlideRenameModule` | saves .xlsm | Yes |
| `xlide_deleteModule` | `#xlideDeleteModule` | saves .xlsm | Yes |
| `xlide_listSheets` | `#xlideListSheets` | none | No |
| `xlide_getWorkbookInfo` | `#xlideGetWorkbookInfo` | none | No |
| `xlide_validateWorkbook` | `#xlideValidateWorkbook` | none | No |
| `xlide_createWorkbook` | `#xlideCreateWorkbook` | creates/overwrites .xlsm | Yes |
| `xlide_readCells` | `#xlideReadCells` | none | No |
| `xlide_readFormulas` | `#xlideReadFormulas` | none | No |
| `xlide_writeCells` | `#xlideWriteCells` | saves .xlsm | Yes |
| `xlide_runOpenpyxl` | `#xlideRunOpenpyxl` | may save .xlsm (controlled by `save` param) | Yes |
| `xlide_exportModules` | `#xlideExportModules` | writes export files + updates workbook JSON config | Yes |
| `xlide_configureExportMode` | `#xlideConfigureExportMode` | updates workbook JSON config | Yes |

---

## Live Share integration — `liveShare.ts`

`LiveShareIntegration` contains infrastructure for XLIDE across VS Code Live Share sessions.

**What works:** The host opens modules through XLIDE normally (creating `xlide-vba://` documents). Live Share mirrors those open documents to the guest. The guest can co-edit and save (Ctrl+S) — the save travels through Live Share's standard editor sync back to the host's `XlideFileSystemProvider.writeFile`, which writes the workbook.

**What does not work:** Independent guest browsing. The code uses the Live Share shared service API (`shareService` / `getSharedService`) under the service name `WilliamSmithE.xlide` to let guests list and open their own workbooks. Microsoft does not allow non-approved extensions to expose guest-accessible shared services — `shareService()` always returns `null`. As a result, the XLIDE sidebar shows nothing on the guest side and guests cannot independently discover or open modules; they can only collaborate on documents the host has already opened.

The host-side RPC handlers (`listWorkbooks`, `listModules`, `listSubs`, `readModule`, `writeModule`) and the guest-side `guestList*` / `guestReadModule` / `guestWriteModule` methods are implemented and would address this gap if Microsoft approval were obtained. Remote modules use `xlide-vba://liveshare/<workbookId>/<moduleName>.bas` URIs so `XlideFileSystemProvider` can route them through the proxy rather than the local Python bridge. `LiveShareIntegration.onDidChange` fires on session role changes so that `XlsmExplorer` and `XlideStatusBar` can refresh.

---

## Status bar — `statusBar.ts`

`XlideStatusBar` manages two `vscode.StatusBarItem` instances:

| Item | Shown when | Text | Click action |
|---|---|---|---|
| Active module | Active editor is an `xlide-vba://` document | `<workbook> | <module>` (or `XLIDE (Live Share)` for remote) | `xlide.refreshExplorer` |
| Live Share | Connected as a Live Share guest | `XLIDE (Live Share): <N workbooks>` | `xlide.refreshExplorer` |

---

## VBA language services

VBA is registered as the `vba` language (extensions `.bas`, `.cls`, `.frm`).

**Syntax coloring** — `syntaxes/vba.tmLanguage.json` provides a TextMate grammar
covering comments, attribute lines, `#If` directives, string/number/date literals,
procedure declarations (`Sub`, `Function`, `Property Get/Let/Set`), declarations
(`Dim`, `Public`, `Private`, `Const`, `Declare`, `Type`, `Enum`), built-in types
and constants, control-flow keywords, and built-in functions.
`language-configuration/vba-language-configuration.json` configures the
apostrophe line comment, brackets, indent rules, and procedure-based folding.

**Symbol intelligence** — `src/vbaSymbolIndex.ts` keeps a workbook-scoped cache
of parsed module symbols. Modules are parsed with a lightweight regex pass
(`parseVbaModule`) that yields each `Sub`, `Function`, and `Property Get/Let/Set`
with name range and body range. The index loads modules lazily through the
Python bridge (`listModules` + `readModule`) and can refresh a single module
after a save.

`src/vbaLanguageProviders.ts` registers the language providers plus diagnostics
and smart-enter editing against the `vba` language under the `xlide-vba` scheme:

| Provider | Behavior |
|---|---|
| `DocumentSymbolProvider` | Outlines the current module from `parseVbaModule` |
| `DefinitionProvider` | Builds an AST `ProjectIndex` and resolves source-backed `object.Member` references through the shared member-completion binder, resolves project type-name tokens through `resolveTypeDefinitions`, then falls back to scope-aware name resolution (`resolveDefinition`); honors a `Module.Member` qualifier via `resolveQualifiedDefinition`, and follows MS-VBAL visibility (locals shadow module members shadow exported cross-module declarations, including enum members exported by their containing `Enum`) |
| `ReferenceProvider` | Uses semantic binding before textual search: source-backed `object.Member` references are matched by their resolved class-member definition spans, project type-name tokens are matched through `resolveTypeDefinitions`, and ordinary identifiers still use `ProjectIndex.referenceScope` plus word-boundary search restricted to the binding scope; honors VS Code's include-declaration toggle |
| `RenameProvider` | Uses the same source-backed member binding before falling back to `referenceScope`, so workbook class members rename only their own declarations/usages; class component rename is intentionally tree-only because the VBA class name is the module/component name rather than an in-source declaration |
| Diagnostics | Debounced structural lint (`lintVbaSource`) flags unbalanced blocks — missing `End Sub`/`Next`/`Loop`/..., stray closers, and inner blocks left unclosed |
| Smart enter (auto-block) | Pressing Enter after a `Sub`/`Function`/`Property` header auto-inserts the matching `End ...` below and leaves the caret on the indented body line |

Language-service business rules are unified across surfaces. Unless a behavior is
called out as a deliberate corner case, completion insert text, hover, signature
help, diagnostics, navigation, rename, tree actions, semantic coloring, snippets,
formatter logic, and code actions must all consume the same analyzer rules or
provider helpers for the same VBA construct.

The `DefinitionProvider`, `ReferenceProvider`, and `RenameProvider` build a fresh
`ProjectIndex` (`src/analyzer/symbols/projectIndex.ts`) from the cached module
sources on each query, with the live editor text overlaid for the current
module. Offset-based symbol spans are converted to editor ranges in the shared
`vbaNavigation.ts` helpers and provider wiring. `VbaSymbolIndex` still backs the
`DocumentSymbolProvider` outline and the workbook-scoped source cache.
Tree-level class module rename uses the same project type-reference helpers:
the component name is changed through `renameModule`, and VS Code reference edits
are applied for the workbook's project-defined class type tokens.

**Structural linting** — `src/vbaLinter.ts` is a pure, `vscode`-free module so it
is unit-tested directly (`tests/vbaLinter.test.ts`). It strips strings/comments,
joins `_` line continuations, then walks a block stack to detect imbalance. The
same module exports `detectProcOpener`/`isProcClosedAhead` used by the
smart-enter feature.

**Conditional compilation model** — The core parser models `#Const`, `#If`,
`#ElseIf`, `#Else`, and `#End If` as `ConditionalDirective` AST nodes at module
and procedure scope. `src/analyzer/conditional/conditionalCompilation.ts` is the
shared, `vscode`-free helper for source-order directive collection, `#Const`
indexing, high-confidence expression evaluation against supplied compiler
constants (`VBA7`, `Win64`, `Win32`, `Mac`, etc.), and `active` / `inactive` /
`unknown` branch activity. If the caller does not supply a compiler environment,
platform constants remain unknown. `createConditionalActivityTracker` is the
shared branch predicate used by `buildModuleSymbols`, `ProjectIndex`, and active
diagnostics, so symbols, same-module/project call signatures, duplicate
declaration checks, type/call validation, and Win64 `Declare PtrSafe`
diagnostics all skip only branches proven inactive and leave unknown branches
visible.

The index also subscribes to `onDidSaveTextDocument` for `xlide-vba://` URIs so
the cache stays in sync with user edits.

**Workbook-wide lint (command + agent tool)** — `src/vbaWorkbookLint.ts`
(`lintWorkbook`) is the shared core that loads every module from the workbook via
the Python bridge, builds a `ProjectIndex` so cross-module rules have the
  current module's visibility-filtered procedure/Declare and bare identifier names, then runs both diagnostic passes
(`lintVbaSource` + `analyzeModule`) per module and flattens their results into
1-based `{moduleName, moduleType, line, column, endColumn, severity, code,
message}` problems, sorted by module/line/column. The
`xlide.lintWorkbook` command (`src/commands.ts`, right-click "Lint All Modules in
Workbook" on a workbook tree node) prints a formatted, blank-line-padded report
to the XLIDE Output channel without switching focus, and shows a summary
notification. Each problem carries a clickable location link built from
`encodeModuleUri(...).with({fragment: 'L<line>,<col>'})` and normalized to the
detectable `xlide-vba://...#L<line>,<col>` triple-slash (empty-authority) form so
the monaco LinkComputer linkifies it and the opener's `extractSelection` reveals
the exact line. The same core is exposed to AI agents as the `xlide_lintWorkbook`
LM tool (`src/agentTools.ts`), returning the structured JSON report so an agent
can verify lint passes in real time after editing modules.

**VBA analyzer (ground-up language service)** — `src/analyzer/` is a pure,
`vscode`-free TypeScript library being built per
`docs/xlide_vba_language_service_roadmap.md`, verified against
`docs/[MS-VBAL].pdf` (v20250520). Phase 1 (lexer), Phase 2 (canonical keyword
table), and Phase 3 (error-tolerant parser/AST) are in place:
`analyzer/lexer/tokenize.ts` is a loss-aware, round-trippable tokenizer,
`analyzer/lexer/keywordTable.ts` is the spec-verified reserved-identifier +
contextual-keyword table with canonical casing, and
`analyzer/parser/parseModule.ts` builds a `ModuleNode` AST (attributes, options,
declarations, Type/Enum, procedures + parameters, and nested block statements)
that never throws on malformed input and emits block-mismatch diagnostics. Every
rule cites an MS-VBAL section; coverage and deviations are tracked in
`docs/spec/MS-VBAL.verification-map.md`. This layer will eventually replace the
interim regex linter in `src/vbaLinter.ts`.

**Host-context member completion** — built on top of the analyzer, this is the
feature that distinguishes XLIDE from a generic VBA syntax extension. It is split
into a pure analyzer layer and a thin VS Code provider:

- `src/analyzer/host/excelObjectModel.ts` is a curated, verified subset of the
  Excel automation object model (Application/Workbook/Worksheet/Range plus the
  commonly used Window, Name(s), Comment(s), ListObject/Row/Column(s),
  PivotTable(s), Chart(s)/ChartObject(s), Shape(s), Font, Interior, Border(s),
  Areas, Hyperlink(s), WorksheetFunction, Style(s), PageSetup and Validation
  types), with each type's properties and methods transcribed from the official
  Office VBA object-model reference (`learn.microsoft.com/office/vba/api/excel.*`)
  and cross-checked against the Excel COM type library. Return types are wired so
  member-access chaining flows into these types (e.g. `Range.Font.`,
  `ws.ListObjects(1).Range.`). It also holds the host-global table
  (`ThisWorkbook` -> `Excel.Workbook`, `Application` -> `Excel.Application`, ...)
  and the `As`-type alias table. The reference corpus under `reference/` is
  development context only: production extension code must not read it from disk
  at runtime, and `reference/**` is excluded from the packaged extension. When a
  reference dump is promoted into runtime behavior, the promotion must generate
  or update checked-in metadata under `src/` with explicit provenance. The Excel
  generator currently emits `Application`, `Workbook`, `Worksheet`, `Range`,
  `Workbooks`, `Worksheets`, and `Sheets` into
  `src/analyzer/host/excelReferenceMembers.ts`, including available member
  signatures and reference documentation; `Excel.Workbook` and `Excel.Worksheet`
  are marked exhaustive for hard unknown-member diagnostics. Reference events
  are counted in coverage but filtered out of object-member surfaces because VBE
  does not expose events as callable object methods/properties; document-module
  event handler authoring uses separate module-scoped metadata in
  `src/analyzer/completion/eventHandlers.ts`. Hand-curated host
  members merge with matching generated entries so completion, hover, and
  signature help use the same enriched member surface instead of a parallel
  fallback path. The generator also writes `docs/excel_reference_coverage.md` so
  promotion gaps stay visible.
  LLM-generated member lists are never used; this is host metadata, not VBA
  grammar.
- `src/analyzer/host/hostModel.ts` exposes pure resolver functions over that
  metadata (`resolveHostGlobal`, `resolveHostAlias`, `getHostMembers`,
  `resolveMemberReturnType`).
- `src/analyzer/completion/memberAccess.ts` tokenizes the source up to the
  cursor, detects a member-access dot, walks the receiver chain (handling call
  parentheses and collection-default `Item` paths for chains like
  `ThisWorkbook.Worksheets(1).Range("A1").` and merged worksheet/chart
  `Workbooks(1).Sheets(1).Range("A1").`), resolves the root (`Me`, a host
  global, a worksheet code name, or a typed local/module
  variable found by parsing the module), follows member return types through the
  chain, and returns the filtered members. When a typed variable proves a more
  specific member surface after assignment from a mixed collection (for example
  `Dim ws As Worksheet: Set ws = Sheets(1)`), the declared type wins over the
  collection's merged item surface. For generic `Object`/`Variant` variables,
  the resolver can refine completion from the latest preceding simple `Set`
  assignment to a known object expression, so `Set obj = Worksheets(1)` narrows
  `obj.` to worksheet members while `Set obj = Sheets(1)` keeps the merged
  worksheet/chart surface. Diagnostic callers can opt out of this refinement so
  late-bound `Object`/`Variant` receivers are not treated as hard absence proof;
  editor completion leaves it enabled. It also accepts a source-backed
  workbook class-member surface from `ProjectIndex.projectClassMembers()`, so
  variables declared as workbook classes (for example `Dim p As Person`) can
  offer public/default-public source members and public fields at `p.` without
  guessing from names. Public constants are not exposed as object members because
  VBE rejects them in class/document/UserForm modules. The same context also
  resolves `Me.` to the current class/document module's source-backed member
  surface, merging with a known host surface when the caller supplies one.
  Source-backed workbook members carry inline `'''` documentation through to
  completion, hover, and member-call signature help, and carry declaration spans
  for source-backed member go-to-definition. Exported member attributes such as
  `Attribute Value.VB_UserMemId = 0` are attached to the same source-backed
  member surface and mark `defaultMember`, but direct object-expression
  inference is not enabled until that behavior has separate oracle coverage.
- `src/vbaMemberCompletion.ts` is the VS Code `CompletionItemProvider` (trigger
  characters `.` and space). For member access it builds the project context
  from the workbook's module list (document code names with workbook,
  worksheet, or chart host type where known, plus the host/source
  `Me` context for the current object module) via the Python bridge and renders
  the resolved members. For workbook class members, open XLIDE module documents
  are read from their live editor text first, so unsaved changes in an open
  `Person` class are reflected the next time completion is requested elsewhere; saved module text
  is read through the bridge when no live editor text exists. At module level in
  document modules, it also offers event-procedure stubs from
  `resolveEventHandlerCompletions`: `ThisWorkbook`/`documentType: workbook`
  gets `Workbook_*` handlers, worksheet document modules get `Worksheet_*`
  handlers, and chart/UserForm handler authoring stays out until those event
  surfaces have curated metadata. Existing handlers are not re-suggested, and
  accepting an event completion inserts either the full `Private Sub ... End
  Sub` stub or only the declaration tail after an existing `Private Sub`
  prefix. Known workbook/worksheet event-handler-shaped `Sub` declarations in
  the wrong module receive non-red `event-handler-module-scope` guidance because
  Excel will not wire them as events there, even though they may compile as
  ordinary procedures. In a declaration
  type position (after `As` / `As New`) or expression-level `New` position it instead
  offers type-name completions via `src/analyzer/completion/typeCompletion.ts`
  (`resolveTypeCompletions`): VBA built-in data types, the Excel host types, and
  project-defined types from `ProjectIndex.visibleTypeNames()` with the live
  editor text overlaid for the current module: object-module names plus visible
  `Type`/`Enum` declarations, preserving duplicate names as ambiguous. Project-defined type
  candidates carry inline `'''` docs from `Type`/`Enum` declarations and from
  object-module header docs, so `As Person` and `New Person` completion/hover
  can show the same documentation model as member surfaces. When the cursor is on a bare
  identifier (statement/expression position, not after `.` or `As`) it offers
  identifier completions via `src/analyzer/completion/identifierCompletion.ts`
  (`resolveIdentifierCompletions`): host-injected globals (`ThisWorkbook`,
  `ActiveSheet`, `Application`, ...), worksheet/document code names, the
  user's in-scope declarations (parameters, locals, module variables/constants,
  procedures, external Declares, enums and their members, user types), visible
  cross-module project declarations from `ProjectIndex.visibleIdentifierSymbols()`,
  built-in VBA runtime
  functions (`MsgBox`, `Left`, `CLng`, `RGB`, ...), and built-in constants
  (`vbOKOnly`, `xlUp`, ...) from runtime/host metadata once a constant-like
  prefix is typed.
  Runtime completion details show the verified signature, and the documentation
  panel includes the runtime kind plus curated parameter types where available;
  constant completion shows the owning enum/type and known value.
  Curated runtime calls are intentionally not duplicated as VS Code snippets.
  Expression-level `New` completion is narrower and offers only creatable
  project classes/UserForms until host/external creatability metadata exists.
  Accepting callable completions applies canonical casing and uses the shared
  VBA call-site rule: standalone call statements insert only the canonical name,
  while expression contexts and explicit `Call` statements may insert `()` with
  the cursor inside the call. Runtime entries that opt out of explicit `Call`
  through verified metadata are filtered at the `Call <target>` position, so
  invalid forms such as `Call DoEvents` are not offered there. Typing a boundary after a known identifier, moving the
  cursor away from the identifier, or leaving the editor applies VBE-style
  canonical casing for keywords, type names, runtime functions, and resolved host
  members.
- The same `src/vbaMemberCompletion.ts` class also registers a VS Code
  `HoverProvider`. It delegates to `src/analyzer/hover/resolveHover.ts`
  (`resolveHover`), a pure resolver that describes the identifier under the
  cursor: `receiver.member` host/reference or source-backed workbook members
  (reusing the same member-access resolver as completion), host globals,
  worksheet code names, type names in declaration/`New` positions (using the same
  primitive/host/project resolver as type completion and semantic coloring), user
  declarations from the live module symbol graph (procedure
  signatures with parameters and return type, variables/parameters/constants
  with their `As` type, enums and members, user types and fields), built-in
  constants, and built-in VBA runtime functions, annotated with the declaring
  module, visibility, or metadata source.
  Unknown members are never guessed. Built-ins resolve last so a user
  declaration of the same name shadows the built-in.

**Built-in VBA runtime metadata** — `src/analyzer/runtime/vbaRuntime.ts` is a
curated, verified subset of the intrinsic VBA runtime library (~85 functions and
statements: `MsgBox`, `InputBox`, the `C*` conversions, string/date/math
helpers, `Array`, `UBound`, `RGB`, ...), plus common runtime constants and enum
members such as `vbOKOnly`, `vbCrLf`, and `vbFalse`. Each `VbaRuntimeFunction` carries a
canonical `signature`, optional `returns`, a `kind` of `function | statement`,
optional explicit-`Call` compatibility metadata, and `source: 'verified'`. Signatures are transcribed from
learn.microsoft.com/office/vba/language and MS-VBAL, never LLM-invented. Names
that collide with intrinsic data types (`Date`, `Time`, `String`, `Error`) are
deliberately omitted so a type in an `As` position is never read as a function.
Like the host model, this is a typed TS module (not the JSON file the roadmap
originally suggested) for compile-time checking. `resolveRuntimeFunction(name)`
and `resolveRuntimeConstant(name)` resolve case-insensitively, while
`runtimeAllowsExplicitCall(fn)` centralizes runtime-specific explicit `Call`
behavior such as the VBE-oracle-verified `DoEvents` special case;
`VBA_RUNTIME_FUNCTIONS` and `VBA_RUNTIME_CONSTANTS` are consumed by hover,
identifier completion, and high-confidence diagnostics.

**Signature help (parameter info)** — `src/analyzer/signature/signatureHelp.ts`
computes the VBE call tip from module text alone. `resolveSignatureHelp(source,
offset, ctx)` returns a `SignatureInfo` (the signature `label`, ordered
`parameters`, and the `activeParameter` index) or `undefined`. It tokenizes the
prefix up to the caret, maintains a paren-frame stack to find the innermost
enclosing *call* paren (a `(` directly preceded by an identifier; grouping/index
parens are skipped), counting top-level commas for the active parameter; when no
call paren is open it falls back to a conservative *parenless call statement*
detector (`Workbooks.Open "file", `) that bails on statement keywords, file-I/O
starters, and top-level `=` assignments. Member-call signatures are resolved
through the same member-completion route used for `object.` completion, so host
members and source-backed project class members share receiver binding,
return-type chains, and inline XML docs. Bare-call signatures are sourced from
same-module user `Sub`/`Function`/`Property` procedures (built from the parsed
AST so `Optional`/`ParamArray`/default detail renders in VBE bracket form), then
`resolveRuntimeFunction`; runtime entries that are not valid explicit `Call`
targets suppress their call tip in that context using the same runtime metadata
as completion and diagnostics. The whole resolver is wrapped in try/catch so it never
disrupts editing, and signatures are never invented — an unknown callee yields
no tip. Verified host signatures live beside the object model in
`excelObjectModel.ts`; source-backed class member signatures are emitted by
`ProjectIndex.projectClassMembers()`.

**Developer documentation (XML doc-comments + external metadata)** —
`src/analyzer/docs/` adds Visual-Studio-style IntelliSense documentation. A
developer annotates a declaration with a `'''` XML doc-comment block (the same
tag vocabulary as C# `///`: `<summary>`, `<param name="...">`, `<returns>`,
`<remarks>`, `<example>`), and/or ships external `*.vbref.xml` metadata files
that document any symbol — including host members and runtime functions — using
`<member name="Module.Symbol">` entries with the **same** vocabulary.
`docModel.ts` is the host-agnostic model (`VbaDoc`) plus the hover/call-tip
Markdown renderers; `docComment.ts` parses inline blocks (and the shared XML
body); `externalDoc.ts` parses metadata files; `docRegistry.ts` resolves a name
(with optional qualifier) to a `VbaDoc`. Inline docs are attached to symbols in
`buildModuleSymbols.ts` (a backward scan over contiguous `'''` lines above each
member), including class-module properties and methods. Because VBA has no
source-level class declaration line, object-module docs use a documented header
convention: a contiguous `'''` block immediately above the first `Option`
directive attaches to the module root. `ProjectIndex.visibleTypeNames()` carries
module/type docs into type-name completion and hover, while the project
class-member surface carries member docs into `object.` completion, member hover,
and source-backed member-call signature help. Generated host reference metadata
also carries `VbaDoc` summaries and parameter notes into completion, hover, and
host member-call signature help. The VBA language configuration continues
`'''` lines when the user presses Enter inside a documentation block, and the
VBA-scoped smart Backspace/Tab commands clear an empty auto-continued `''' ` or
`' ` marker before deleting or indenting. Hover (`resolveHover`) and signature help
(`resolveSignatureHelp`) now carry a `documentation?` field (and per-parameter
docs for call tips), with the precedence **source inline comment > developer
external metadata > built-in host/reference metadata**. The vscode side
(`src/vbaDocMetadata.ts`, `DocMetadataLoader`) discovers metadata files anywhere
in the workspace via the `xlide.docs.metadataGlob` setting (default
`**/*.vbref.xml`), parses them into a live `DocRegistry`, and reloads on file
change; the registry is passed into the hover and signature-help contexts by
`src/vbaMemberCompletion.ts`. The full standard and usage paths live in
`docs/vba-doc-comments.md`.

**Active diagnostics engine** — `src/analyzer/diagnostics/` computes
high-confidence semantic problems directly from module text:

- `src/analyzer/diagnostics/ruleMetadata.ts` is the typed rule catalogue
  (`DIAGNOSTIC_RULES`): each rule carries a stable `code`, `title`,
  `defaultSeverity`, `category`, `vbeCompileEquivalent`, `diagnosticKind`,
  `source: 'XLIDE'`, an MS-VBAL `specReference`, and a `confidence`. Only
  high-confidence rules ship.
  The `undeclared-variable` rule ships for project-backed `Option Explicit`
  write/read positions: bare assignment and `Set` targets, assignment RHS and
  call-argument reads, control-flow block headers, member receivers, and indexed
  bases. It deliberately skips type-name, label, named-argument, and unresolved
  external-style call positions to avoid false positives. The arbitrary-expression
  `unknown-call` rule is deliberately absent for the same reason. The one cross-module call rule that does ship, `unknown-call`
  (`unknownCallStatement`), is restricted to the unambiguous call forms where the
  callee is a bare (non-member) identifier (see below).

Diagnostic severity policy:

| Diagnostic kind | Default surface | Rule metadata expectation |
| --- | --- | --- |
| Deterministic VBE compile failure | Error / red squiggly | `vbeCompileEquivalent: true` with a spec reference or oracle-verified behavior |
| Deterministic runtime failure | Error / red squiggly | `vbeCompileEquivalent: false`, `diagnosticKind: "deterministic-runtime-error"`, and focused runtime-oracle evidence or equivalent deterministic proof |
| Runtime risk or XLIDE-invalid guidance | Warning / yellow squiggly | `vbeCompileEquivalent: false`, explicit `category`, and tests that prove the analyzer has enough information |
| XLIDE-only guidance or style | Warning / yellow squiggly or lower | `vbeCompileEquivalent: false` and non-compile category such as `style`, `excel-host`, or `project-symbol` |
| Uncertain, incomplete while typing, host-dependent, or heuristic-only behavior | No diagnostic | No active rule until the behavior is spec-backed or oracle-verified |

- `src/analyzer/diagnostics/analyzeModule.ts` exposes
  `analyzeModule(source, opts)` returning `VbaDiagnostic[]` (code, message,
  severity, offset `span`). It reuses the lexer, parser, and symbol graph and
  implements the rules: unterminated string (odd-quote-count, escape-aware),
  duplicate procedure (Property Get/Let/Set may share a name), duplicate
  declaration in a flat procedure scope, duplicate module-level variable,
  assignment to a `Const` (bare `name =` only - excludes `.member`, indexing,
  `Set`, and comparisons), a configurable `Option Explicit`-missing
  reminder (silent on empty/attribute-only modules), an `unknown-call`
  ("Sub or Function not defined") rule, an `invalid-proc-header`
  ("Invalid procedure declaration") rule, an `unbalanced-parens` rule, and an
  `argument-count` ("Wrong number of arguments") rule. `bareCallStatementTarget`
  in `src/analyzer/call/callContext.ts` powers `unknown-call`: it accepts the
  three call forms whose callee is a bare (non-member) identifier - a lone
  identifier, a parenless call with arguments
  (`MsgBox "hi"`), and an explicit `Call name` - and flags the callee when the
  name resolves to no project procedure/Declare, runtime function/statement, host
  global, `Application` member, or in-scope declaration. It excludes member
  calls (`.`), labels (`:`), assignments (any top-level `=`), and the
  implicit-host-member form `Cells(1, 1)` / `Range("A1")` (a non-`Call`
  identifier immediately followed by `(`). `checkCallParens` powers call syntax
  diagnostics: explicit `Call` statements with arguments require parentheses,
  and VBE-oracle-verified standalone zero-argument calls cannot use empty
  parentheses unless they are prefixed with valid `Call` syntax or used in an expression.
  The rule uses the shared callable signature path for known same-module and
  exported standard-module procedures/Declares such as `myFunction()`, verified
  zero-argument runtime calls such as `DoEvents()`, and member/property
  statements such as `ThisWorkbook.CanCheckIn()`, `Application.Calculate()`, and
  `ActiveSheet.Range()`. Required-argument calls such as `MsgBox()` stay on the
  `argument-count` path. Non-empty standalone member/property calls such as
  `ActiveSheet.Range("A1")` and parenless member call statements such as
  `p.Save "ok"` are compile-accepted by VBE and stay on the
  signature-validation path. `invalid-explicit-call-target` uses the same
  runtime metadata to reject VBE-oracle-verified special cases such as
  `Call DoEvents` and `Call DoEvents()`, while allowing bare `DoEvents` and
  expression `DoEvents()`. `checkProcedureHeader` powers
  `invalid-proc-header`: after the procedure name in a `Sub`/`Function`/
  `Property` header, the only legal next token is `(` (or `As` for a
  `Function`/`Property Get`); any other token (e.g. the second word in
  `Sub My Sub`) is flagged. `checkUnbalancedParens` powers `unbalanced-parens`:
  it scans the module token stream tracking `(`/`)` depth, resets at each
  statement boundary (a newline or a depth-0 `:`), and flags a dangling `(` or
  an unmatched `)` - parentheses inside strings/comments/date-literals and
  `[bracketed]` names are distinct token kinds so they never miscount.
  `checkInvalidExpressionSyntax` covers narrow, high-confidence expression
  syntax errors such as consecutive non-unary binary operators (`***`) and
  statements ending with a binary operator; broader incomplete-expression
  recovery remains deferred to avoid noisy realtime diagnostics.
  `checkArgumentCount` powers `argument-count`: it validates call statements
  and expression calls against same-module, unique exported project, and
  verified runtime signatures with explicit parameter lists (for example
  `MsgBox()` is flagged because `Prompt` is required), and validates valid
  object member-call contexts when the shared member-completion binder resolves
  a known source-backed or host/reference signature, including parenthesized
  calls such as `Set wb = Workbooks.Open(...)` / `Range(Cell1, [Cell2])`,
  explicit `Call` statements, and parenless call statements such as
  `p.Save "ok"`; current class
  `Me.Member(...)` calls use that same path. It honours `Optional` (lowers the
  minimum) and `ParamArray` (removes the maximum), validates named-argument
  names against the parameters, and skips unresolved or ambiguous callees. The
  same known member signatures also feed argument type diagnostics when
  parameter types are explicit. The
  `argument-type-mismatch` and `assignment-type-mismatch` rules are red
  deterministic-runtime-error diagnostics when XLIDE can prove a literal cannot
  be coerced: focused Excel/VBE compile oracle cases confirm representative
  examples compile, and focused runtime oracle probes confirm they then raise
  runtime error 13. `argument-object-type-mismatch` is split out as a red
  compile-equivalent diagnostic after an oracle case confirmed `String` to
  `Object` is rejected by VBE Compile. `assignment-object-type-mismatch`
  covers `Set` assignments where both sides have deterministic object types,
  including project classes, source-backed object members, `Nothing`, and
  explicit `Implements` compatibility. These rules use only declared
  parameter/local types, `Function`/`Property Get` return-name types, curated
  runtime return metadata, same-module return types, known source-backed or
  host/reference member signatures, and deterministic literal/expression
  inference; unknown and `Variant` operands suppress diagnostics. For
  workbook-backed modules, the provider also passes a
  project-wide map of exported standard-module `Sub`/`Function`/`Declare`
  signatures, so argument count/type checks can cross module boundaries when
  the target is unambiguous. Ambiguous bare exported names stay silent, while
  `ModuleName.ProcedureName` resolves through the named standard module only;
  non-standard member cases stay silent until the binder can prove the target.
  `string-arithmetic-coercion` is a related red
  deterministic-runtime-error diagnostic for numeric contexts containing a
  provably nonnumeric string literal in an arithmetic expression; focused oracle
  cases confirm the representative expression compiles but raises runtime error
  13 when executed. Source-backed workbook class members and public fields feed
  the same assignment validator for member assignments such as
  `p.Age = "blah"` when the receiver and setter/write type are known; focused
  multi-module oracle cases confirm the nonnumeric string case raises runtime
  error 13 and the numeric-string control succeeds. `readonly-member-assignment`
  is a compile-equivalent
  diagnostic for source-backed project class properties whose member surface has
  no setter; focused oracle evidence rejects this as `Can't assign to read-only
  property`. `member-not-found` is another source-backed class-member rule: it
  fires only when a receiver resolves to an unambiguous and exhaustive
  `ProjectIndex.projectClassMembers()` surface, or a promoted exhaustive host
  surface, and the member name is absent. Plain class modules are
  source-exhaustive, including current-class `Me.Member` references;
  document modules and UserForms stay silent until their host/designer
  catalogues are complete enough to prove absence, except for known workbook
  host references such as `ThisWorkbook`/`Me` inside `ThisWorkbook`. Focused
  oracle evidence rejects unknown
  class property assignment and unknown class method calls with
  `Method or data member not found`, while known property and public-field
  controls compile. The diagnostic context disables completion-only
  `Set`-assignment refinement, so declared `Object`/`Variant` receivers remain
  late-bound even after assignments such as `Set obj = ActiveSheet`; VBE oracle
  compile controls accept unknown members there, and hard unknown-member,
  member-call, and member-assignment diagnostics stay silent.
  `unexpected-declaration-token` is
  a compile-equivalent
  declaration diagnostic for extra same-statement tokens after a complete
  `As` type name, such as `Dim s As String junk`; the representative `Dim`
  shape is backed by focused VBE oracle evidence, and broad fixed-length string
  behavior remains out of scope until its full grammar is verified. A separate
  `object-module-public-member` rule is module-kind-sensitive: in class,
  document, and UserForm modules it rejects explicit Public constants, arrays,
  fixed-length strings, user-defined Types, and Declare statements. Each branch
  is backed by focused VBE oracle evidence, while standard modules and non-Public
  declarations stay outside the rule.
  `event-handler-module-scope` is an information diagnostic for known
  workbook/worksheet event handler names declared outside the matching document
  module. It uses the same module-scoped event metadata as completion, and it is
  deliberately non-red because Excel treats those declarations as ordinary
  procedures rather than wired event handlers.
  `invalid-declaration-name` flags unbracketed MS-VBAL reserved identifiers in
  declaration-name positions while accepting bracketed `FOREIGN-NAME` forms such
  as `[In]`.
  `invalid-as-type-name` uses the shared type-position scanner and resolver
  for `As`, return, parameter, UDT field, `New`, `TypeOf ... Is`, and
  `Implements` positions. Resolved project/primitive/host types stay quiet;
  unresolved reserved identifiers, runtime functions, visible project
  declarations that are known not to be types, and ambiguous visible project
  type names are compile-equivalent errors. Broad unknown type names still wait
  for the external-reference story so referenced-library classes are not
  flagged prematurely.
  `set-required` fires when a plain assignment targets a known object variable,
  `Function`/`Property Get` return name, or source-backed object-valued member
  (`Property Set` or public field) that requires `Set`; `set-requires-object`
  fires when `Set` targets a known intrinsic scalar variable or source-backed
  scalar member. Both rules stay silent for `Variant`, unknown types, and
  ambiguous project members.
  `scalar-member-access` fires only when the receiver is a declared
  intrinsic scalar (`String`, numeric, `Boolean`, or `Date`); focused oracle
  cases show named scalar members are VBE Compile errors (`Invalid qualifier`),
  while a trailing scalar dot is a VBE Compile `Syntax error` after explicit
  focus retesting. Unknown, `Variant`, object-like, project class, and UDT
  receivers stay silent until the binder can prove more.
  `undeclared-variable` runs only when `Option Explicit` is present and the
  caller passes `knownIdentifiers` from
  `ProjectIndex.visibleIdentifierNames(moduleName)`, so unknown assignment
  targets, read references, member receivers, indexed bases, and block-header
  expressions become compile-equivalent `Variable not defined` diagnostics while
  public standard-module globals, enum members, document/UserForm code names,
  runtime functions/constants, host globals/enum constants, and `Application` members suppress false
  positives. Type-name positions, labels, named-argument labels, and arguments
  to unresolved external-style calls remain silent until the binder can prove
  those broader reference shapes. `unknown-call` runs only when the caller
  passes `knownProcedures` (the current module's visibility-filtered procedure
  names from `ProjectIndex.visibleProcedureNames(moduleName)`); without it that
  rule is skipped so a single module is never analysed in isolation. The whole analyzer is wrapped in
  try/catch so a parse hiccup returns `[]` and never breaks editing, and accepts
  `severities` overrides (including `'off'`) per rule.

This engine is merged with the structural block-balance linter
(`src/vbaLinter.ts`, which owns the "Missing End .../unexpected terminator"
family) inside `registerVbaDiagnostics` in `src/vbaLanguageProviders.ts`: both
run on open and debounced (300 ms) on every edit, on real `.vba` files and on
virtual `xlide-vba` module documents, with no save. Everything is computed from
the live editor text plus deterministic project context. For workbook-backed
documents the provider overlays the current live module text into a fresh
`ProjectIndex`, then passes visibility-filtered procedure names, visible bare
identifier names, cross-module standard-module signatures, and source-backed
project class-member surfaces into `analyzeModule`.
Settings `xlide.diagnostics.enabled` and `xlide.diagnostics.optionExplicit`
gate it and re-run open documents on change.

**Project-wide symbol graph** — `src/analyzer/symbols/` projects the parser AST
into a host-agnostic symbol model so XLIDE can offer document symbols, workspace
symbols, and go-to-definition over the whole workbook project rather than a
single module:

- `src/analyzer/symbols/symbolModel.ts` defines the `VbaSymbol` shape (name,
  kind, `nameSpan` for the identifier, `fullSpan` for the declaration,
  visibility, `asType`, exported `Attribute` metadata, and nested children).
- `src/analyzer/symbols/buildModuleSymbols.ts` walks one `ModuleNode` and emits
  a hierarchical module symbol (procedures with parameter/local children,
  `Type` with fields, `Enum` with members, module variables/consts, and
  `Declare`s with PtrSafe/Lib/Alias/parameter/return metadata).
  Identifier spans are located with the real lexer, so a `nameSpan` never lands
  inside a comment or string. Dotted exported attribute lines are mapped by the
  target before the dot, so `Attribute Value.VB_UserMemId = 0` attaches to the
  `Value` member while retaining the attribute source span.
- `src/analyzer/symbols/projectIndex.ts` is the `ProjectIndex` that aggregates
  modules and answers `documentSymbols`, `workspaceSymbols`, conservative
  `resolveDefinition` (locals/params -> same-module declarations -> exported
  declarations in other modules, with enum members inheriting project
  visibility from their containing `Enum`), `resolveQualifiedDefinition` (the exported
  member of a named module, for `Module.Member` references), `referenceScope`
  (the binding scope of a name for scope-restricted reference/rename search),
  `visibleProcedureNames` (same-module procedures/Declares plus exported
  standard-module procedures/Declares callable as bare identifiers from a given
  module),
  `visibleIdentifierNames` (same-module declarations, exported standard-module
  globals/types/enums and enum members, plus document/UserForm code names for
  Option Explicit diagnostics), `visibleIdentifierSymbols` (the source-backed
  declaration objects for the same visible bare-identifier surface, consumed by
  identifier completion),
  `visibleTypeNames` (class/document/UserForm module names plus visible
  `Type`/`Enum` declarations for `As`/`New` binding), `visibleNonTypeNames`
  (visible declarations that are known not to be type names, used only after
  type resolution fails), `resolveTypeDefinitions` (visible project type-name
  definitions with object modules resolving to the module start because the
  object type name is the VB component name),
  `projectClassMembers`
  (source-backed member surfaces with signatures, docs, definition spans,
  module-level `Implements` names, and default-member facts from
  `VB_UserMemId = 0` attributes), and
  `duplicateProcedures`. Cross-module visibility follows MS-VBAL: explicit
  `Public`/`Global` and default-`Public` procedures are exported;
  `Private`/`Dim`/`Friend` and unmodified module variables stay module-private.
  This index now drives the live `DefinitionProvider`, `ReferenceProvider`, and
  `RenameProvider` (see "Symbol intelligence").

**Type-name semantic coloring and hover** - TextMate grammar handles static VBA
tokens only, while workbook classes, document modules, UserForms, UDTs, and enums
are dynamic project symbols. `src/analyzer/semantic/typeSemanticTokens.ts`
therefore parses the live module and resolves actual type-name positions (`As
Person`, parameter types, return types, module/local variables, UDT fields, and
`New Person` expressions, `TypeOf value Is Person` tests, and module-level
`Implements Person` statements)
through the shared type resolver used by completion and hover: project-visible
names from `ProjectIndex.visibleTypeNames()`, VBA
primitive types, and Excel host object types. Semantic tokens mark primitives as
`type`, host/project object types as `class`, enums as `enum`, and UDTs as
`struct`; colliding project-visible names fall back to generic `type` for
coloring/hover and are flagged by diagnostics as ambiguous type-name errors.
The VS Code provider in
`src/vbaLanguageProviders.ts` overlays the live editor text for the current
module before resolving visible types, so `Dim customer As Person`, `Dim ws As
Worksheet`, and `Dim amount As Currency` color as soon as their type is known.
Unresolved names and non-type positions are ignored.

---

## Key design decisions

| Decision | Rationale |
|---|---|
| `FileSystemProvider` over `TextDocumentContentProvider` | Read/write virtual FS — Ctrl+S writes back with no custom save command |
| Long-lived Python process over per-call subprocess | Amortises ~200 ms Python startup across all requests |
| `cwd=python/` on spawn | Makes the `xlide` package importable without pip-installing it into the extension |
| pyOpenVBA for VBA, openpyxl for cells | pyOpenVBA owns the OVBA binary format; openpyxl reads/writes sheet data with `keep_vba=True` so macros are preserved |
| No COM, no Office | Works on Windows, macOS, Linux, WSL, and remote containers |
| Confirmation on write tools | Prevents AI agents from silently mutating production workbooks |

---

## Dependencies

| Package | Version | Role |
|---|---|---|
| `pyOpenVBA` | `>=3.0.1` | VBA module read/write (pure Python) |
| `openpyxl` | `>=3.1.0` | Excel cell data read/write |

TypeScript dev: `typescript`, `esbuild`, `@types/vscode`, `@types/node`.

---

## Files to keep up to date

| Change | Files to touch |
|---|---|
| New JSON-RPC method | `python/server.py`, `python/xlide/vba_io.py` or `excel_io.py`, `src/agentTools.ts` + `package.json` if exposed as LM tool, `docs/architecture.md` |
| New VS Code command | `src/commands.ts`, `package.json` (`contributes.commands`, `menus`), `docs/architecture.md` |
| New AI agent (LM) tool | `package.json` (`contributes.languageModelTools`), `src/agentTools.ts` (registration), `.github/copilot-instructions.md` (tool reference + workflow), `docs/architecture.md` |
| New workbook-wide lint behavior | `src/vbaWorkbookLint.ts` (shared core), `src/commands.ts` (`xlide.lintWorkbook` report formatting + clickable links), `src/agentTools.ts` (`xlide_lintWorkbook`), `package.json` (command/menu/LM tool), `.github/copilot-instructions.md`, `docs/architecture.md` |
| New Python source file | `python/xlide/__init__.py` (if re-exported), `docs/architecture.md` |
| Dependency added/removed | `python/requirements.txt`, `README.md` |
| New VBA language feature | `src/vbaSymbolIndex.ts` (parsing/index), `src/vbaLinter.ts` (structural analysis), `src/vbaLanguageProviders.ts` (provider), `syntaxes/vba.tmLanguage.json` (coloring), `language-configuration/vba-language-configuration.json` (brackets/indent/folding), `docs/architecture.md` |
| New analyzer grammar rule | `src/analyzer/**` (lexer/parser), matching fixtures in `tests/`, an MS-VBAL section cite in code, a row in `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New host object-model member/type/constant | `src/analyzer/host/excelObjectModel.ts` or generated `src/analyzer/host/excelReferenceMembers.ts` (metadata transcribed or generated with provenance), `tests/vbaMemberCompletion.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum table), `docs/architecture.md` |
| New host-member call signature | `src/analyzer/host/excelObjectModel.ts` (`memberSignatures` entry, transcribed + source-verified), `tests/vbaSignatureHelp.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum table), `docs/architecture.md` |
| New built-in VBA runtime function/statement | `src/analyzer/runtime/vbaRuntime.ts` (signature transcribed + source-verified), `tests/vbaRuntime.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New built-in VBA runtime constant | `src/analyzer/runtime/vbaRuntime.ts` (constant/type/value transcribed + source-verified), `tests/vbaRuntime.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New symbol-graph kind/resolution rule | `src/analyzer/symbols/**`, `tests/vbaSymbolGraph.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New definition/reference/rename scope rule | `src/analyzer/symbols/projectIndex.ts` (`resolveDefinition`/`resolveQualifiedDefinition`/`referenceScope`), `src/analyzer/index.ts` (barrel export), `src/vbaNavigation.ts` / `src/vbaLanguageProviders.ts` (provider wiring + span->range mapping), `tests/vbaSymbolGraph.test.ts`, `docs/architecture.md` |
| New active diagnostic rule | `src/analyzer/diagnostics/{ruleMetadata,analyzeModule}.ts` (rule + MS-VBAL `specReference`), `tests/vbaDiagnostics.test.ts`, `src/vbaLanguageProviders.ts` (provider merge + any new config), `package.json` (settings), `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New completion/hover resolver or rule | `src/analyzer/completion/**` or `src/analyzer/hover/**`, `src/analyzer/index.ts` (barrel export), `src/vbaMemberCompletion.ts` (provider wiring), matching `tests/vba*.test.ts`, `docs/architecture.md` |
| New signature-help rule/source | `src/analyzer/signature/signatureHelp.ts`, `src/analyzer/index.ts` (barrel export), `src/vbaMemberCompletion.ts` (`provideSignatureHelp` + `registerSignatureHelpProvider`), `tests/vbaSignatureHelp.test.ts`, `docs/architecture.md` |
| New doc-comment tag or metadata behavior | `src/analyzer/docs/**`, `src/analyzer/index.ts` (barrel export), `src/vbaDocMetadata.ts` (loader/glob), `src/vbaMemberCompletion.ts` (context wiring), `package.json` (`xlide.docs.*` settings), `tests/vbaDocComments.test.ts`, `docs/vba-doc-comments.md`, `docs/architecture.md` |
| Live Share RPC surface change | `src/liveShare.ts`, `docs/architecture.md` |
