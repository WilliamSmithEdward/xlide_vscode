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
    moduleDump.ts       Shared export/config logic for UI commands and AI tools
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
| `stat()` | Returns a synthetic `FileStat` (file, mtime=now) |
| All others | Throw `FileSystemError.NoPermissions` |

VS Code treats the file as fully editable — Ctrl+S triggers `writeFile` with no extra command needed.

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
- Name matches `ThisWorkbook` or `Sheet\d*` → `document`
- Anything else → `standard`

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
| `listModules` | `path` | — | `[{name, type}]` |
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

`moduleDump.ts` is the single source of truth for export/config behavior.

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

`xlide.dumpModulesToFolder` is a deprecated alias for `xlide.exportModulesToFolder` kept for backward compatibility.
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
| `DefinitionProvider` | Builds an AST `ProjectIndex` and resolves the identifier with scope-aware name resolution (`resolveDefinition`); honors a `Module.Member` qualifier via `resolveQualifiedDefinition`, and follows MS-VBAL visibility (locals shadow module members shadow exported cross-module declarations) |
| `ReferenceProvider` | Computes the identifier's binding scope (`ProjectIndex.referenceScope`), then runs a word-boundary search (skipping strings/comments) restricted to that scope — a procedure-local stays in its procedure, a `Private` member in its module, an exported symbol across the project minus privately-shadowing modules/locals; honors VS Code's include-declaration toggle |
| `RenameProvider` | Uses the same `referenceScope` to rewrite only the in-scope occurrences; `prepareRename`/`provideRenameEdits` refuse identifiers that do not resolve to a known declaration; VS Code applies the `WorkspaceEdit` and Ctrl+S persists each module through the virtual filesystem |
| Diagnostics | Debounced structural lint (`lintVbaSource`) flags unbalanced blocks — missing `End Sub`/`Next`/`Loop`/..., stray closers, and inner blocks left unclosed |
| Smart enter (auto-block) | Pressing Enter after a `Sub`/`Function`/`Property` header auto-inserts the matching `End ...` below and leaves the caret on the indented body line |

The `DefinitionProvider`, `ReferenceProvider`, and `RenameProvider` build a fresh
`ProjectIndex` (`src/analyzer/symbols/projectIndex.ts`) from the cached module
sources on each query; offset-based symbol spans are converted to editor ranges
in `vbaLanguageProviders.ts`. `VbaSymbolIndex` still backs the
`DocumentSymbolProvider` outline and the workbook-scoped source cache.

**Structural linting** — `src/vbaLinter.ts` is a pure, `vscode`-free module so it
is unit-tested directly (`tests/vbaLinter.test.ts`). It strips strings/comments,
joins `_` line continuations, then walks a block stack to detect imbalance. The
same module exports `detectProcOpener`/`isProcClosedAhead` used by the
smart-enter feature.

The index also subscribes to `onDidSaveTextDocument` for `xlide-vba://` URIs so
the cache stays in sync with user edits.

**Workbook-wide lint (command + agent tool)** — `src/vbaWorkbookLint.ts`
(`lintWorkbook`) is the shared core that loads every module from the workbook via
the Python bridge, builds a `ProjectIndex` so the unknown-call rule has the full
set of `knownProcedures`, then runs both diagnostic passes
(`lintVbaSource` + `analyzeModule`) per module and flattens their results into
1-based `{moduleName, moduleType, line, column, endColumn, severity, code,
message}` problems, sorted by module/line/column. The
`xlide.lintWorkbook` command (`src/commands.ts`, right-click "Lint All Modules in
Workbook" on a workbook tree node) prints a formatted, blank-line-padded report
to the XLIDE Output channel, switches focus to Output, and shows a summary
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
  and the `As`-type alias table. The reference corpus under `reference/` is used
  only as a transcription source; none of it is bundled or generated into the
  extension. LLM-generated member lists are never used; this is host metadata,
  not VBA grammar.
- `src/analyzer/host/hostModel.ts` exposes pure resolver functions over that
  metadata (`resolveHostGlobal`, `resolveHostAlias`, `getHostMembers`,
  `resolveMemberReturnType`).
- `src/analyzer/completion/memberAccess.ts` tokenizes the source up to the
  cursor, detects a member-access dot, walks the receiver chain (handling call
  parentheses for chaining like `ws.Range("A1").Offset(1, 0).`), resolves the
  root (`Me`, a host global, a worksheet code name, or a typed local/module
  variable found by parsing the module), follows member return types through the
  chain, and returns the filtered members.
- `src/vbaMemberCompletion.ts` is the VS Code `CompletionItemProvider` (trigger
  characters `.` and space). For member access it builds the project context
  from the workbook's module list (worksheet code names and the `Me` type for
  the current document module) via the Python bridge and renders the resolved
  members. In a declaration type position (after `As` / `As New`) it instead
  offers type-name completions via `src/analyzer/completion/typeCompletion.ts`
  (`resolveTypeCompletions`): VBA built-in data types, the Excel host types, and
  project-defined types — user `Type`s/`Enum`s in the current module, public
  (non-`Private`) `Type`s/`Enum`s read from the workbook's other modules (via the
  bridge `readModule` call, cached per workbook with a short TTL), plus
  class/UserForm module names from the workbook. When the cursor is on a bare
  identifier (statement/expression position, not after `.` or `As`) it offers
  identifier completions via `src/analyzer/completion/identifierCompletion.ts`
  (`resolveIdentifierCompletions`): host-injected globals (`ThisWorkbook`,
  `ActiveSheet`, `Application`, ...), worksheet/document code names, the
  user's in-scope declarations (parameters, locals, module variables/constants,
  procedures, enums and their members, user types), and built-in VBA runtime
  functions (`MsgBox`, `Left`, `CLng`, `RGB`, ...) from the runtime metadata.
- The same `src/vbaMemberCompletion.ts` class also registers a VS Code
  `HoverProvider`. It delegates to `src/analyzer/hover/resolveHover.ts`
  (`resolveHover`), a pure resolver that describes the identifier under the
  cursor: `receiver.member` host members (reusing the exported
  `resolveReceiverTypeAt` from `memberAccess.ts`), host globals, worksheet code
  names, user declarations from the live module symbol graph (procedure
  signatures with parameters and return type, variables/parameters/constants
  with their `As` type, enums and members, user types and fields), and built-in
  VBA runtime functions, annotated with the declaring module and visibility.
  Unknown members are never guessed. Built-ins resolve last so a user
  declaration of the same name shadows the built-in.

**Built-in VBA runtime metadata** — `src/analyzer/runtime/vbaRuntime.ts` is a
curated, verified subset of the intrinsic VBA runtime library (~85 functions and
statements: `MsgBox`, `InputBox`, the `C*` conversions, string/date/math
helpers, `Array`, `UBound`, `RGB`, ...). Each `VbaRuntimeFunction` carries a
canonical `signature`, optional `returns`, a `kind` of `function | statement`,
and `source: 'verified'`. Signatures are transcribed from
learn.microsoft.com/office/vba/language and MS-VBAL, never LLM-invented. Names
that collide with intrinsic data types (`Date`, `Time`, `String`, `Error`) are
deliberately omitted so a type in an `As` position is never read as a function.
Like the host model, this is a typed TS module (not the JSON file the roadmap
originally suggested) for compile-time checking. `resolveRuntimeFunction(name)`
resolves case-insensitively; `VBA_RUNTIME_FUNCTIONS` is the full list consumed
by hover and identifier completion.

**Signature help (parameter info)** — `src/analyzer/signature/signatureHelp.ts`
computes the VBE call tip from module text alone. `resolveSignatureHelp(source,
offset, ctx)` returns a `SignatureInfo` (the signature `label`, ordered
`parameters`, and the `activeParameter` index) or `undefined`. It tokenizes the
prefix up to the caret, maintains a paren-frame stack to find the innermost
enclosing *call* paren (a `(` directly preceded by an identifier; grouping/index
parens are skipped), counting top-level commas for the active parameter; when no
call paren is open it falls back to a conservative *parenless call statement*
detector (`Workbooks.Open "file", `) that bails on statement keywords, file-I/O
starters, and top-level `=` assignments. The callee's signature is sourced, in
order, from: host-member signatures (`resolveHostMemberSignature`, backed by the
verified `EXCEL_OBJECT_MODEL.memberSignatures` table — e.g. the full
`Workbooks.Open(...)`), user `Sub`/`Function`/`Property` procedures (built from
the parsed AST so `Optional`/`ParamArray`/default detail renders in VBE bracket
form), then `resolveRuntimeFunction`. The whole resolver is wrapped in try/catch
so it never disrupts editing, and signatures are never invented — an unknown
callee yields no tip. Host-member receiver types are resolved by reusing
`resolveReceiverTypeAt` from the member-completion layer, so chained receivers
(`ActiveSheet.Range("A1").Offset(`) work. Verified host signatures live beside
the object model in `excelObjectModel.ts`; `resolveHostMemberSignature` in
`hostModel.ts` looks them up case-insensitively.

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
member). Hover (`resolveHover`) and signature help (`resolveSignatureHelp`) now
carry a `documentation?` field (and per-parameter docs for call tips), with the
precedence **inline comment > external metadata > curated library** — i.e.
developer-defined metadata overrides the built-in library. The vscode side
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
  `defaultSeverity`, `category`, `vbeCompileEquivalent`, `source: 'XLIDE'`, an
  MS-VBAL `specReference`, and a `confidence`. Only high-confidence rules ship.
  The broad `undeclared-variable`
  rule and the arbitrary-expression `unknown-call` rule are deliberately absent —
  they would need an expression binder plus a complete host catalogue and would
  otherwise produce false positives, which the project's no-false-positive rule
  forbids. The one cross-module rule that does ship, `unknown-call`
  (`unknownCallStatement`), is restricted to the unambiguous call forms where the
  callee is a bare (non-member) identifier (see below).

Diagnostic severity policy:

| Diagnostic kind | Default surface | Rule metadata expectation |
| --- | --- | --- |
| Deterministic VBE compile failure | Error / red squiggly | `vbeCompileEquivalent: true` with a spec reference or oracle-verified behavior |
| Deterministic XLIDE-invalid or runtime-risk rule | Error or warning, depending on blast radius | `vbeCompileEquivalent: false`, explicit `category`, and tests that prove the analyzer has enough information |
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
  `argument-count` ("Wrong number of arguments") rule. `callStatementTarget`
  powers `unknown-call`: it accepts the three call forms whose callee is a bare
  (non-member) identifier - a lone identifier, a parenless call with arguments
  (`MsgBox "hi"`), and an explicit `Call name` - and flags the callee when the
  name resolves to no project procedure, runtime function/statement, host
  global, `Application` member, or in-scope declaration. It excludes member
  calls (`.`), labels (`:`), assignments (any top-level `=`), and the
  implicit-host-member form `Cells(1, 1)` / `Range("A1")` (a non-`Call`
  identifier immediately followed by `(`). `checkProcedureHeader` powers
  `invalid-proc-header`: after the procedure name in a `Sub`/`Function`/
  `Property` header, the only legal next token is `(` (or `As` for a
  `Function`/`Property Get`); any other token (e.g. the second word in
  `Sub My Sub`) is flagged. `checkUnbalancedParens` powers `unbalanced-parens`:
  it scans the module token stream tracking `(`/`)` depth, resets at each
  statement boundary (a newline or a depth-0 `:`), and flags a dangling `(` or
  an unmatched `)` - parentheses inside strings/comments/date-literals and
  `[bracketed]` names are distinct token kinds so they never miscount.
  `checkArgumentCount` powers `argument-count`: it validates a call statement
  against the parameter list of a Sub/Function defined in the *same* module (the
  only place the AST gives a ground-truth signature), reusing `extractCall`
  (built on `callStatementTarget`) to pull the callee and its top-level argument
  slots; it honours `Optional` (lowers the minimum) and `ParamArray` (removes
  the maximum), validates named-argument names against the parameters, and skips
  host/runtime/cross-module callees plus any duplicated/ambiguous name. The
  `unknown-call` rule runs only when the caller
  passes `knownProcedures` (the project-wide procedure-name set from
  `ProjectIndex.procedureNames()`); without it that rule is skipped so a single
  module is never analysed in isolation. The whole analyzer is wrapped in
  try/catch so a parse hiccup returns `[]` and never breaks editing, and accepts
  `severities` overrides (including `'off'`) per rule.

This engine is merged with the structural block-balance linter
(`src/vbaLinter.ts`, which owns the "Missing End .../unexpected terminator"
family) inside `registerVbaDiagnostics` in `src/vbaLanguageProviders.ts`: both
run on open and debounced (300 ms) on every edit, on real `.vba` files and on
virtual `xlide-vba` module documents, with no save. Everything is computed from
the live editor text; the only cross-module input is the project procedure-name
set, which the provider reads from the `VbaSymbolIndex` cache (`getAllModules`,
a Python round-trip only on the first, uncached load) and passes to
`analyzeModule` as `knownProcedures` for the bare-call rule.
Settings `xlide.diagnostics.enabled` and `xlide.diagnostics.optionExplicit`
gate it and re-run open documents on change.

**Project-wide symbol graph** — `src/analyzer/symbols/` projects the parser AST
into a host-agnostic symbol model so XLIDE can offer document symbols, workspace
symbols, and go-to-definition over the whole workbook project rather than a
single module:

- `src/analyzer/symbols/symbolModel.ts` defines the `VbaSymbol` shape (name,
  kind, `nameSpan` for the identifier, `fullSpan` for the declaration,
  visibility, `asType`, and nested children).
- `src/analyzer/symbols/buildModuleSymbols.ts` walks one `ModuleNode` and emits
  a hierarchical module symbol (procedures with parameter/local children,
  `Type` with fields, `Enum` with members, module variables/consts, `Declare`s).
  Identifier spans are located with the real lexer, so a `nameSpan` never lands
  inside a comment or string.
- `src/analyzer/symbols/projectIndex.ts` is the `ProjectIndex` that aggregates
  modules and answers `documentSymbols`, `workspaceSymbols`, conservative
  `resolveDefinition` (locals/params -> same-module declarations -> exported
  declarations in other modules), `resolveQualifiedDefinition` (the exported
  member of a named module, for `Module.Member` references), `referenceScope`
  (the binding scope of a name for scope-restricted reference/rename search),
  and `duplicateProcedures`. Cross-module visibility follows MS-VBAL: explicit
  `Public`/`Global` and default-`Public` procedures are exported;
  `Private`/`Dim`/`Friend` and unmodified module variables stay module-private.
  This index now drives the live `DefinitionProvider`, `ReferenceProvider`, and
  `RenameProvider` (see "Symbol intelligence").

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
| New host object-model member/type | `src/analyzer/host/excelObjectModel.ts` (member transcribed + source-verified), `tests/vbaMemberCompletion.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum table), `docs/architecture.md` |
| New host-member call signature | `src/analyzer/host/excelObjectModel.ts` (`memberSignatures` entry, transcribed + source-verified), `tests/vbaSignatureHelp.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum table), `docs/architecture.md` |
| New built-in VBA runtime function/statement | `src/analyzer/runtime/vbaRuntime.ts` (signature transcribed + source-verified), `tests/vbaRuntime.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New symbol-graph kind/resolution rule | `src/analyzer/symbols/**`, `tests/vbaSymbolGraph.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New definition/reference/rename scope rule | `src/analyzer/symbols/projectIndex.ts` (`resolveDefinition`/`resolveQualifiedDefinition`/`referenceScope`), `src/analyzer/index.ts` (barrel export), `src/vbaLanguageProviders.ts` (provider wiring + span->range mapping), `tests/vbaSymbolGraph.test.ts`, `docs/architecture.md` |
| New active diagnostic rule | `src/analyzer/diagnostics/{ruleMetadata,analyzeModule}.ts` (rule + MS-VBAL `specReference`), `tests/vbaDiagnostics.test.ts`, `src/vbaLanguageProviders.ts` (provider merge + any new config), `package.json` (settings), `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New completion/hover resolver or rule | `src/analyzer/completion/**` or `src/analyzer/hover/**`, `src/analyzer/index.ts` (barrel export), `src/vbaMemberCompletion.ts` (provider wiring), matching `tests/vba*.test.ts`, `docs/architecture.md` |
| New signature-help rule/source | `src/analyzer/signature/signatureHelp.ts`, `src/analyzer/index.ts` (barrel export), `src/vbaMemberCompletion.ts` (`provideSignatureHelp` + `registerSignatureHelpProvider`), `tests/vbaSignatureHelp.test.ts`, `docs/architecture.md` |
| New doc-comment tag or metadata behavior | `src/analyzer/docs/**`, `src/analyzer/index.ts` (barrel export), `src/vbaDocMetadata.ts` (loader/glob), `src/vbaMemberCompletion.ts` (context wiring), `package.json` (`xlide.docs.*` settings), `tests/vbaDocComments.test.ts`, `docs/vba-doc-comments.md`, `docs/architecture.md` |
| Live Share RPC surface change | `src/liveShare.ts`, `docs/architecture.md` |
