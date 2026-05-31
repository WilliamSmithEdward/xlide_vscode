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
| `DefinitionProvider` | Resolves an identifier across all modules in the workbook; honors `Module.Member` qualifiers and `Private` visibility |
| `ReferenceProvider` | Word-boundary search across all modules, skipping string literals and apostrophe comments |
| `RenameProvider` | `prepareRename` checks the identifier is a known procedure; `provideRenameEdits` returns a `WorkspaceEdit` that rewrites every module; VS Code applies the edit and Ctrl+S persists each module through the virtual filesystem |
| Diagnostics | Debounced structural lint (`lintVbaSource`) flags unbalanced blocks — missing `End Sub`/`Next`/`Loop`/..., stray closers, and inner blocks left unclosed |
| Smart enter (auto-block) | Pressing Enter after a `Sub`/`Function`/`Property` header auto-inserts the matching `End ...` below and leaves the caret on the indented body line |

The `ReferenceProvider` excludes the procedure declaration token itself, so
"Find All References" returns only call sites, not the definition.

**Structural linting** — `src/vbaLinter.ts` is a pure, `vscode`-free module so it
is unit-tested directly (`tests/vbaLinter.test.ts`). It strips strings/comments,
joins `_` line continuations, then walks a block stack to detect imbalance. The
same module exports `detectProcOpener`/`isProcClosedAhead` used by the
smart-enter feature.

The index also subscribes to `onDidSaveTextDocument` for `xlide-vba://` URIs so
the cache stays in sync with user edits.

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
  Excel automation object model (Application/Workbook/Worksheet/Range), with each
  type's properties and methods transcribed from the official Office VBA
  object-model reference (`learn.microsoft.com/office/vba/api/excel.*`). It also
  holds the host-global table (`ThisWorkbook` -> `Excel.Workbook`, `Application`
  -> `Excel.Application`, ...) and the `As`-type alias table. LLM-generated
  member lists are never used; this is host metadata, not VBA grammar.
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
  declarations in other modules), and `duplicateProcedures`. Cross-module
  visibility follows MS-VBAL: explicit `Public`/`Global` and default-`Public`
  procedures are exported; `Private`/`Dim`/`Friend` and unmodified module
  variables stay module-private.

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
| New Python source file | `python/xlide/__init__.py` (if re-exported), `docs/architecture.md` |
| Dependency added/removed | `python/requirements.txt`, `README.md` |
| New VBA language feature | `src/vbaSymbolIndex.ts` (parsing/index), `src/vbaLinter.ts` (structural analysis), `src/vbaLanguageProviders.ts` (provider), `syntaxes/vba.tmLanguage.json` (coloring), `language-configuration/vba-language-configuration.json` (brackets/indent/folding), `docs/architecture.md` |
| New analyzer grammar rule | `src/analyzer/**` (lexer/parser), matching fixtures in `tests/`, an MS-VBAL section cite in code, a row in `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New host object-model member/type | `src/analyzer/host/excelObjectModel.ts` (member transcribed + source-verified), `tests/vbaMemberCompletion.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum table), `docs/architecture.md` |
| New built-in VBA runtime function/statement | `src/analyzer/runtime/vbaRuntime.ts` (signature transcribed + source-verified), `tests/vbaRuntime.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New symbol-graph kind/resolution rule | `src/analyzer/symbols/**`, `tests/vbaSymbolGraph.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New completion/hover resolver or rule | `src/analyzer/completion/**` or `src/analyzer/hover/**`, `src/analyzer/index.ts` (barrel export), `src/vbaMemberCompletion.ts` (provider wiring), matching `tests/vba*.test.ts`, `docs/architecture.md` |
| Live Share RPC surface change | `src/liveShare.ts`, `docs/architecture.md` |
