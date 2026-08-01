# XLIDE – Architecture

## Overview

XLIDE is a VS Code extension that turns Excel macro files (`.xlsm`, `.xlsb`, `.xlam`) into first-class editable documents. VBA modules open in the editor like normal source files, Ctrl+S writes them back into the workbook, and 18 VS Code language model tools expose workbook-aware operations to Copilot and compatible agents.

Everything runs in the extension host. There is no backend process, no interpreter, and no third-party library between XLIDE and the workbook file:

```
VS Code Extension (TypeScript)
        |
        |  in-process function calls
        |
Workbook engine  (src/vba/**, no COM, no Office install required)
        |-- cfb.ts            [MS-CFB] compound file reader / writer
        |-- ovba.ts           [MS-OVBA] compression + decompression
        |-- vbaProject.ts     dir stream, module streams, PROJECT stream
        |-- zip.ts / xlsx.ts  OOXML package + worksheet cell data
        +-- workbookService.ts  the operations the extension calls
```

---

## Repository layout

```
xlide_vscode/
  src/
    extension.ts        Activation entry point — registers all providers and commands
    workbookEngine.ts   WorkbookEngine class — in-process dispatcher for every workbook operation
    xlsmExplorer.ts     XlsmExplorer — TreeDataProvider for the XLIDE workbook tree in VS Code Explorer
    xlideSidebar.ts     Polished XLIDE Activity Bar/sidebar WebviewView
    xlideSidebarModel.ts Pure model for sidebar status/action/configuration sections
    xlideFileSystem.ts  XlideFileSystemProvider — virtual xlide-vba:// filesystem
    commands.ts         Thin command composition root; handlers live in the per-domain modules under src/commands/
    commands/           Per-domain command modules: analysisCommands.ts, moduleSyncCommands.ts, vbaTestCommands.ts, workbookCrudCommands.ts, supportBundleCommands.ts, miscCommands.ts, shared.ts (CommandDeps + cross-domain helpers)
    agentTools.ts       LanguageModelTool registrations for AI agent use
    workbookModuleOperations.ts  Shared workbook module write/rename/delete service used by both UI commands and agent tools
    moduleExport.ts     Shared module export logic for UI commands and AI tools
    workbookSettings.ts Strict workbook settings sidecar path, schema validation, and persistence
    workbookModuleSyncSettings.ts Effective workbook import/export sync settings and provenance
    globalSettings.ts  Machine-scoped VS Code XLIDE setting validation, normalization, and provenance
    liveShare.ts        LiveShareIntegration — host/guest Live Share bridge over the VSLS service API
    statusBar.ts        XlideStatusBar — two status bar items (active module, Live Share guest indicator)
    vsls.d.ts           Ambient type declarations for the VS Code Live Share extension API
    vbaSymbolIndex.ts   VbaSymbolIndex — workbook-scoped cache of VBA module sources
    vbaLanguageProviders.ts  Composition root that registers the vba language subsystems; implementations live in vbaLiveDiagnostics.ts, vbaCompletionProvider.ts, vbaHoverSignatureProvider.ts, vbaNavigationProviders.ts, vbaSemanticTokensProvider.ts, vbaCodeActions.ts, and vbaTypingAutomation.ts
    vbaSourceScan.ts    Pure shared VBA source-scan utilities — stripVba, line start offsets, logical lines, identifier search/validation (no vscode dependency)
    vbaStructuralDiagnostics.ts     Pure structural block-balance analysis (analyzeVbaStructure) and the block opener/closer grammar (no vscode dependency)
    vbaSmartEnter.ts    Pure Smart Enter block completion, With-member continuation, loop-iterator sync, and procedure-header paren repair (no vscode dependency)
    vbaSmartBlockSnippets.ts        Shared smart-block snippet catalogue projected into keyword completion and conformance tests (no vscode dependency)
    vbaModuleAnalysis.ts    Shared module analysis core used by live diagnostics, current-module analysis, and workbook analysis; merges structural analysis, semantic analysis, and analysis suppression directives
    vbaOpenDocuments.ts Shared same-workbook open-document source overlay helper for editor-backed project analysis
    vbaProjectAnalysis.ts  Shared ProjectIndex construction and analyzer-option derivation for project-aware analysis, diagnostics, semantic tokens, completion, and hover surfaces
    vbaProjectIndexService.ts  VbaProjectIndexService — one shared incremental ProjectIndex per workbook (cold build, then per-module folds), read by diagnostics, completion/hover/signature help, navigation, and semantic tokens
    vbaWorkbookAnalysis.ts  Workbook-wide analysis core (analyzeWorkbook) reused by commands and the xlide_analyzeWorkbook agent tool; flattens vbaModuleAnalysis results into 1-based problems with diagnostic metadata and summary counts
    workbookAnalysisResultsModel.ts  Pure workbook analysis results view model and copy-report formatting shared by analysis result UI tests and the webview
    workbookAnalysisWebview.ts  Dedicated VS Code webview panel for current-module/workbook analysis results, filters, workbook-scoped analysis settings provenance/reset controls, copy actions, and click-through navigation
    analyzer/
      lexer/
        keywordTable.ts MS-VBAL 3.3.5.2 reserved-identifier + contextual keyword tables with canonical casing
        tokenKinds.ts   TokenKind/Trivia/VbaToken types and WSC helpers (MS-VBAL 3.3)
        trivia.ts       Leading whitespace / line-continuation scanner (MS-VBAL 3.2.2)
        tokenize.ts     Loss-aware, round-trippable VBA tokenizer (MS-VBAL 3.3.1-3.3.5, 3.4)
        strippedLines.ts Lexer-derived stripped-line substrate (strings/comments blanked, columns preserved) consumed by Smart Enter and keyword completion (audit #74)
      parser/
        nodes.ts        AST node types + spans + ParseDiagnostic (MS-VBAL 4.2/5.x)
        parserState.ts  Logical-statement splitter + statement cursor (MS-VBAL 3.3.1 EOS)
        parseModule.ts  Error-tolerant module parser -> ModuleNode AST (MS-VBAL 5.x)
      codeActions/
        diagnosticCodeActions.ts  Pure analyzer quick-fix resolver keyed by diagnostic rule code
      diagnostics/
        analyzeModule.ts  Analysis pass entry point; parses/lexes once per pass and runs the rule registry over a shared statement walk
        registry.ts     Ordered diagnostic rule registry assembling the per-family rule modules
        rules/          Per-family rule modules: lexical, duplicates, declarations, callArity, argumentTypes, assignments, objectState, undeclared, arrays, expressions, moduleKind, controlFlow, runtimeValues, plus shared.ts cross-family helpers
        ruleMetadata.ts Typed rule catalogue: stable codes, severities, categories, VBE equivalence, MS-VBAL spec references
        analysisContext.ts  Per-pass shared contracts and the memoized statement-token cache
        walker.ts       Shared procedure/statement traversal utilities
        typeInference.ts  Type-inference engine with per-pass memoized signature tables and type environments
        callExtraction.ts Call extraction and arity validation helpers
        constExpr.ts    Constant-expression collection and evaluation
        dataflow.ts     Shared straight-line dataflow walker
        analysisSuppressions.ts  Shared XLIDE analysis suppression directive scanner/filter for live diagnostics and workbook analysis
      semantic/
        typeSemanticTokens.ts  Pure resolver for type-name semantic tokens and hover in declaration/New positions
      index.ts          Public, vscode-free analyzer surface (lexer + parser)

    webview/            Shared webview scaffold: templates.ts (assets/webview loader), html.ts, page.ts, refresh.ts, panelRegistry.ts, styles.ts

    vba/
      cfb.ts            [MS-CFB] compound file binary reader/writer (canonical rebuild on save)
      ovba.ts           [MS-OVBA] run-length compression/decompression, byte-identical to the spec's reference decoder
      vbaProject.ts     vbaProject.bin: dir stream parse/serialize, module add/rename/delete, PROJECT stream, signature detection
      zip.ts            ZIP reader/writer preserving untouched entries' original compressed bytes
      xlsx.ts           OOXML package: sheet enumeration, cell/formula read, cell write, shared strings
      workbookService.ts  The operation layer the extension calls: module CRUD, protection info, sheets, cells, atomic workbook writes

  assets/
    webview/            Externalized webview template assets (HTML/CSS/JS) for workbookAnalysis, moduleSync, vbaTests, vbaTestResults, and globalSettings panels
    testhost/           Externalized VBA test host sources loaded at runtime: XlideTestModalWatcher.cs, run-vba-tests.ps1

  docs/
    architecture.md     This file
    roadmap_version_2.x.md    Closed v2 launch-hardening roadmap
    roadmap_version_2.1.0.md  Completed v2.1.0 red-squiggle closeout
    roadmap_version_2.2.0.md  Static-analysis completeness roadmap (superseded by 2.4.0)
    roadmap_version_2.4.0.md  Completed static-analysis completeness sprint
    static_analysis_completeness_2.4.0.md  v2.4.0 completeness report (release gate)
    roadmap_version_2.5.0.md  Completed: expression binder + syntax-corpus completeness
    static_analysis_completeness_2.5.0.md  v2.5.0 completeness report (release gate)
    roadmap_version_2.6.0.md  Next: deferred binder families + performance backlog

  package.json          Extension manifest, contributes, LM tool declarations
  tsconfig.json         Strict TypeScript config (module: Node16)
  esbuild.js            Bundle script — produces out/extension.js
  .vscode/
    launch.json         F5 Extension Development Host config
    tasks.json          Default build task (npm run watch)
    settings.json       Workspace editor settings
```

---

## Virtual filesystem — `xlide-vba://`

Clicking a module in the XLIDE Explorer workbook tree opens it under the custom
scheme:

```
xlide-vba:///C:/path/to/workbook.xlsm/Module1.bas
```

`XlideFileSystemProvider` implements `vscode.FileSystemProvider`:

| Method | Action |
|---|---|
| `readFile(uri)` | Calls `readModule` on the workbook engine; returns UTF-8 bytes |
| `writeFile(uri, content)` | Calls `writeModule`; saves the .xlsm in place |
| `stat()` | Returns a `FileStat` whose `mtime` derives from the real workbook file mtime; when the workbook file changes out-of-band (for example a concurrent Excel VBE edit), module mtimes move forward so VS Code's save-conflict detection triggers instead of silently overwriting the newer workbook copy |
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

### Workbook identity and siloing

Each local workbook is a separate VBA project boundary. The `xlide-vba://` URI
embeds the workbook path, and workbook-aware services must key cache entries,
project indexes, diagnostics, open-document overlays, reference/rename scopes,
and mutation commands by that workbook identity before considering module names.
The shared helpers in `src/xlideFileSystem.ts` (`workbookIdentityKey`,
`sameWorkbookPath`, and `moduleIdentityKey`) define that boundary. This is
deliberate: two open workbooks can both contain `Module1`, `Person`, or
`ThisWorkbook`, and those names must not share completion, hover, diagnostics,
references, rename edits, analysis context, or live source overlays. On Windows,
path comparisons are case-insensitive; on other platforms they are exact unless
a caller has explicitly normalized the path. VBA module-name identity is
case-insensitive inside the workbook boundary.

Cross-workbook actions, such as copying modules from one workbook to another,
must be exposed as explicit user-chosen workflows with source/destination
selection, preview, conflict handling, and safety checks. They must not happen
implicitly through project analysis or editor IntelliSense.

Same-workbook live source overlays are centralized in `src/vbaOpenDocuments.ts`.
That helper scans open `xlide-vba://` documents, decodes the workbook/module
identity through the same URI helpers, and replaces cached module source only
when the open editor belongs to the requested workbook. Completion, hover,
signature help, live diagnostics, semantic tokens, definition/reference/rename,
current-module analysis, and tree-level class rename all use that overlay path
instead of carrying local workspace-document scans.
`tests/vbaMultiWorkbookIsolation.test.ts` locks the shared overlay and
project-analysis inputs with two same-named open workbooks so provider surfaces
stay workbook-siloed by construction.
Larger project-analysis regressions should use machine-readable workbook
fixtures under `tests/fixtures/vbaProjects/*.json`, loaded by
`tests/helpers/vbaProjectFixtures.ts` and asserted by
`tests/vbaProjectWorkbookFixtures.test.ts`. Those fixtures carry module source
plus expected project context, member/type/identifier completion, diagnostics,
semantic-token, hover, and signature-help behavior so multi-surface workbook
scenarios do not need bespoke one-off test wiring.

---

## Sidebar tree — `XlsmExplorer`

`TreeDataProvider<XlideNode>` with three levels:

| Level | Node kind | Children source |
|---|---|---|
| 0 | `xlsm` — one per file found by `findFiles('**/*.{xlsm,xlsb,xlam}')` | modules |
| 1 | `module` — name + type (standard / class / document) | subs |
| 2 | `sub` — procedure name, kind, 1-based line number | none |

Clicking a `module` node opens the module via `xlide.openModule`. Clicking a `sub` node opens the module and moves the cursor to that line.

Module type is inferred from the VBA source and name (`classifyModuleType` in
`src/vba/workbookService.ts`, mirrored for sync planning in
`src/moduleSyncPlan.ts`):
- `Attribute VB_Base` carries two GUIDs → `userform`
- `Attribute VB_Base` carries a Workbook/Worksheet/Chart CLSID → `document`
- Name is `ThisWorkbook` or matches a localized sheet pattern (`Sheet\d*`,
  `Feuil\d*`, `Hoja\d*`, `Tabelle\d*`, `Foglio\d*`, `Planilha\d*`) → `document`
- Anything else → `standard`; the dir stream's MODULETYPE record upgrades
  non-standard modules to `class`

`Attribute VB_PredeclaredId = True` alone is deliberately NOT treated as a
document marker — predeclared singleton-style classes (e.g. stdVBA's
`stdArray`/`stdLambda`) carry it too.

When the hidden `VB_Base` attribute is available, `listModules` also returns an
optional `documentType` for document modules (`workbook`, `worksheet`, or
`chart`). The broad `type` field remains `document` so existing import/export
behavior has one module-kind contract, while language-service features can use
the subtype when they need workbook-vs-worksheet-vs-chart semantics.

---

## Workbook engine — `WorkbookEngine`

`WorkbookEngine.call(method, params)` dispatches straight into `src/vba/workbookService.ts` on the extension host thread. There is no process to start, probe, recover, or configure: the first call after activation is as cheap as the hundredth, and a workspace with no Excel workbooks never opens one.

The call shape is deliberately the same request/response surface the extension used when the work lived behind a JSON-RPC child process, so every caller — explorer, virtual filesystem, analysis, commands, agent tools, Live Share — is unchanged. Failures reject with `BridgeError` carrying a JSON-RPC-style code (`-32601` unknown method, `-32602` bad params, `-32000` operation failed).

**Write safety.** Mutating operations rebuild the whole container and land through `atomicWrite`: the new bytes go to a temp file beside the target and are then renamed over it, so a crash mid-write cannot leave a half-written workbook. Per [MS-OVBA], a mutating save also drops every `__SRP_*` performance-cache stream and clears the `_VBA_PROJECT` cache body — stale compiled p-code for a module set that no longer matches is what crashes Excel on open. A non-mutating save leaves both untouched, because the caches still describe the project exactly.

---

## Windows Excel COM behavior

The commands `xlide.openWorkbook` and `xlide.runMacroAtCursor` use PowerShell COM automation on Windows.

Setting:

- `xlide.attachToRunningExcel` (default `true`)
  - `true`: tries to attach to a running `Excel.Application` and reuse an already-open workbook (matched by full path or workbook name) before opening.
  - `false`: always opens through a new COM-created Excel application path.

---

## Engine methods

| Method | Required params | Optional params | Returns |
|---|---|---|---|
| `listModules` | `path` | — | `[{name, type, documentType?}]` |
| `listSubs` | `path`, `module` | — | `[{name, kind, line}]` |
| `readModule` | `path`, `module` | — | `{source}` |
| `readModules` | `path` | `full` (bool, default `false`) | `[{name, type, documentType?, source}]` — batch read used by analysis, sync plans, and test discovery |
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

Failures reject with a `BridgeError` whose `code` follows the JSON-RPC convention (`-32601`, `-32602`, `-32000`).

---

## Protected & signed workbooks

A VBA project password protects the project from being *opened in the VBE*, not the bytes on disk, so `writeModule`, `renameModule`, and `deleteModule` edit password-locked projects in place. Each reports `signatureDropped: true` when the workbook carried a digital signature, because rewriting module streams invalidates the signature the workbook was signed with.

On the TypeScript side, `notifySignatureDropped(filePath, signatureDropped)` in `xlideFileSystem.ts` shows a one-time-per-workbook warning when a signature is invalidated. `writeFile` and the three write agent tools/commands all forward the flag.

`getProtectionInfo` reports `{isPasswordProtected, isSigned}` from the parsed project: the dir stream's protection state, plus `detectSignature`, which looks for the `_VBA_PROJECT_CUR` / `_VBA_PROJECT_SIGNATURE` streams in the compound file. `XlsmExplorer` lazily probes this when a workbook is expanded and renders `[locked]`/`[signed]` badges on the workbook node. `getWorkbookInfo` folds the same two flags into its summary.

`validateWorkbook` runs structural checks over the parsed project: every dir-declared module must resolve to a readable stream, names must be unique and non-empty, and the PROJECT stream must agree with the dir stream. `createWorkbook` copies `assets/templates/blank.xlsm` (ThisWorkbook, Sheet1, an empty Module1) to the target path.

---

## Module export / import

`moduleExport.ts` is the single source of truth for writing exported module
files. `workbookSettings.ts` owns the workbook settings sidecar path, strict
schema validation, generic workbook-over-global provenance resolution, and
normalized read/patch/write persistence. `workbookModuleSyncSettings.ts` owns
the effective import/export folder and mode model, including source provenance
for workbook overrides, built-in defaults, and unsaved session edits.
`moduleSyncPlan.ts` builds the UI preview model used by bulk import/export and
carries that resolved settings metadata into the webview.

Both lanes call into these shared owners:

- UI commands (`xlide.exportModulesToFolder`, `xlide.importModulesFromFolder`,
  and tree/menu routes that open the same preview GUI)
- AI tools (`xlide_exportModules`, `xlide_configureExportMode`)

**Export** reads all VBA modules live from Excel macro workbooks (`.xlsm`, `.xlsb`, `.xlam`) over JSON-RPC (`listModules` then `readModule` per module) and writes module files to a folder. The UI bulk export command opens a webview diff preview where the user can change the export folder, switch export mode, autosave settings after a short debounce, click each module, compare workbook-vs-repo text, check/uncheck modules, and apply only the selected changes. In `trueUp` mode, stale `.bas`/`.cls` repo module files appear as removable diff rows instead of being removed invisibly.

- Output file extension is `.bas` for standard modules and `.cls` for class/document modules.
- Export mode is per-workbook and persisted in the workbook-local JSON config:
  - `exportAll` (default): export every workbook module; create missing files and update changed files; do not delete stale files
  - `trueUp`: `exportAll` plus remove stale `.bas`/`.cls` module files that no longer exist in the workbook

**Import** (`xlide.importModulesFromFolder`) reads `.bas`/`.cls` files from the configured (or user-chosen) folder and opens the same webview diff preview. Existing modules can be updated through `writeModule`. Standard and class modules can be created from imported files. Document modules and UserForm `.cls` code-behind modules can be updated when the workbook already has a same-named module, but they cannot be created directly from import; missing document/UserForm-code-behind rows show `Skipping import`, remain visible in the preview, and are skipped on apply with an audit entry rather than failing the whole import. `.frm` designer files are ignored by import/export sync. Import mode defaults to `updateOnly` (`Import/Update (No Deletes)`). `trueUpStandardClass` (`Import/Update + Delete Missing`) performs the same create/update pass, then previews workbook-only standard/class modules as removable rows; document modules and UserForm code-behind modules are excluded from true-up removals.

Import/export settings live in the preview GUI so folder and mode edits use the
same resolver, planner, and persistence path as apply. The preview shows quiet
source labels for folder and mode values. Workbook-specific settings are
written beside the workbook; global/default settings live in machine-scoped VS
Code configuration:

```
<workbook-filename>.xlide_settings.json
```

Workbook settings schema:

```json
{
  "exportFolder": "C:/absolute/path/to/export/folder",
  "exportMode": "exportAll",
  "importMode": "updateOnly",
  "analysis": {
    "visibleSeverities": ["error", "warning", "information"],
    "untrackedRules": []
  }
}
```

`trueUp` export treats the selected folder as the workbook's module folder: it proposes/removes only root `.bas` and `.cls` module files that do not map to a live workbook module. Other file types, nested files, and `.frm` designer files are outside import/export sync.

On later runs, `exportFolder` is used as the default folder in the preview GUI.
Mode edits write only the edited sync mode and preserve unrelated workbook
settings, so export saves do not stamp import defaults and import saves do not
stamp export defaults. XLIDE reads and writes only
`<workbook-filename>.xlide_settings.json`; older sidecar names are not part of
the supported settings contract. If that sidecar exists but contains invalid
JSON, unknown keys, invalid sync modes, or invalid analysis settings, XLIDE
reports the settings file as invalid instead of treating it as empty defaults.

---

## AI agent tools

Declared in `package.json` under `contributes.languageModelTools` and registered at activation via `vscode.lm.registerTool`. Copilot can invoke them inline or via `#` references in chat.

For agentic editing, the workbook/XLIDE virtual module structure is the source
of truth. Agents should discover with `xlide_getWorkbookInfo`/`xlide_listModules`,
read with `xlide_readModule`, and write with `xlide_writeModule`. Exported
`.bas`/`.cls` files are sync artifacts unless the user explicitly asks
to operate on export files.

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
| `xlide_analyzeWorkbook` | `#xlideAnalyzeWorkbook` | none | No |
| `xlide_runVbaTests` | `#xlideRunVbaTests` | runs tests + writes artifacts | Yes |
| `xlide_createWorkbook` | `#xlideCreateWorkbook` | creates .xlsm (fails if the file exists) | Yes |
| `xlide_readCells` | `#xlideReadCells` | none | No |
| `xlide_readFormulas` | `#xlideReadFormulas` | none | No |
| `xlide_writeCells` | `#xlideWriteCells` | saves .xlsm | Yes |
| `xlide_exportModules` | `#xlideExportModules` | writes export files + updates workbook JSON config | Yes |
| `xlide_configureExportMode` | `#xlideConfigureExportMode` | updates workbook JSON config | Yes |

---

## Live Share integration — `liveShare.ts`

`LiveShareIntegration` contains infrastructure for XLIDE across VS Code Live Share sessions.

**What works:** The host opens modules through XLIDE normally (creating `xlide-vba://` documents). Live Share mirrors those open documents to the guest. The guest can co-edit and save (Ctrl+S) — the save travels through Live Share's standard editor sync back to the host's `XlideFileSystemProvider.writeFile`, which writes the workbook.

**What does not work:** Independent guest browsing. The code uses the Live Share shared service API (`shareService` / `getSharedService`) under the service name `WilliamSmithE.xlide` to let guests list and open their own workbooks. Microsoft does not allow non-approved extensions to expose guest-accessible shared services — `shareService()` always returns `null`. As a result, the XLIDE sidebar shows nothing on the guest side and guests cannot independently discover or open modules; they can only collaborate on documents the host has already opened.

The host-side RPC handlers (`listWorkbooks`, `listModules`, `listSubs`, `readModule`, `writeModule`) and the guest-side `guestList*` / `guestReadModule` / `guestWriteModule` methods are implemented and would address this gap if Microsoft approval were obtained. Remote modules use `xlide-vba://liveshare/<workbookId>/<moduleName>.bas` URIs so `XlideFileSystemProvider` can route them through the proxy rather than the local workbook engine. `LiveShareIntegration.onDidChange` fires on session role changes so that `XlsmExplorer` and `XlideStatusBar` can refresh.

---

## Status bar — `statusBar.ts`

`XlideStatusBar` manages two `vscode.StatusBarItem` instances:

| Item | Shown when | Text | Click action |
|---|---|---|---|
| Active module | Active editor is an `xlide-vba://` document | `<workbook> | <module>` (or `XLIDE (Live Share)` for remote) | `xlide.refreshExplorer` |
| Live Share | Connected as a Live Share guest | `XLIDE (Live Share): <N workbooks>` | `xlide.refreshExplorer` |

---

## Activity Bar sidebar — `xlideSidebar.ts`

XLIDE contributes a dedicated Activity Bar container (`xlide`) and a polished
WebviewView (`xlide.sidebar`). The workbook/module navigation tree remains in
the VS Code Explorer as `xlide.explorer`; the Activity Bar sidebar is the
product shell for setup status, common actions, compact settings entry points,
and support actions.

`src/xlideSidebarModel.ts` owns the pure sidebar model so the UI shape can be
tested without VS Code APIs. `src/xlideSidebar.ts` renders that model into a
VS Code-themed webview with section spacing, dividers, setup-only status dots,
a compact welcome note that points users to the Explorer-hosted XLIDE tree, and
a dedicated Workbook Actions group containing the target workbook picker and
workbook-scoped action buttons. It does not render module-scoped actions or
module-scoped information; those stay in editor or tree contexts where the
module target is explicit. It refreshes on `xlide.*` configuration changes,
workspace-folder changes, active-editor changes, workbook file create/delete
events, and `.xlide_settings.json` sidecar changes, and does not require Excel
COM to render.

There is no Setup section: the workbook engine runs in-process, so nothing has
to be installed, detected, or repaired before the tree and the actions work. The
Workbook Actions section lets the user pick the
target workbook and keeps workbook-scoped actions directly under that picker. If
no explicit sidebar target is selected, the context can fall back to the active
`xlide-vba://` editor or a single-workbook workspace; if no target exists,
workbook-scoped sidebar buttons are disabled rather than falling back to an
unrelated editor. The selected target is workspace UI state, not a global
setting or workbook sidecar value. The Settings section offers a compact global
settings launcher that opens `xlide.openGlobalSettings`, a dedicated webview for
VS Code machine/profile settings, without rendering workbook-scoped
import/export, analysis, or sidecar rows. The Support section keeps
troubleshooting actions together. Workbook-facing GUIs read
`<workbook-filename>.xlide_settings.json`
through `workbookSettings.ts` and the effective settings helpers that production
commands use.
Workbook-specific settings are not stored globally.

---

## Configuration scopes and precedence

XLIDE supports two configuration scopes:

1. Global/editor settings live in VS Code machine/profile settings and are
   declared as machine-scoped `xlide.*` contributions in `package.json`.
2. Workbook-specific settings live beside the workbook in
   `<workbook-filename>.xlide_settings.json`.

Precedence is deterministic: built-in defaults are the floor, VS Code global
settings override those defaults for the current machine/profile, and
workbook-specific sidecar values override global defaults only for that
workbook. There are no legacy sidecar names and no secondary JSON settings
surface outside workbook sidecars.

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
Because VS Code language configuration is static JSON, tests keep its block
indent/folding regexes aligned with the shared smart-block rules in
`src/vbaSmartEnter.ts` instead of letting it become a second behavioral source.
Block keyword snippets are also projected from the same `VBA_SMART_BLOCK_SNIPPETS`
catalogue (`src/vbaSmartBlockSnippets.ts`), so adding a block archetype requires
updating one pure contract and then satisfying the completion, Smart Enter, and
static JSON conformance tests.
The dependency-free `src/vbaSourceScan.ts` owns shared VBA source-text helpers such
as identifier validation, comment/string-safe identifier occurrence search, line
start offsets, and leading-whitespace detection. Providers, workbook analysis,
code actions, and tree/module commands should reuse those helpers rather than
carrying local regex copies.

**Symbol intelligence** — `src/vbaSymbolIndex.ts` keeps a workbook-scoped cache
of module sources and metadata. The cache loads modules lazily through the
workbook engine (`readModules`, falling back to `listModules` + `readModule`) and
can refresh a single module after a save. Symbol extraction itself lives in the
analyzer (`src/analyzer/symbols/projectIndex.ts`), which consumers feed with the
cached sources.

`src/vbaLanguageProviders.ts` is the composition root that registers the
language providers plus diagnostics and smart-enter editing against the `vba`
language under the `xlide-vba` scheme; the implementations live in the
per-subsystem modules listed in the repository layout (`vbaLiveDiagnostics.ts`,
`vbaCompletionProvider.ts`, `vbaHoverSignatureProvider.ts`,
`vbaNavigationProviders.ts`, `vbaSemanticTokensProvider.ts`,
`vbaCodeActions.ts`, `vbaTypingAutomation.ts`):

| Provider | Behavior |
|---|---|
| `DocumentSymbolProvider` | Outlines the current module from the analyzer `ProjectIndex` via `documentOutlineSymbolsForSource` |
| `DefinitionProvider` | Builds an AST `ProjectIndex` and resolves source-backed `object.Member` references through the shared member-completion binder, resolves project type-name tokens through `resolveTypeDefinitions`, then falls back to scope-aware name resolution (`resolveDefinition`); honors a `Module.Member` qualifier via `resolveQualifiedDefinition`, and follows MS-VBAL visibility (locals shadow module members shadow exported cross-module declarations, including enum members exported by their containing `Enum`) |
| `ReferenceProvider` | Project type-name tokens are matched through `resolveTypeDefinitions`; everything else flows through the shared pure `collectSymbolReferences` (`src/vbaReferenceResolution.ts`), which returns the **union** of (a) member-access / module-qualified occurrences (`Module.Proc`, `obj.Member`, `Me.Member`) validated by their resolved class-member definition spans, and (b) bare-identifier occurrences each re-resolved in their own module context and kept only when they bind *solely* to the target declarations. Honors VS Code's include-declaration toggle |
| `RenameProvider` | Shares the exact same `collectSymbolReferences` resolver as references, so a standard-module procedure/variable, class member, property family, Declare, Enum member, or local/parameter renames its declaration **and every bare call (same- and cross-module) plus its validated qualified/member-access references** — regardless of whether rename is invoked from the declaration, a bare call, or a qualified reference. The per-occurrence bind check keeps a same-named member on an unrelated receiver, a same-named procedure in another module, a shadowing local, and ambiguous bare calls from being renamed. VBA component/module rename is intentionally tree-only because standard and class module names are workbook component names rather than in-source declarations; Enum-type-qualified access (`Color.Red`) is not matched (the analyzer exposes no `EnumType.Member` resolver, matching go-to-definition) |
| `CodeActionProvider` | Delegates XLIDE diagnostics to the pure `resolveDiagnosticCodeActions` resolver and converts returned offset edits into VS Code quick fixes; first supported fixes add `Option Explicit`, move misplaced `Option` statements, split local `Dim` initializers, insert missing block closers, insert missing explicit-`Call` and expression-call argument-list parentheses, remove illegal empty parentheses from standalone zero-argument calls, rewrite invalid `Call DoEvents()`-style runtime statements, add/remove `Set` for proven object/scalar assignments, and expose an XLIDE source action for analyzing the current workbook-backed module. Call-related quick fixes consume shared `callContext` range helpers so diagnostics and repairs do not parse chained callees, empty parentheses, or invalid explicit `Call` syntax differently |
| Diagnostics | Debounced `vbaModuleAnalysis` results merge structural block-balance, semantic analysis, and suppression directives before rendering VS Code diagnostics; live diagnostics pass the active cursor offset through the shared incomplete-expression detector so dangling member access, trailing binary operators, and unmatched opening parentheses are hidden only while their active statement is being edited, then restored on selection/editor change, current-module analysis, or workbook analysis. Structural block diagnostics pin ranges to the opener/closer syntax phrase, and analyzer declaration-order rules prefer the exact offending token over whole-line or whole-parameter spans |
| Smart enter (auto-block) | Pressing Enter after a safe block opener inserts the matching closer through the shared smart-block helper. The default `xlide.editor.blockLayout = "comfy"` layout uses a spacer line, one editable body line one real tab deeper than the opener, another spacer line, then the closer at opener indentation; `"compact"` places the editable body directly above the closer. If a matching closer already exists ahead, Smart Enter only normalizes the newly created body line indentation. Supported openers include procedures, `If ... Then`, `With`, `For`, `Do`, `While`, `Select Case`, `Type`, `Enum`, and `#If`, with `With` seeding a leading `.` for member completion. Pressing Enter after a leading-dot member line inside an active `With` block keeps the same indentation and seeds another leading `.` |
| Loop iterator sync | Editing the iterator token in a simple `For` / `For Each` opener or its matching `Next name` updates the paired token, using the same lexer-derived string/comment stripping as the rest of Smart Enter and conservative block matching |

The expanded Smart Enter layout is the default `comfy` block style. The
`compact` style removes spacer lines and is exposed as the VS Code extension
setting `xlide.editor.blockLayout`, contributed by `package.json` with machine
scope rather than
workbook sidecar configuration. Both modes are expressed in the same
smart-block helper/catalogue so Enter auto-blocking and Tab snippets share one
behavior contract.

Language-service business rules are unified across surfaces. Unless a behavior is
called out as a deliberate corner case, completion insert text, hover, signature
help, diagnostics, navigation, rename, tree actions, semantic coloring, snippets,
formatter logic, and code actions must all consume the same analyzer rules or
provider helpers for the same VBA construct.

The `DefinitionProvider`, `ReferenceProvider`, and `RenameProvider` read the
shared per-workbook incremental `ProjectIndex`
(`src/vbaProjectIndexService.ts`, over
`src/analyzer/symbols/projectIndex.ts`) with the live editor text overlaid for
the current module — definition/references use the `live` view, while rename
uses the `strict` view so reference edits never silently skip modules that
fail to index. Offset-based symbol spans are converted to editor ranges in the shared
`vbaNavigation.ts` helpers and provider wiring. `VbaSymbolIndex` still backs
the workbook-scoped source cache those queries read from.
Tree-level module rename uses source-backed project helpers before changing the
component name through `renameModule`: class modules rewrite workbook
project-defined class type tokens, while standard modules rewrite bound
module-qualified qualifier tokens for exported members and visible qualified
type names.

Project-aware diagnostics, analysis surfaces, semantic type coloring, completion,
and hover derive project context through `src/vbaProjectAnalysis.ts`. That
helper owns module-kind normalization, live-current-module overlay, project
procedure signatures, visibility-filtered procedure/identifier/type/non-type
names, source-backed member surfaces, and the named read-only live-analysis path
that ignores temporarily invalid modules. Completion, definition, references,
diagnostics, semantic coloring, hover/signature contexts, and current-module
analysis use that path. It also exposes `projectEditorSymbolContextForModule()`,
the provider-facing bridge for external project procedures/symbols plus the
analysis options used by editor surfaces. `src/vbaMemberCompletion.ts` derives a
single workbook-aware project context per completion/hover/signature/casing
request from the shared `VbaProjectIndexService` (via
`src/vbaEditorProjectContext.ts`), then projects it into member, type,
identifier, hover, signature-help, event-handler, and canonical-casing
resolver contexts. Rename and tree-level module rename are the deliberate
mutation-safety exception: they use strict project indexes so reference edits do
not silently skip modules that cannot be parsed.
Before those project indexes are built for live editor providers, workbook
modules are passed through `src/vbaOpenDocuments.ts` so unsaved open modules from
the same workbook participate in cross-module diagnostics, type coloring,
navigation, completion, hover, and signature help without crossing workbook
boundaries.

**Structural analysis** — `src/vbaStructuralDiagnostics.ts` is a pure, `vscode`-free module so it
is unit-tested directly (`tests/vbaStructuralAnalysis.test.ts`). It strips strings/comments,
joins `_` line continuations (via `src/vbaSourceScan.ts`), then walks a block stack to detect
imbalance. Missing-block diagnostics carry stable codes plus the expected closer and
deterministic insertion line, so quick fixes can insert `End Sub`, `End If`,
`Next`, and related closers without parsing diagnostic text. The sibling
`src/vbaSmartEnter.ts` exports the smart-enter helpers used by the auto-block
feature; its stripped-line substrate is derived from the analyzer lexer
(`src/analyzer/lexer/strippedLines.ts`), not from `stripVba`.
The planned convergence onto the token-based analyzer (audit #74) is tracked by
two corpus-wide comparison harnesses: `tests/structuralEngineComparison.test.ts`
diffs this engine against the parser's block-balance recovery diagnostics (ten
verified divergence classes, with regressions in both directions, so the legacy
engine stays), and `tests/smartEnterSubstrateComparison.test.ts` pins Smart
Enter's now-production lexer substrate against `stripVba` (migrated; the single
deliberate divergence is the lexer correctly blanking trailing `: Rem ...`
comments per MS-VBAL 3.3.5.2, which `stripVba` leaks).
Both harnesses fail on any new, unexplained drift between the engines.

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

**Workbook-wide analysis (command + agent tool)** — `src/vbaWorkbookAnalysis.ts`
(`analyzeWorkbook`) loads every module from the workbook via the workbook engine,
builds shared project-analysis options so cross-module rules have the current
module's visibility-filtered procedure/Declare names, bare identifier names,
visible type/non-type names, exported signatures, and source-backed member
surfaces, then delegates each module to `src/vbaModuleAnalysis.ts`. That shared
module core merges
`analyzeVbaStructure`, `analyzeModule`, and analysis suppression directives before
workbook analysis flattens results into 1-based `{moduleName, moduleType, line,
column, endColumn, severity, code, message, category, vbeCompileEquivalent,
diagnosticKind}` problems, sorted by module/line/column. Semantic analyzer rules
and structural block-balance codes resolve through the shared diagnostic
metadata catalogue before the workbook summary is counted, so the command GUI
and `xlide_analyzeWorkbook` agent JSON do not maintain separate rule buckets. The
`xlide.analyzeWorkbook` command (`src/commands/analysisCommands.ts`, right-click "Analyze Workbook"
on a workbook tree node) and `xlide.analyzeCurrentModule` both open the
dedicated analysis results panel from `src/workbookAnalysisWebview.ts`. The panel is
driven by the pure model in `src/workbookAnalysisResultsModel.ts`, shows severity
and VBE-equivalence filters, module grouping, counts, suppressed-diagnostic
visibility, copy-report/copy-JSON actions, and click-through navigation that
opens the exact module span in an adjacent editor group through
`encodeModuleUri`. The Output channel remains a
support log, not the primary analysis-results surface. The same core is exposed to
AI agents as the `xlide_analyzeWorkbook` LM tool (`src/agentTools.ts`), returning
the structured JSON report so an agent can verify analysis passes in real time after
editing modules.

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
interim regex analyzer in `src/vbaStructuralDiagnostics.ts`; Smart Enter and
keyword completion already sit on the lexer-derived substrate (audit #74,
`src/analyzer/lexer/strippedLines.ts`, gated by
`tests/smartEnterSubstrateComparison.test.ts`), while the
`tests/structuralEngineComparison.test.ts` harness documents the exact
behavioral gaps that must close before block-balance diagnostics can move
without changing user-visible output.

**Host-context member completion** — built on top of the analyzer, this is the
feature that distinguishes XLIDE from a generic VBA syntax extension. It is split
into a pure analyzer layer and a thin VS Code provider:

- `src/analyzer/host/excelObjectModel.ts` is a curated, verified subset of the
  Excel automation object model (Application/Workbook/Worksheet/Range plus the
  commonly used Window, Name(s), Comment(s), ListObject/Row/Column(s),
  PivotTable(s), Chart(s)/ChartObject(s), Shape(s), Font, Interior, Border(s),
  Areas, Hyperlink(s), WorksheetFunction, Style(s), PageSetup, Validation,
  Slicer/timeline, shape-formatting, chart-layout, conditional-format subtype,
  legacy drawing, sparkline, XML, publish, and web-option types), with each
  type's properties and methods transcribed from the official
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
  generator currently emits 229 promoted Excel runtime surfaces into
  `src/analyzer/host/excelReferenceMembers.ts`, including available member
  signatures and reference documentation. Of those, 36 proven surfaces are
  marked exhaustive for hard unknown-member diagnostics, while
  `WorksheetFunction`, Pivot families, QueryTable/data-connection families,
  chart internals, ShapeRange, comments, sort/filter helpers, form-control
  families, Slicer/timeline objects, shape-formatting internals, chart-layout
  internals, conditional-format subtypes, legacy drawing objects, sparkline
  objects, and XML/publish/web workbook surfaces are promoted for completion,
  signature help, hover, and
  receiver-chain inference only until oracle evidence proves a family can
  safely support hard absence diagnostics. Reference events are counted in
  coverage but filtered out of object-member surfaces because VBE does not
  expose events as callable object methods/properties; document-module
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
  chain, unwraps parenthesized receiver expressions whose inner receiver chain
  is already deterministic, and returns the filtered members. For leading-dot
  member access inside
  `With`, the resolver scans the enclosing procedure up to the dot, builds the
  active `With` receiver stack, and resolves nested `With .Member` expressions
  outer-to-inner before applying the same member-chain logic. When a typed
  variable proves a more specific member surface after assignment from a mixed
  collection (for example
  `Dim ws As Worksheet: Set ws = Sheets(1)`), the declared type wins over the
  collection's merged item surface. For generic `Object`/`Variant` variables,
  the resolver can refine completion from the latest preceding simple `Set`
  assignment to a known object expression, so `Set obj = Worksheets(1)` narrows
  `obj.` to worksheet members while `Set obj = Sheets(1)` keeps the merged
  worksheet/chart surface. Diagnostic callers can opt out of this refinement so
  late-bound `Object`/`Variant` receivers are not treated as hard absence proof;
  editor completion leaves it enabled. It also accepts source-backed workbook
  member surfaces from `ProjectIndex.projectMemberSurfaces(moduleName)`, so
  variables declared as workbook classes (for example `Dim p As Person`) can
  offer public/default-public source members and public fields at `p.`, and
  variables declared as visible user-defined types can offer their fields at
  `point.` without guessing from names. Public constants are not exposed as
  object members because VBE rejects them in class/document/UserForm modules. The same context also
  resolves `Me.` to the current class/document module's source-backed member
  surface, merging with a known host surface when the caller supplies one.
  Source-backed workbook members carry inline `'''` documentation through to
  completion, hover, and member-call signature help, and carry declaration spans
  for source-backed member go-to-definition. Exported member attributes such as
  `Attribute Value.VB_UserMemId = 0` are attached to the same source-backed
  member surface and mark `defaultMember`, but direct object-expression
  inference is not enabled until that behavior has separate oracle coverage.
- `src/vbaMemberCompletion.ts` is the VS Code `CompletionItemProvider` (trigger
  characters `.` and space). For member access it loads the workbook's module
  list via the workbook engine, then uses `src/vbaProjectAnalysis.ts` for
  module-kind normalization, live-current-module overlay, project type/member
  derivation, and invalid-module tolerance while the user is mid-edit. The
  resulting context includes document code names with workbook, worksheet, or
  chart host type where known, plus the host/source `Me` context for the current
  object module. For workbook class members, open XLIDE module documents are
  read from their live editor text first, so unsaved changes in an open
  `Person` class are reflected the next time completion is requested elsewhere;
  saved module text is read through the bridge when no live editor text exists.
  At module level in
  document modules, it also offers event-procedure stubs from
  `resolveEventHandlerCompletions`: `ThisWorkbook`/`documentType: workbook`
  gets `Workbook_*` handlers, worksheet document modules get `Worksheet_*`
  handlers, and chart document modules get `Chart_*` handlers. UserForm
  handler authoring stays out: designer-backed form/control event metadata is a
  permanent won't-implement (see the Won't Implement section in
  `docs/spec/MS-VBAL.verification-map.md`). Existing handlers are not re-suggested, and
  accepting an event completion inserts either the full `Private Sub ... End
  Sub` stub or only the declaration tail after an existing `Private Sub`
  prefix. Known workbook/worksheet/chart event-handler-shaped `Sub`
  declarations in the wrong module receive non-red
  `event-handler-module-scope` guidance because
  Excel will not wire them as events there, even though they may compile as
  ordinary procedures. In a declaration type position (after `As`) or a `New`
  type position (after `As New` or expression-level `New`) it instead offers
  type-name completions via `src/analyzer/completion/typeCompletion.ts`
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
  Block keyword completions are projected from `src/vbaSmartBlockSnippets.ts`'s shared
  `VBA_SMART_BLOCK_SNIPPETS` catalogue and remain explicit full-block scaffolds for Tab-driven
  shortcut gestures, while Smart Enter handles the line-by-line workflow after a
  user-typed opener. Close-keyword suggestions still consume the same
  smart-block stack as Smart Enter, so active closer labels such as `Next cell`
  do not fork from the block model. Loop snippets use a transformed iterator
  mirror on the `Next` line rather than a second linked placeholder, so leaving
  the iterator field does not keep a cross-line placeholder selection alive.
  After insertion, simple loop iterator edits are synchronized bidirectionally
  between the opener and the matching `Next name`.
  Keyword snippets and Smart Enter use the same literal-tab block indentation
  unit, materialized with the current line's base indentation and
  `keepWhitespace`, so all block archetypes preserve real tabs in their bodies.
  `As New` and expression-level `New` completion are narrower and offer only
  creatable project classes/UserForms until host/external creatability metadata
  exists.
  Accepting callable completions applies canonical casing and uses the shared
  VBA call-site rule: standalone call statements insert only the canonical name,
  while expression contexts and explicit `Call` statements may insert `()` with
  the cursor inside the call. Runtime entries that opt out of explicit `Call`
  through verified metadata are filtered at the `Call <target>` position, so
  invalid forms such as `Call DoEvents` are not offered there. Typing a
  same-line boundary after a known identifier applies a single-token casing edit,
  while pressing Enter, moving the cursor to another line, switching editors, or
  leaving the editor applies all safe VBE-style canonical casing edits on the
  line just left. Both paths use the same analyzer resolver for keywords, type
  names, runtime functions, project identifiers, and resolved host/source
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
identifier completion, and high-confidence diagnostics. Diagnostics treat these
runtime signatures as a last-resort metadata source: source/project callables win
first, and non-callable source names in the active procedure/module suppress bare
runtime fallback even when the declaration has no explicit `As` type.

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
members and source-backed project member surfaces share receiver binding,
return-type chains, and inline XML docs. Bare-call signatures are sourced from
same-module user `Sub`/`Function`/`Property` procedures (built from the parsed
AST so `Optional`/`ParamArray`/default detail renders in VBE bracket form), then
`resolveRuntimeFunction`; runtime entries that are not valid explicit `Call`
targets suppress their call tip in that context using the same runtime metadata
as completion and diagnostics. The whole resolver is wrapped in try/catch so it never
disrupts editing, and signatures are never invented — an unknown callee yields
no tip. Verified host signatures live beside the object model in
`excelObjectModel.ts`; source-backed callable member signatures are emitted by
`ProjectIndex.projectMemberSurfaces(moduleName)`.

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
`user_guides/vba-doc-comments.md`.

**Active diagnostics engine** — `src/analyzer/diagnostics/` computes
high-confidence semantic problems directly from module text:

- `src/analyzer/diagnostics/ruleMetadata.ts` is the typed rule catalogue
  (`DIAGNOSTIC_RULES`): each rule carries a stable `code`, `title`,
  `defaultSeverity`, `category`, `vbeCompileEquivalent`, `diagnosticKind`,
  `source: 'XLIDE'`, an MS-VBAL `specReference`, and a `confidence`. Only
  high-confidence rules ship. The same module also owns metadata for structural
  block-balance diagnostics and the Problems source-label helper used by live
  diagnostics: `XLIDE/VBE` for compile-equivalent findings, `XLIDE/runtime` for
  deterministic runtime failures, `XLIDE/risk` for runtime-risk warnings, and
  `XLIDE/style` for guidance/style rules. The diagnostic `code` remains the
  stable rule id for quick fixes and future filtering.
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
| Uncertain, incomplete while typing, host-dependent, or heuristic-only behavior | No diagnostic | No active rule until the behavior is spec-backed or oracle-verified; if the final syntax is oracle-verified invalid but useful while typing, live diagnostics may suppress the diagnostic only for the active edit line |

- `src/analyzer/diagnostics/analyzeModule.ts` exposes
  `analyzeModule(source, opts)` returning `VbaDiagnostic[]` (code, message,
  severity, offset `span`). It lexes and parses once per pass, reuses the
  symbol graph, and runs the rule families registered in
  `src/analyzer/diagnostics/registry.ts` (one module per family under
  `src/analyzer/diagnostics/rules/`, sharing one statement walk and the
  memoized caches in `analysisContext.ts`/`typeInference.ts`). The rules
  include: unterminated string (odd-quote-count, escape-aware),
  duplicate procedure (Property Get/Let/Set may share a name), duplicate
  declaration in a flat procedure scope, duplicate module-level variable,
  assignment to a `Const` (bare `name =` only - excludes `.member`, indexing,
  `Set`, and comparisons), a configurable `Option Explicit`-missing
  reminder (silent on empty/attribute-only modules), an `unknown-call`
  ("Sub or Function not defined") rule, an `invalid-proc-header`
  ("Invalid procedure declaration") rule, an `unbalanced-parens` rule, and an
  `argument-count` ("Wrong number of arguments") rule. `src/analyzer/call/callContext.ts`
  owns shared VBA call-statement classification for signature help, completion
  parenthesis insertion, and diagnostics. `bareCallStatementTarget` powers
  `unknown-call`: it accepts the three call forms whose callee is a bare
  (non-member) identifier - a lone identifier, a parenless call with arguments
  (`MsgBox "hi"`), and an explicit `Call name` - and flags the callee when the
  name resolves to no project procedure/Declare, runtime function/statement, host
  global, `Application` member, or in-scope declaration. It excludes member
  calls (`.`), labels (`:`), assignments (any top-level `=`), and the
  implicit-host-member form `Cells(1, 1)` / `Range("A1")` (a non-`Call`
  identifier immediately followed by `(`). `checkCallParens` powers call syntax
  diagnostics using the same shared classifier: explicit `Call` statements with
  arguments require parentheses, and VBE-oracle-verified standalone
  zero-argument calls cannot use empty parentheses unless they are prefixed with
  valid `Call` syntax or used in an expression.
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
  names against the parameters, and skips unresolved or ambiguous callees. Bare
  runtime signatures are consulted only after visible source names have failed
  to bind, so local/module declarations such as `Dim Format` or `Const Format`
  suppress runtime-shaped arity, type, return-inference, expression-call, and
  runtime-only `Call` diagnostics. The
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
  parameter/local/module/project value types resolved through the shared
  expression resolver, module-qualified `ModuleName.ValueName` reads resolved
  through the matching source-backed module surface, `Function`/`Property Get`
  return-name types, zero-argument Function/Property Get value references,
  curated runtime return metadata, same-module return types, known
  source-backed or host/reference member signatures, proven source-backed or
  host/reference member return expressions, and deterministic
  literal/expression inference; unknown, ambiguous, untyped, and `Variant`
  operands suppress diagnostics. Bare member functions with required
  arguments do not infer as value references unless the expression is an actual
  call. For
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
  diagnostic for source-backed project properties whose member surface has
  no setter; focused oracle evidence rejects this as `Can't assign to read-only
  property`. `member-not-found` is another source-backed member rule: it
  fires only when a receiver resolves to an unambiguous and exhaustive
  `ProjectIndex.projectMemberSurfaces(moduleName)` surface, or a promoted exhaustive host
  surface, and the member name is absent. Plain class modules are
  source-exhaustive, visible UDT field surfaces are exhaustive, standard-module
  qualifier surfaces are source-exhaustive even when no visible member rows are
  available, and current-class `Me.Member` references use the same contract;
  leading-dot members inside simple or nested `With` blocks share that receiver
  contract because diagnostics call the same member resolver as completion.
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
  shape is backed by focused VBE oracle evidence. Recognized fixed-length string
  suffixes such as `As String * 10` are parser-owned metadata
  (`asType: String`, `fixedLength: 10`) used by symbols, hover, UDT member hover,
  and object-module public-member diagnostics. The declaration trailing-token
  rule consumes that suffix before looking for extra tokens, and
  `fixed-length-string-size` uses the same parsed suffix to flag literal sizes
  and deterministic same-module/procedure `Const` integer expressions outside
  the VBE-verified `1..65526` range. Unknown/non-deterministic length
  expressions and assignment/truncation behavior remain on the verification
  roadmap. A separate
  `object-module-public-member` rule is module-kind-sensitive: in class,
  document, and UserForm modules it rejects explicit Public constants, arrays,
  fixed-length strings, user-defined Types, and Declare statements. Each branch
  is backed by focused VBE oracle evidence, while standard modules and non-Public
  declarations stay outside the rule.
  `event-declaration-module-kind` is a red module-kind diagnostic for `Event`
  declarations in standard modules; class, document, and UserForm modules remain
  accepted, and inactive conditional-compilation branches stay filtered.
  `withevents-declaration` covers the parser-proven `WithEvents` declaration
  restrictions: standard modules, procedure-local declarations, `As New`, and
  array declarators are red, while object-module module-level declarations stay
  quiet and event-source type compatibility remains deferred.
  `friend-declaration` covers parser-proven `Friend` declaration restrictions:
  standard-module procedures and module variable declarations are red, while
  object-module procedures and inactive conditional branches stay quiet.
  `implements-statement-placement` is a red module-kind/placement diagnostic for
  `Implements` statements in standard modules, procedure bodies, or after a
  procedure in an object module. Declaration-section object-module statements
  remain accepted, inactive conditional branches are filtered, and interface
  member completeness is left to the project binder.
  `module-declaration-after-procedure` covers the shared declaration-section
  ordering rule for parsed module declarations: active `Declare`, `Event`,
  module variable/`Const`, `Type`, and `Enum` declarations after an active
  procedure are red and pinned to the declaration keyword, while inactive
  conditional branches and later procedures stay quiet.
  `raiseevent-undeclared-event` is a red same-module binding diagnostic for
  settled `RaiseEvent` statements whose target does not resolve to an active
  `Event` declaration in the containing module. It scans tokenized procedure
  body lines so line-numbered statements are covered, while partial
  `RaiseEvent` statements and event signature/arity checks remain deferred.
  `event-handler-module-scope` is an information diagnostic for known
  workbook/worksheet/chart event handler names declared outside the matching
  document module. It uses the same module-scoped event metadata as completion,
  and it is deliberately non-red because Excel treats those declarations as
  ordinary procedures rather than wired event handlers.
  `invalid-declaration-name` flags unbracketed MS-VBAL reserved identifiers in
  declaration-name positions while accepting bracketed `FOREIGN-NAME` forms such
  as `[In]`.
  `invalid-as-type-name` and `invalid-new-type-name` use the shared
  type-position scanner and resolver for `As`, `As New`, return, parameter, UDT
  field, expression-level `New`, `TypeOf ... Is`, and `Implements` positions.
  Resolved project/primitive/host types stay quiet in ordinary type positions;
  unresolved reserved identifiers, runtime functions, visible project
  declarations that are known not to be types, and ambiguous visible project
  type names are compile-equivalent errors. Resolved non-creatable types are
  compile-equivalent errors after `New`/`As New`, using the same creatability
  predicate as completion. Broad unknown type names still wait for the
  external-reference story so referenced-library classes are not flagged
  prematurely.
  `set-required` fires when a plain assignment targets a known object variable,
  `Function`/`Property Get` return name, visible exported standard-module
  object global, or source-backed object-valued member (`Property Set` or
  public field) that requires `Set`; `set-requires-object` fires when `Set`
  targets a known intrinsic scalar variable, visible exported standard-module
  scalar global, or source-backed scalar member. Both rules route bare
  assignment targets through the shared assignment-target resolver, so local
  shadows win and `Variant`, unknown types, and ambiguous project targets stay
  silent.
  `array-assignment-to-scalar` and `array-bound-requires-array` use the shared
  resolver for simple bare target/source/argument shapes before falling back to
  legacy local shape maps. Visible exported standard-module arrays/scalars feed
  these diagnostics, while local shadows, ambiguous exported globals, indexed
  values, member-expression arguments, `Variant`, and unknown shapes stay quiet.
  `fixed-array-redim` and `scalar-redim` use the same assignment-target resolver
  for bare `ReDim` targets, with fixed-array bounds carried on variable symbols
  from parser metadata. Visible exported standard-module scalars and fixed
  arrays feed these diagnostics, while dynamic arrays, scalar `Variant` and
  implicit Variant targets, local dynamic shadows, ambiguous exported globals,
  undeclared targets, and inactive branches stay quiet.
  `erase-requires-array` uses the same assignment-target resolver for simple
  bare `Erase` targets. Visible exported standard-module non-Variant
  scalar/object globals feed the diagnostic, while arrays, `Variant` and
  implicit Variant targets, local shadows, ambiguous exported globals,
  unresolved names, non-simple member/index targets, and inactive branches stay
  quiet.
  `for-each-control-variable-type` resolves simple `For Each` control variables
  through the shared assignment-target resolver, and `for-each-source-type`
  resolves simple `In` source names through the shared expression resolver.
  Visible exported standard-module scalars/arrays feed the relevant diagnostics,
  while `Variant`, object-like values, local shadows, ambiguous exported
  globals, unresolved names, member-chain sources, and inactive branches stay
  quiet.
  `scalar-member-access` fires only when a bare receiver is a declared
  intrinsic scalar (`String`, numeric, `Boolean`, or `Date`), including visible
  exported standard-module scalar globals proven through the shared
  member-receiver resolver; focused oracle cases show named scalar members are
  VBE Compile errors (`Invalid qualifier`), while a trailing scalar dot is a
  VBE Compile `Syntax error` after explicit focus retesting. Local shadows,
  ambiguous exported globals, qualified/member-chain tokens, unknown,
  `Variant`, object-like, project class, and UDT receivers stay silent until
  the binder can prove more. `invalid-expression-syntax` asks that same
  member-receiver resolver before reporting a generic incomplete trailing-dot
  member access, so known scalar receivers are left to `scalar-member-access`
  while local shadows, ambiguous exported globals, object-like receivers, and
  unknown receivers still use the syntax fallback.
  `undeclared-variable` runs only when `Option Explicit` is present and the
  caller passes project context (`knownIdentifiers` from
  `ProjectIndex.visibleIdentifierNames(moduleName)` plus, when available,
  `projectVisibleSymbols` from `ProjectIndex.visibleIdentifierSymbols(moduleName)`).
  The rule asks the shared source resolver for local/module/project bindings at
  each procedure site, so unknown assignment targets, read references, member
  receivers, indexed bases, and block-header expressions become
  compile-equivalent `Variable not defined` diagnostics while public
  standard-module globals, enum members, document/UserForm code names, runtime
  functions/constants, host globals/enum constants, and `Application` members
  suppress false positives. Type-name positions, labels, named-argument labels,
  and arguments to unresolved external-style calls remain silent until the
  binder can prove those broader reference shapes. `unknown-call` runs only when
  the caller passes `knownProcedures` (the current module's visibility-filtered
  procedure names from `ProjectIndex.visibleProcedureNames(moduleName)`); it and
  `non-callable-call` use the same source resolver so local/module bindings
  take precedence over exported project procedures and duplicate project
  targets remain quiet. Without `knownProcedures`, `unknown-call` is skipped so
  a single module is never analysed in isolation. The whole analyzer is wrapped in
  try/catch so a parse hiccup returns `[]` and never breaks editing, and accepts
  `severities` overrides (including `'off'`) per rule.

`src/vbaModuleAnalysis.ts` merges this engine with the structural block-balance
analyzer (`src/vbaStructuralDiagnostics.ts`, which owns the "Missing End .../unexpected
terminator" family) and analysis suppression directives. `registerVbaDiagnostics` in
`src/vbaLiveDiagnostics.ts` runs that shared module core on open and debounced
(300 ms) on every edit, on real `.vba` files and on virtual `xlide-vba` module
documents, with no save. Everything is computed from the live editor text plus
deterministic project context. For workbook-backed documents the provider
overlays the current live module text into a fresh shared project-analysis
context, then passes its visibility-filtered procedures, visible bare
identifiers, visible project type names, non-type-name exclusions,
cross-module standard-module signatures, and source-backed project member
surfaces into `analyzeVbaModuleSource`.
Live editor diagnostics also pass the active cursor offset into
`incompleteExpressionEditSpan`, which scopes to the active colon-separated
statement and suppresses overlapping hard syntax diagnostics for incomplete
member access, trailing binary operators, and unmatched opening parentheses.
Current-module analysis and workbook analysis do not pass that offset, so completed
invalid source remains diagnostic-strict outside the active edit.
Setting `xlide.diagnostics.enabled` gates it and re-runs open documents on
change. Analysis rule severity overrides are keyed by stable diagnostic code and
resolved through `xlide.analysis.ruleSeverityOverrides` globally or the
workbook sidecar per workbook. Contributed `xlide.*` VS Code settings are
machine-scoped in `package.json`. `globalSettings.ts` validates, normalizes,
and resolves those contributed settings with explicit provenance (`default`,
`machine`, or `unknown`). Malformed values surface as `XLIDE/settings`
diagnostics on VBA documents; invalid rule severity overrides are rejected by
the same guardrail model used by live diagnostics, current-module analysis, and
workbook analysis.
Before diagnostics are displayed, `src/analyzer/diagnostics/analysisSuppressions.ts`
scans tokenized apostrophe comments for explicit `@xlide-analysis` directives and
applies the same lexical file, member, line, next-line, and paired block filter
inside `vbaModuleAnalysis`. Directive diagnostics, such as
malformed code lists, unknown diagnostic codes, late `disable-file`, invalid
`disable-next-member` placement, or stray/unbalanced/mismatched block pairs, are
added back as `XLIDE/style` warnings and are not suppressed by suppression
directives. Workbook analysis includes the suppressed diagnostic count in its
summary so hidden problems remain auditable.

**Project-wide symbol graph** — `src/analyzer/symbols/` projects the parser AST
into a host-agnostic symbol model so XLIDE can offer document symbols, workspace
symbols, and go-to-definition over the whole workbook project rather than a
single module:

- `src/analyzer/symbols/symbolModel.ts` defines the `VbaSymbol` shape (name,
  kind, `nameSpan` for the identifier, `fullSpan` for the declaration,
  visibility, `asType`, array/fixed-array metadata, exported `Attribute`
  metadata, and nested children).
- `src/analyzer/symbols/buildModuleSymbols.ts` walks one `ModuleNode` and emits
  a hierarchical module symbol (procedures with parameter/local children,
  `Type` with fields, `Enum` with members, module variables/consts, and
  `Declare`s with PtrSafe/Lib/Alias/parameter/return metadata).
  Identifier spans are located with the real lexer, so a `nameSpan` never lands
  inside a comment or string. Dotted exported attribute lines are mapped by the
  target before the dot, so `Attribute Value.VB_UserMemId = 0` attaches to the
  `Value` member while retaining the attribute source span.
- `src/analyzer/symbols/nameResolution.ts` owns the shared bare-identifier
  precedence helper. It records the context (`expression`, `call`,
  `assignmentTarget`, `memberReceiver`, `typeName`, or `newExpression`), winning
  source tier (`local`, `module`, or `project`), ambiguity, definitions, and a
  short explanation. `ProjectIndex`, Option Explicit, const-assignment,
  assignment-target typing, `Set` target typing, array/scalar shape
  diagnostics, `ReDim`/`Erase` target-shape diagnostics, `For Each` control and
  source-shape diagnostics, scalar member-receiver diagnostics, runtime
  integer-constant folding, ambiguous enum-member, unknown/non-callable call
  diagnostics, and runtime-shadow diagnostics consume this helper so editor
  navigation and red-squiggly fallback rules do not grow separate precedence
  ladders.
- `src/analyzer/symbols/projectIndex.ts` is the `ProjectIndex` that aggregates
  modules and answers `documentSymbols`, `workspaceSymbols`, conservative
  `resolveDefinition` (via the shared bare-identifier resolver: locals/params ->
  same-module declarations -> exported declarations in other modules, with enum
  members inheriting project visibility from their containing `Enum`),
  `resolveQualifiedDefinition` (the exported
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
  `projectMemberSurfaces` (source-backed member surfaces for object modules
  and visible UDT fields with signatures, docs, definition spans, module-level
  `Implements` names, and default-member facts from `VB_UserMemId = 0`
  attributes), and
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
`src/vbaSemanticTokensProvider.ts` overlays the live editor text for the current
module before resolving visible types, so `Dim customer As Person`, `Dim ws As
Worksheet`, and `Dim amount As Currency` color as soon as their type is known.
Unresolved names and non-type positions are ignored.

---

## Key design decisions

| Decision | Rationale |
|---|---|
| `FileSystemProvider` over `TextDocumentContentProvider` | Read/write virtual FS — Ctrl+S writes back with no custom save command |
| In-process engine over an external backend | No runtime to install, detect, or recover; no per-call IPC or process-startup cost, and no setup gate between install and first use |
| Own the container formats ([MS-CFB], [MS-OVBA], OOXML) | The binary formats are specified and stable; owning them removes the last third-party dependency and lets writes preserve untouched bytes exactly |
| Rebuild-and-rename on every mutating save | A canonical rebuild keeps the compound file self-consistent, and the atomic rename means a crash mid-write cannot corrupt the workbook |
| No COM, no Office | Works on Windows, macOS, Linux, WSL, and remote containers |
| Confirmation on write tools | Prevents AI agents from silently mutating production workbooks |

---

## Dependencies

XLIDE ships with no runtime dependencies. VBA and worksheet access are
implemented in `src/vba/**` against the published container formats; the only
Node built-ins used are `fs`, `path`, and `zlib`.

TypeScript dev: `typescript`, `esbuild`, `vitest`, `@types/vscode`, `@types/node`.

---

## Files to keep up to date

| Change | Files to touch |
|---|---|
| New engine method | `src/vba/workbookService.ts` (implementation), `src/workbookEngine.ts` (dispatch), `tests/vbaNativeWorkbook.test.ts`, `src/agentTools.ts` + `package.json` if exposed as LM tool, `docs/architecture.md` |
| New VS Code command | the matching per-domain module under `src/commands/` (wired through `src/commands.ts`), `package.json` (`contributes.commands`, `menus`), `docs/architecture.md` |
| New AI agent (LM) tool | `package.json` (`contributes.languageModelTools`), `src/agentTools.ts` (registration), `.github/copilot-instructions.md` (tool reference + workflow), `docs/architecture.md` |
| New workbook-wide analysis behavior | `src/vbaModuleAnalysis.ts` (shared module analysis core), `src/vbaWorkbookAnalysis.ts` (workbook analysis core), `src/workbookAnalysisResultsModel.ts` (results view model/copy report), `src/workbookAnalysisWebview.ts` (analysis results GUI), `src/commands/analysisCommands.ts` (command wiring + location navigation), `src/agentTools.ts` (`xlide_analyzeWorkbook`), `package.json` (command/menu/LM tool), `.github/copilot-instructions.md`, `docs/architecture.md` |
| Dependency added/removed | `package.json`, `README.md`, `docs/architecture.md` |
| New VBA language feature | `src/vbaSymbolIndex.ts` (parsing/index), `src/vbaStructuralDiagnostics.ts` (structural analysis), the matching provider subsystem module registered in `src/vbaLanguageProviders.ts`, `syntaxes/vba.tmLanguage.json` (coloring), `language-configuration/vba-language-configuration.json` (brackets/indent/folding), `docs/architecture.md` |
| New analyzer grammar rule | `src/analyzer/**` (lexer/parser), matching fixtures in `tests/`, an MS-VBAL section cite in code, a row in `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New host object-model member/type/constant | `src/analyzer/host/excelObjectModel.ts` or generated `src/analyzer/host/excelReferenceMembers.ts` (metadata transcribed or generated with provenance), `tests/vbaMemberCompletion.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum table), `docs/architecture.md` |
| New host-member call signature | `src/analyzer/host/excelObjectModel.ts` (`memberSignatures` entry, transcribed + source-verified), `tests/vbaSignatureHelp.test.ts`, `docs/spec/MS-VBAL.verification-map.md` (addendum table), `docs/architecture.md` |
| New built-in VBA runtime function/statement | `src/analyzer/runtime/vbaRuntime.ts` (signature transcribed + source-verified), `tests/vbaRuntime.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New built-in VBA runtime constant | `src/analyzer/runtime/vbaRuntime.ts` (constant/type/value transcribed + source-verified), `tests/vbaRuntime.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New symbol-graph kind/resolution rule | `src/analyzer/symbols/**`, `tests/vbaSymbolGraph.test.ts`, `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New definition/reference/rename scope rule | `src/analyzer/symbols/projectIndex.ts` (`resolveDefinition`/`resolveQualifiedDefinition`/`referenceScope`), `src/analyzer/index.ts` (barrel export), `src/vbaNavigation.ts` / `src/vbaNavigationProviders.ts` (provider wiring + span->range mapping), `tests/vbaSymbolGraph.test.ts`, `docs/architecture.md` |
| New active diagnostic rule | `src/analyzer/diagnostics/ruleMetadata.ts` plus the matching `src/analyzer/diagnostics/rules/<family>.ts` registered in `registry.ts` (rule + MS-VBAL `specReference`), the family's `tests/diagnostics/<family>.test.ts`, `src/vbaLiveDiagnostics.ts` (provider merge + any new config), `package.json` (settings), `docs/spec/MS-VBAL.verification-map.md`, `docs/architecture.md` |
| New analyzer quick fix | `src/analyzer/codeActions/diagnosticCodeActions.ts`, `src/vbaCodeActions.ts` (`CodeActionProvider` adapter only when needed), `tests/vbaCodeActions.test.ts`, `docs/roadmap_version_2.x.md`, `docs/architecture.md` |
| New editor source action | `src/vbaCodeActions.ts`, the matching per-domain module under `src/commands/`, `package.json` (`contributes.commands`/menus), focused tests when a pure helper is added, `docs/roadmap_version_2.x.md`, `docs/architecture.md` |
| New analysis suppression directive behavior | `src/analyzer/diagnostics/analysisSuppressions.ts`, `src/vbaLiveDiagnostics.ts`, `src/vbaWorkbookAnalysis.ts`, `src/commands/analysisCommands.ts`, `tests/vbaAnalysisSuppressions.test.ts`, `docs/xlide_vba_analysis_suppression_comments.md`, `docs/roadmap_version_2.x.md`, `docs/architecture.md` |
| New completion/hover resolver or rule | `src/analyzer/completion/**` or `src/analyzer/hover/**`, `src/analyzer/index.ts` (barrel export), `src/vbaMemberCompletion.ts` (provider wiring), matching `tests/vba*.test.ts`, `docs/architecture.md` |
| New signature-help rule/source | `src/analyzer/signature/signatureHelp.ts`, `src/analyzer/index.ts` (barrel export), `src/vbaMemberCompletion.ts` (`provideSignatureHelp` + `registerSignatureHelpProvider`), `tests/vbaSignatureHelp.test.ts`, `docs/architecture.md` |
| New doc-comment tag or metadata behavior | `src/analyzer/docs/**`, `src/analyzer/index.ts` (barrel export), `src/vbaDocMetadata.ts` (loader/glob), `src/vbaMemberCompletion.ts` (context wiring), `package.json` (`xlide.docs.*` settings), `tests/vbaDocComments.test.ts`, `user_guides/vba-doc-comments.md`, `docs/architecture.md` |
| Live Share RPC surface change | `src/liveShare.ts`, `docs/architecture.md` |
