# XLIDE – Architecture

## Overview

XLIDE is a VS Code extension that turns Excel macro files (`.xlsm`, `.xlsb`, `.xlam`) into first-class editable documents. VBA modules open in the editor like normal source files, Ctrl+S writes them back into the workbook, and six Language Model tools expose every operation to Copilot and other VS Code AI agents.

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
| `writeModule` | `path`, `module`, `source` | — | `{ok}` |
| `renameModule` | `path`, `module`, `newName` | — | `{ok}` |
| `deleteModule` | `path`, `module` | — | `{ok}` |
| `readCells` | `path`, `sheet`, `range` | — | `{data: [[…]]}` |
| `writeCells` | `path`, `sheet`, `startCell`, `data` | — | `{ok}` |

Errors are returned as `{"error": {"code": -32000, "message": "…"}}`.

---

## Module export

`moduleDump.ts` is the single source of truth for export/config behavior.

Both lanes call into this shared implementation:

- UI commands (`xlide.exportModulesToFolder`, `xlide.configureExportMode`)
- AI tools (`xlide_exportModules`, `xlide_configureExportMode`)

The export operation exports all modules from a workbook by reading each module live over JSON-RPC (`listModules` then `readModule` per module).

- User picks a destination folder with a folder picker.
- Output file extension is `.bas` for standard modules and `.cls` for class/document modules.
- Export mode is per-workbook and persisted in the workbook-local JSON config:
  - `trueUp` (default): replace existing, add new, remove no-longer-existing modules
  - `replaceExistingOnly`: replace files that already exist; do not add missing files; do not remove stale files
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
| `xlide_listModules` | `#xlideListModules` | none | No |
| `xlide_listSubs` | `#xlideListSubs` | none | No |
| `xlide_readModule` | `#xlideReadModule` | none | No |
| `xlide_writeModule` | `#xlideWriteModule` | saves .xlsm | Yes |
| `xlide_readCells` | `#xlideReadCells` | none | No |
| `xlide_writeCells` | `#xlideWriteCells` | saves .xlsm | Yes |
| `xlide_exportModules` | `#xlideExportModules` | writes export files + updates workbook JSON config | Yes |
| `xlide_configureExportMode` | `#xlideConfigureExportMode` | updates workbook JSON config | Yes |

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

`src/vbaLanguageProviders.ts` registers four providers against the `vba`
language under the `xlide-vba` scheme:

| Provider | Behavior |
|---|---|
| `DocumentSymbolProvider` | Outlines the current module from `parseVbaModule` |
| `DefinitionProvider` | Resolves an identifier across all modules in the workbook; honors `Module.Member` qualifiers and `Private` visibility |
| `ReferenceProvider` | Word-boundary search across all modules, skipping string literals and apostrophe comments |
| `RenameProvider` | `prepareRename` checks the identifier is a known procedure; `provideRenameEdits` returns a `WorkspaceEdit` that rewrites every module; VS Code applies the edit and Ctrl+S persists each module through the virtual filesystem |

The index also subscribes to `onDidSaveTextDocument` for `xlide-vba://` URIs so
the cache stays in sync with user edits.

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
| New VBA language feature | `src/vbaSymbolIndex.ts` (parsing/index), `src/vbaLanguageProviders.ts` (provider), `syntaxes/vba.tmLanguage.json` (coloring), `language-configuration/vba-language-configuration.json` (brackets/indent/folding), `docs/architecture.md` |
