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
    commands.ts         Command handlers: open, new, rename, delete module
    agentTools.ts       Six LanguageModelTool registrations for AI agent use

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
