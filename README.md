# XLIDE - Excel VBA for VS Code

Edit Excel VBA code directly in VS Code. Browse modules in a sidebar tree,
edit with syntax highlighting and symbol navigation (Go to Definition,
Find All References, Rename Symbol), save changes back to the `.xlsm` file
with Ctrl+S, and expose every operation to GitHub Copilot via the Language
Model API.

---

## Requirements

- **VS Code** 1.95+
- **Python 3.10+** -- the VBA read/write backend runs as a child process
- Python packages: `pyOpenVBA >= 3.0.1`, `openpyxl >= 3.1.0`

No COM automation, no Office installation, no win32com -- works on Windows,
macOS, Linux, and remote containers.

---

## Development setup

```bash
git clone https://github.com/WilliamSmithEdward/xlide_vscode.git
cd xlide_vscode

# TypeScript side
npm install
npm run compile        # type-check + esbuild bundle -> out/extension.js

# Python side (optional venv)
python -m venv .venv
.venv\Scripts\activate   # or: source .venv/bin/activate
pip install -r python/requirements.txt
```

Press **F5** in VS Code to launch an Extension Development Host with the
extension loaded and the watch compiler running.

---

## Architecture

```
xlide_vscode/
  src/
    extension.ts            # activate() -- wires everything together
    pythonBridge.ts         # JSON-RPC 2.0 client over child_process stdio
    xlideFileSystem.ts      # xlide-vba:// virtual FileSystemProvider
    xlsmExplorer.ts         # Sidebar TreeDataProvider
    commands.ts             # VS Code command registrations
    agentTools.ts           # vscode.lm.registerTool() for Copilot
    moduleDump.ts           # Shared export-to-folder logic (UI + AI lane)
    vbaSymbolIndex.ts       # In-memory cross-module symbol index
    vbaLanguageProviders.ts # DocumentSymbol / Definition / References / Rename
  python/
    server.py               # JSON-RPC 2.0 server (stdin/stdout, newline-delimited)
    xlide/
      vba_io.py             # pyOpenVBA wrappers -- listModules, readModule, writeModule
      excel_io.py           # openpyxl wrappers -- readCells, writeCells
  syntaxes/
    vba.tmLanguage.json     # TextMate grammar (MS-VBAL spec-accurate)
  language-configuration/
    vba-language-configuration.json   # Brackets, indent rules, folding
  walkthrough/              # Markdown content for VS Code Getting Started tab
  docs/
    architecture.md         # Full architecture reference
```

### Key design decisions

| Decision | Rationale |
|---|---|
| Long-lived Python process | Amortises ~200 ms Python startup across all requests |
| `FileSystemProvider` over `TextDocumentContentProvider` | Read/write virtual FS -- Ctrl+S triggers `writeFile` with no custom save command |
| Virtual URI scheme `xlide-vba://` | Decouples workbook path + module name from the editor's file concept |
| Shared `moduleDump.ts` | Export logic is single-source-of-truth for both UI commands and Copilot agent tools |
| No COM / no Office | Portability -- pyOpenVBA reads the OVBA binary format directly |
| Confirmation on write tools | Prevents AI agents from silently mutating production workbooks |

### JSON-RPC methods (Python bridge)

| Method | Params | Returns |
|---|---|---|
| `listModules` | `{ path }` | `[{ name, type }]` |
| `listSubs` | `{ path, module }` | `[{ name, kind, line }]` |
| `readModule` | `{ path, module }` | `{ source }` |
| `writeModule` | `{ path, module, source }` | `{}` |
| `renameModule` | `{ path, module, newName }` | `{}` |
| `deleteModule` | `{ path, module }` | `{}` |
| `readCells` | `{ path, sheet, range }` | `{ values }` |
| `writeCells` | `{ path, sheet, startCell, data }` | `{}` |

### VBA language ID

Registered as `vba` in `package.json` with extensions `.bas`, `.cls`, `.frm`.
The TextMate grammar in `syntaxes/vba.tmLanguage.json` is scoped to
`source.vba` and covers all reserved identifiers from MS-VBAL v20250520
(section 3.3.5.2: statement-keywords, marker-keywords, operator-identifiers,
reserved-names, special-forms, reserved-type-identifiers, literal-identifiers,
def-type directives, and implementation-reserved identifiers).

---

## Build commands

| Command | Purpose |
|---|---|
| `npm run compile` | Type-check + dev bundle |
| `npm run watch` | Incremental type-check + esbuild watch |
| `npm run package` | Production bundle (minified) |
| `vsce package --no-dependencies` | Build `.vsix` for distribution |

---

## Copilot agent tools

| Tool name | Reference | Reads/Writes | Confirm |
|---|---|---|---|
| `xlide_listModules` | `#xlideListModules` | R | No |
| `xlide_listSubs` | `#xlideListSubs` | R | No |
| `xlide_readModule` | `#xlideReadModule` | R | No |
| `xlide_writeModule` | `#xlideWriteModule` | W | Yes |
| `xlide_readCells` | `#xlideReadCells` | R | No |
| `xlide_writeCells` | `#xlideWriteCells` | W | Yes |
| `xlide_exportModules` | `#xlideExportModules` | W | Yes |
| `xlide_configureExportMode` | `#xlideConfigureExportMode` | W | Yes |

---

## Per-workbook export config

Stored beside each workbook as `<workbookname>.extension.repo.json`:

```json
{
  "exportFolder": "C:/absolute/path/to/export",
  "exportMode": "trueUp",
  "managedFiles": ["Module1.bas", "Sheet1.cls"]
}
```

`trueUp` (default) -- replace existing, add new, delete stale files tracked in
`managedFiles`. `replaceExistingOnly` -- only replaces files already on disk.

---

## Further reading

- [docs/architecture.md](docs/architecture.md) -- full architecture reference
- [MS-VBAL specification](https://docs.microsoft.com/openspecs/office_file_formats/ms-vbal)
- [pyOpenVBA](https://github.com/DecimalTurn/pyOpenVBA)
