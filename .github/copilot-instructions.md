# XLIDE Agent Instructions

XLIDE is a VS Code extension for editing Excel VBA and cell data. The agent has 16 tools for reading and writing workbooks.

---

## Source of Truth

Workbook-backed XLIDE operations are canonical. When the user asks to inspect or edit VBA in a workbook, discover modules with `xlide_getWorkbookInfo` or `xlide_listModules`, read with `xlide_readModule`, and write with `xlide_writeModule`. Do not edit exported `.bas`, `.cls`, or `.frm` repo files as a substitute for workbook edits unless the user explicitly asks to work on export artifacts or repository sync files.

Use `xlide_exportModules` only when the user wants to export/sync workbook modules to disk. After direct workbook edits, verify with `xlide_lintWorkbook`; export to repo files is optional and user-directed.

---

## Headline capabilities

Two XLIDE features are first-class and you should use them proactively:

### 1. Real-time VBA syntax + lint verification
XLIDE ships a pure, false-positive-free VBA analyzer (the same engine that powers the live editor diagnostics). **After every `xlide_writeModule` (or any batch of VBA edits), call `xlide_lintWorkbook` to verify lint passes before you report success.** It returns a structured JSON report of every error/warning with `moduleName`, `line`, `column`, `severity`, `code`, and `message`. An empty `problems` array means the workbook's VBA is clean. Rules cover block balance, unterminated strings, duplicate declarations, `Const` assignment, malformed procedure headers, argument-count mismatches, unbalanced parentheses, `Sub`/`Function`-not-defined and more - each carrying an MS-VBAL spec reference. Treat a non-empty report as a build failure and fix it.

### 2. Custom XML tooltips / definitions / metadata for VBA
XLIDE supports Visual-Studio-style documentation that drives hovers and call tips. You can author it two ways, and **developer-defined metadata overrides the built-in library**:
- **Inline `'''` XML doc-comments** directly above any `Sub`/`Function`/`Property`/`Type`/`Enum`/`Declare`/module-variable, using `<summary>`, `<param name="...">`, `<returns>`, `<remarks>`, `<example>` (a plain `''' text` line is treated as the summary). When you write or generate VBA, add these doc-comments so the symbols get rich tooltips.
- **External `*.vbref.xml` metadata files** placed anywhere in the workspace (glob configurable via `xlide.docs.metadataGlob`), documenting `<member name="Module.Symbol">` (or bare `Symbol`) entries with the same XML vocabulary - ideal for documenting host members or sharing docs across a team, and `<signature>` can give a call tip to a procedure that cannot otherwise be resolved.

The full standard lives in `docs/vba-doc-comments.md`.

---

## Canonical Workflow

### Step 1 — Discover workbooks
If the user has not specified a file path, always call `xlide_listWorkbooks` first to find available `.xlsm`/`.xlsb`/`.xlam` files.

### Step 2 — Understand the workbook
Call `xlide_getWorkbookInfo` once per workbook. It returns sheets (name + used dimensions), VBA modules (name + type), and named ranges in a single round-trip. Do NOT call `xlide_listModules` + `xlide_listSheets` separately when `xlide_getWorkbookInfo` covers both. Do not infer module state from local export files.

### Step 3 — Operate
Use the targeted tool for the task (see tool reference below). Prefer specific tools over `xlide_runOpenpyxl` when a specific tool exists.

### Step 4 — Verify VBA edits
After writing or editing any VBA, call `xlide_lintWorkbook` and resolve any reported problems before finishing.

---

## Tool Reference

### Discovery (no confirmation required)
| Tool | When to use |
|---|---|
| `xlide_listWorkbooks` | User hasn't given a file path |
| `xlide_getWorkbookInfo` | First look at any workbook — sheets + modules + named ranges |
| `xlide_listModules` | Need only the VBA module list |
| `xlide_listSubs` | Need procedures in a specific module |
| `xlide_listSheets` | Need only sheet names and dimensions |
| `xlide_readModule` | Read canonical VBA source from the workbook-backed XLIDE module |
| `xlide_lintWorkbook` | **Verify VBA syntax/lint across the whole workbook** - run after editing modules; returns per-problem `moduleName`/`line`/`column`/`severity`/`code`/`message` |
| `xlide_readCells` | Read computed cell values (formulas already evaluated) |
| `xlide_readFormulas` | Read raw formula strings (e.g. `=SUM(A1:A10)`) — use when reproducing or auditing spreadsheet logic |

### Write / Modify (require user confirmation)
| Tool | When to use |
|---|---|
| `xlide_writeModule` | Write or **create** canonical workbook VBA source - if the module name does not exist it is created automatically |
| `xlide_renameModule` | Rename a VBA module |
| `xlide_deleteModule` | Delete a VBA module (irreversible — warn the user) |
| `xlide_writeCells` | Write values to a cell range |
| `xlide_runOpenpyxl` | Anything not covered above: styling, fills, fonts, borders, column widths, number formats, charts, conditional formatting, sheet operations, named ranges — full openpyxl API |
| `xlide_exportModules` | Export/sync workbook modules to files on disk when explicitly requested |
| `xlide_configureExportMode` | Set the persistent export mode for a workbook |

---

## xlide_runOpenpyxl Usage

The code runs with these variables available:
- `wb` — the open `openpyxl.Workbook`
- `openpyxl` — the full openpyxl module
- `json` — stdlib json
- `result` — assign your return value here

**Read-only query** — set `save: false` to avoid writing:
```python
result = {
    "sheets": wb.sheetnames,
    "named_ranges": [nr.name for nr in wb.defined_names.definedName],
}
```

**Styling example:**
```python
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
ws = wb["Sheet1"]
ws["A1"].font = Font(bold=True, size=14)
ws["A1"].fill = PatternFill("solid", fgColor="4472C4")
ws["A1"].alignment = Alignment(horizontal="center")
result = "styled"
```

**Column width / row height:**
```python
ws = wb["Sheet1"]
ws.column_dimensions["A"].width = 20
ws.row_dimensions[1].height = 30
result = "done"
```

---

## Key Constraints

- **ASCII only** in all cell values and VBA source — no Unicode, emoji, or accented characters. They will mangle on round-trip.
- `xlide_readCells` returns cached/computed values (Excel last-saved result). Use `xlide_readFormulas` to see the formula string.
- `xlide_writeModule` creates the module if it doesn't exist. Use `kind` in the source header only for class modules; standard modules need no special header.
- Document modules (Sheet1, Sheet2, ThisWorkbook) cannot be deleted — only written.
- All write tools auto-save the workbook after every call.
