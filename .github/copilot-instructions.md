# XLIDE Agent Instructions

XLIDE is a VS Code extension for editing Excel VBA and cell data. The agent has 15 tools for reading and writing workbooks.

---

## Canonical Workflow

### Step 1 — Discover workbooks
If the user has not specified a file path, always call `xlide_listWorkbooks` first to find available `.xlsm`/`.xlsb`/`.xlam` files.

### Step 2 — Understand the workbook
Call `xlide_getWorkbookInfo` once per workbook. It returns sheets (name + used dimensions), VBA modules (name + type), and named ranges in a single round-trip. Do NOT call `xlide_listModules` + `xlide_listSheets` separately when `xlide_getWorkbookInfo` covers both.

### Step 3 — Operate
Use the targeted tool for the task (see tool reference below). Prefer specific tools over `xlide_runOpenpyxl` when a specific tool exists.

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
| `xlide_readModule` | Read VBA source of a module |
| `xlide_readCells` | Read computed cell values (formulas already evaluated) |
| `xlide_readFormulas` | Read raw formula strings (e.g. `=SUM(A1:A10)`) — use when reproducing or auditing spreadsheet logic |

### Write / Modify (require user confirmation)
| Tool | When to use |
|---|---|
| `xlide_writeModule` | Write or **create** VBA source — if the module name does not exist it is created automatically |
| `xlide_renameModule` | Rename a VBA module |
| `xlide_deleteModule` | Delete a VBA module (irreversible — warn the user) |
| `xlide_writeCells` | Write values to a cell range |
| `xlide_runOpenpyxl` | Anything not covered above: styling, fills, fonts, borders, column widths, number formats, charts, conditional formatting, sheet operations, named ranges — full openpyxl API |
| `xlide_exportModules` | Export all VBA modules to files on disk |
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
