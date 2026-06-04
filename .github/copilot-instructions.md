# XLIDE Agent Instructions

XLIDE is a VS Code extension for editing Excel VBA and workbook data. The agent
surface has 19 tools for discovering workbooks, reading/writing VBA, analyzing
projects, running tests, syncing modules, and inspecting worksheet data.

---

## Source of Truth

Workbook-backed XLIDE operations are canonical. When the user asks to inspect or
edit VBA in a workbook, discover with `xlide_getWorkbookInfo` or
`xlide_listModules`, read with `xlide_readModule`, and write with
`xlide_writeModule`. Do not edit exported `.bas`, `.cls`, or `.frm` repo files
as a substitute for workbook edits unless the user explicitly asks to work on
export artifacts or repository sync files.

Use `xlide_exportModules` only when the user wants to export/sync workbook
modules to disk. After direct workbook edits, verify with
`xlide_analyzeWorkbook`; export to repo files is optional and user-directed.

---

## Headline Capabilities

Two XLIDE features are first-class and should be used proactively.

### 1. VBA Analysis Verification

XLIDE ships a deterministic VBA analyzer, the same engine that powers live
editor diagnostics. After every `xlide_writeModule`, or any batch of VBA edits,
call `xlide_analyzeWorkbook` before reporting success.

The tool returns a structured JSON report with `moduleName`, `line`, `column`,
`severity`, `code`, and `message` for each finding. An empty `problems` array
means the workbook's VBA analysis is clean. Treat a non-empty report as a build
failure and fix it before finishing when the task is an edit.

### 2. VBA Documentation Metadata

XLIDE supports Visual-Studio-style documentation that drives hovers, completion
detail, and call tips. Developer-defined metadata overrides built-in library
metadata.

- Inline `'''` XML doc-comments can be placed directly above any
  `Sub`/`Function`/`Property`/`Type`/`Enum`/`Declare`/module-variable, using
  `<summary>`, `<param name="...">`, `<returns>`, `<remarks>`, and `<example>`.
  A plain `''' text` line is treated as the summary.
- External `*.vbref.xml` metadata files can document `<member
  name="Module.Symbol">` or bare `Symbol` entries with the same XML vocabulary.
  External `<signature>` entries can provide call tips for procedures XLIDE
  cannot otherwise resolve.

The full standard lives in `docs/vba-doc-comments.md`.

---

## Canonical Workflow

### Step 1 - Discover Workbooks

If the user has not specified a file path, call `xlide_listWorkbooks` first to
find available `.xlsm`, `.xlsb`, and `.xlam` files.

### Step 2 - Understand the Workbook

Call `xlide_getWorkbookInfo` once per workbook. It returns sheets, VBA modules,
and named ranges in a single round-trip. Do not call `xlide_listModules` and
`xlide_listSheets` separately when `xlide_getWorkbookInfo` covers both. Do not
infer module state from local export files.

### Step 3 - Operate

Use the targeted tool for the task. Prefer specific tools over
`xlide_runOpenpyxl` when a specific tool exists.

### Step 4 - Verify VBA Edits

After writing or editing any VBA, call `xlide_analyzeWorkbook` and resolve any
reported problems before finishing. If workbook tests exist and Excel COM is
available, run `xlide_runVbaTests` when the task affects tested behavior.

---

## Tool Reference

### Read / Discover

| Tool | When to use |
|---|---|
| `xlide_listWorkbooks` | User has not given a workbook path. |
| `xlide_getWorkbookInfo` | First look at any workbook: sheets, modules, and named ranges. |
| `xlide_validateWorkbook` | Check protection/signature/project issues before risky work. |
| `xlide_listModules` | Need only the VBA module list. |
| `xlide_listSubs` | Need procedures in a specific module. |
| `xlide_readModule` | Read canonical VBA source from the workbook. |
| `xlide_analyzeWorkbook` | Verify VBA syntax/analysis across the whole workbook. |
| `xlide_listSheets` | Need only sheet names and dimensions. |
| `xlide_readCells` | Read cached/computed cell values. |
| `xlide_readFormulas` | Read raw formula strings, such as `=SUM(A1:A10)`. |

### Write / Mutate

These tools require user confirmation.

| Tool | When to use |
|---|---|
| `xlide_writeModule` | Write or create canonical workbook VBA source. |
| `xlide_renameModule` | Rename a VBA module. |
| `xlide_deleteModule` | Delete a VBA module. Warn before destructive edits. |
| `xlide_createWorkbook` | Create a new macro workbook. |
| `xlide_writeCells` | Write values to a cell range. |
| `xlide_runOpenpyxl` | Use openpyxl for sheet operations not covered by a specific tool. |
| `xlide_exportModules` | Export/sync workbook modules to files on disk when explicitly requested. |
| `xlide_configureExportMode` | Set the persistent export mode for a workbook. |

### Execute / Test

These tools require user confirmation and may require Windows Excel COM.

| Tool | When to use |
|---|---|
| `xlide_runVbaTests` | Run marked `@xlide-test` workbook tests and inspect artifacts/results. |

---

## xlide_runOpenpyxl Usage

The code runs with these variables available:

- `wb` - the open `openpyxl.Workbook`
- `openpyxl` - the full openpyxl module
- `json` - stdlib json
- `result` - assign your return value here

Read-only query: set `save: false` to avoid writing.

```python
result = {
    "sheets": wb.sheetnames,
    "named_ranges": [nr.name for nr in wb.defined_names.definedName],
}
```

Styling example:

```python
from openpyxl.styles import Font, PatternFill, Alignment

ws = wb["Sheet1"]
ws["A1"].font = Font(bold=True, size=14)
ws["A1"].fill = PatternFill("solid", fgColor="4472C4")
ws["A1"].alignment = Alignment(horizontal="center")
result = "styled"
```

Column width / row height:

```python
ws = wb["Sheet1"]
ws.column_dimensions["A"].width = 20
ws.row_dimensions[1].height = 30
result = "done"
```

---

## Key Constraints

- Keep cell values and VBA source ASCII-only unless the user explicitly asks
  otherwise; non-ASCII may not round-trip safely through every workbook path.
- `xlide_readCells` returns cached/computed values. Use `xlide_readFormulas` to
  see formula strings.
- `xlide_writeModule` creates the module if it does not exist. Use the source
  header only when creating class-like modules; standard modules need no special
  header.
- Document modules such as `Sheet1`, `Sheet2`, and `ThisWorkbook` cannot be
  deleted. They can only be written.
- Workbook write tools save the workbook after the call. Sync tools write files
  and workbook settings. Test tools write test artifacts.
