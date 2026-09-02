# XLIDE Agent Instructions

XLIDE is a VS Code extension for editing Excel VBA and workbook data. The agent
surface has 18 tools for discovering projects, reading/writing VBA, analyzing
projects, running tests, syncing modules, and inspecting worksheet data.

---

## Source of Truth

**The project is the only source of truth for VBA. Exported files on disk are
not.**

Reach VBA one of two ways, both of which read and write the project itself:

- **XLIDE tools** - discover with `xlide_getProjectInfo` or
  `xlide_listModules`, read with `xlide_readModule`, write with
  `xlide_writeModule`.
- **The XLIDE virtual file system** - modules opened from the XLIDE tree live at
  `xlide-vba://` URIs. Editing and saving one of those documents writes straight
  back into the project, so ordinary editor edits are project edits.

Both paths go through XLIDE's in-process project engine, which parses and
rewrites the OLE compound file, the VBA project streams, and the OOXML package
directly. There is no external runtime and no intermediate copy on disk.

**Do not read, edit, or create exported `.bas`, `.cls`, or `.frm` files in the
project as a way of working on VBA**, even when they are present and look
current. They are one-way export artifacts: they can be stale, they can be
partial, and editing one changes nothing in the project. Treat them as source
material only when the user explicitly asks you to work on the export
artifacts, the repository sync files, or an import from disk.

Use `xlide_exportModules` only when the user wants to export/sync project
modules to disk. After direct project edits, verify with
`xlide_analyzeProject`; export to repo files is optional and user-directed.

---

## Headline Capabilities

Two XLIDE features are first-class and should be used proactively.

### 1. VBA Analysis Verification

XLIDE ships a deterministic VBA analyzer, the same engine that powers live
editor diagnostics. After every `xlide_writeModule`, or any batch of VBA edits,
call `xlide_analyzeProject` before reporting success.

The tool returns a structured JSON report with `moduleName`, `line`, `column`,
`severity`, `code`, and `message` for each finding. An empty `problems` array
means the project's VBA analysis is clean. Treat a non-empty report as a build
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

If the user has not specified a file path, call `xlide_listProjects` first to
find available `.xlsm`, `.xlsb`, and `.xlam` files.

### Step 2 - Understand the Project

Call `xlide_getProjectInfo` once per workbook. It returns sheets, VBA modules,
and named ranges in a single round-trip. Do not call `xlide_listModules` and
`xlide_listSheets` separately when `xlide_getProjectInfo` covers both. Do not
infer module state from local export files.

### Step 3 - Operate

Use the targeted tool for the task.

### Step 4 - Verify VBA Edits

After writing or editing any VBA, call `xlide_analyzeProject` and resolve any
reported problems before finishing. If workbook tests exist and Excel COM is
available, run `xlide_runVbaTests` when the task affects tested behavior.

---

## Tool Reference

### Read / Discover

| Tool | When to use |
|---|---|
| `xlide_listProjects` | User has not given a project path. |
| `xlide_getProjectInfo` | First look at any workbook: sheets, modules, and named ranges. |
| `xlide_validateProject` | Check protection/signature/project issues before risky work. |
| `xlide_listModules` | Need only the VBA module list. |
| `xlide_listSubs` | Need procedures in a specific module. |
| `xlide_readModule` | Read canonical VBA source from the project. |
| `xlide_analyzeProject` | Verify VBA syntax/analysis across the whole project. |
| `xlide_listSheets` | Need only sheet names and dimensions. |
| `xlide_readCells` | Read cached/computed cell values. |
| `xlide_readFormulas` | Read raw formula strings, such as `=SUM(A1:A10)`. |

### Write / Mutate

These tools require user confirmation.

| Tool | When to use |
|---|---|
| `xlide_writeModule` | Write or create canonical project VBA source. |
| `xlide_renameModule` | Rename a VBA module. |
| `xlide_deleteModule` | Delete a VBA module. Warn before destructive edits. |
| `xlide_createProject` | Create a new macro-enabled file. |
| `xlide_writeCells` | Write values to a cell range. |
| `xlide_exportModules` | Export/sync project modules to files on disk when explicitly requested. |
| `xlide_configureExportMode` | Set the persistent export mode for a project. |

### Execute / Test

These tools require user confirmation and may require Windows Excel COM.

| Tool | When to use |
|---|---|
| `xlide_runVbaTests` | Run marked `@xlide-test` workbook tests and inspect artifacts/results. |

---

## Key Constraints

- Keep cell values and VBA source ASCII-only unless the user explicitly asks
  otherwise; non-ASCII may not round-trip safely through every project path.
- `xlide_readCells` returns cached/computed values. Use `xlide_readFormulas` to
  see formula strings.
- `xlide_writeModule` creates the module if it does not exist. Use the source
  header only when creating class-like modules; standard modules need no special
  header.
- Document modules such as `Sheet1`, `Sheet2`, and `ThisWorkbook` cannot be
  deleted. They can only be written.
- Project write tools save the project after the call. Sync tools write files
  and project settings. Test tools write test artifacts.
