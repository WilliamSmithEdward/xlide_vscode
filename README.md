# XLIDE: VBA for VS Code

[![release](https://vsmarketplacebadges.dev/version-short/WilliamSmithE.xlide.svg?style=flat&color=orange&label=release)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)
[![installs](https://vsmarketplacebadges.dev/installs-short/WilliamSmithE.xlide.svg?style=flat&color=blue)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)
[![rating](https://vsmarketplacebadges.dev/rating-short/WilliamSmithE.xlide.svg?style=flat&color=blue)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide&ssr=false#review-details)
[![CI](https://img.shields.io/github/actions/workflow/status/WilliamSmithEdward/xlide_vscode/ci.yml?branch=main&label=CI)](https://github.com/WilliamSmithEdward/xlide_vscode/actions/workflows/ci.yml)
[![license](https://img.shields.io/badge/license-MIT-brightgreen)](LICENSE)
[![Office](https://img.shields.io/badge/Office-Excel%2C%20Word%2C%20PowerPoint%2C%20Access-blue)](#new-in-400-every-office-vba-host)

[Install XLIDE from the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)\
\
[See XLIDE's Sister Project Which Puts XLIDE Directly in the VBA Editor](https://github.com/WilliamSmithEdward/xlide_vbide)

---

XLIDE gives Microsoft Office VBA projects a modern VS Code workspace. Read and write
directly to VBA files directly without COM.

Add macro-enabled Office files to your VS Code project and XLIDE detects them
automatically: Excel workbooks, templates, and add-ins (`.xlsm`, `.xlsb`,
`.xlam`, `.xltm`, `.xls`, `.xlt`, `.xla`), Word documents and templates
(`.docm`, `.dotm`, `.doc`, `.dot`), PowerPoint presentations and add-ins
(`.pptm`, `.potm`, `.ppsm`, `.ppam`, `.ppt`, `.ppa`), and Access databases
(`.accdb`, `.mdb`, `.mda`, read-only). Navigate each file's VBA project in the XLIDE
tree, open modules directly from the file, edit with host-aware IntelliSense
(Excel modules complete against Excel's object model, Word modules against
Word's, and so on), and save changes back with normal VS Code save behavior.
When you want a file-based workflow, use the full local disk push/pull
workflow for the selected file: preview detailed side-by-side diffs, export
modules to `.bas` and `.cls` files, commit them to your favorite version
control platform, and import reviewed files back.

For new programmers, XLIDE makes the VBA language easier to approach.
It shows useful completion lists, explains symbols, highlights likely mistakes,
and keeps each file's project structure visible while you learn how Office
automation fits together.

For experienced VBA developers, XLIDE brings serious engineering workflows to
existing files: project-wide symbol navigation, rename and reference tools,
full static analysis, live diagnostics, file-level analysis reports,
module-qualified IntelliSense, import/export diff previews, source-control
friendly `.bas`/`.cls` sync, explicit per-file settings, performance
diagnostics, and VBA unit tests that execute through Excel, Word, or
PowerPoint when you need runtime confidence.

For everyone, XLIDE opens a new agentic AI surface for Office VBA. Compatible
AI assistants can inspect the real file, read and write VBA modules, analyze
code, run tests, inspect worksheet data, and sync modules with files through
explicit XLIDE tools. That means AI help can work with the document itself, not
just copied snippets or stale exports, making Office automation more
reviewable, testable, and collaborative.

---

## Why Use XLIDE

1. **Make VBA easier to learn and maintain**
   VBA is powerful, but its rules, project structure, and the host object
   models can feel hidden. XLIDE makes them visible with project navigation,
   IntelliSense, hover explanations, signature help, semantic coloring, and
   immediate feedback while you type.

2. **Catch VBA mistakes earlier**
   XLIDE shows live red and yellow squiggles for many high-confidence VBA
   problems: missing block closers, duplicate names, undeclared variables,
   invalid parameter/property shapes, array misuse, bad `Set` usage, common
   `#If` pitfalls, and more.

3. **Get IntelliSense that understands your file and its host**
   Completion and tooltips know about your modules, classes, functions,
   constants, enums, user-defined types, XML documentation comments, and the
   object model of the host the file belongs to - Excel members in a workbook,
   Word members in a document, PowerPoint and Access members in theirs.

4. **Edit project VBA in a real code editor**
   Work with project modules in VS Code using normal editor habits: tabs,
   search, save, rename, Go to Definition, Find All References, hover,
   signature help, formatting-friendly indentation, and semantic coloring.

5. **Keep project code reviewable**
   Export modules to `.bas` and `.cls` files, preview exactly what will change,
   and use source control workflows without guessing what is inside a binary
   project.

6. **Import and export safely**
   XLIDE previews create, update, overwrite, and delete actions before applying
   them. Project-specific settings live beside the project so each project can
   keep its own sync rules.

7. **Analyze the whole file**
   Run file-wide analysis over any container - workbook, document,
   presentation, or database - and review findings in a dedicated report
   instead of hunting through modules one by one.

8. **Run macros and VBA tests when Office is available**
   On Windows with Microsoft Office installed, XLIDE runs macros and
   `@xlide-test` unit tests through explicit automation of the file's own
   application - Excel, Word, or PowerPoint.

9. **Give AI assistants real file context**
   XLIDE exposes tools for file discovery, VBA reads/writes, analysis,
   tests, sheet/cell access, formulas, and module sync so agents can work from
   the actual Office file instead of stale exported copies.

---

## Who It Is For

XLIDE is useful if you:

- Are learning VBA for a class, internship, first automation project, or career
  change.
- Use Excel heavily and want to start programming without losing sight of the
  project.
- Are a student, analyst, accountant, engineer, researcher, or operations user
  turning repeated spreadsheet work into reusable automation.
- Own business-critical Excel workbooks with VBA.
- Maintain shared macros for finance, operations, reporting, engineering, or
  internal tools.
- Want better visibility into old VBA projects before changing them.
- Need to review project code with teammates.
- Prefer VS Code editing, search, navigation, and source control.
- Want AI help that can inspect the project directly.

Excel remains where the workbook runs. XLIDE adds a better workspace around the
VBA project.

---

## Get Started

1. Install the extension:
   [XLIDE on the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)

2. Open a folder in VS Code that contains an `.xlsm`, `.xlsb`, or `.xlam`
   project.

3. Expand your project in the XLIDE view, open a module, and start editing.
   Press `Ctrl+S` to save code back to the project.

There is nothing else to install: XLIDE reads and writes projects itself, with
no external runtime and no setup step.

Start here for a guided first-hour walkthrough:
[Getting started with XLIDE](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md)

---

## Everyday Workflow

### Edit VBA modules

Open a project in the XLIDE tree, choose a module, edit it in VS Code, and save.
XLIDE writes the module back to the project.

### Fix red squiggles

Live diagnostics appear as you type. For a bigger pass, run **XLIDE: Analyze
Project** from the project menu and review the analysis report.

Guide:
[Analysis and ignores](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md)

### Sync code with files

Use **Export Modules to Folder** to create or update `.bas` and `.cls` files.
Use **Import Modules from Folder** to bring reviewed files back into the
selected project. XLIDE shows a project-scoped preview before applying
changes.

Guide:
[Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md)

### Design UserForms

Open a UserForm and the designer opens with it: the canvas, a properties pane,
and the form's markup in one tab. Drag from the toolbox to add controls, drag
and resize on the canvas, and double-click a control for its event handler.
Ctrl-click or Shift-click builds a multi-selection, and so does dragging a band
across the form; the selection aligns, resizes, moves, copies (Ctrl+C, Ctrl+V)
and deletes as one. Every gesture is an ordinary text edit of the markup, so
Ctrl+Z undoes it and the project is written only when you save.

The designer is native: it reads and writes the form's binary storage directly
and never needs Excel open.

A VB6 form (`.frm`, and a `.ctl` or `.pag`) opens in the same designer from
its project's Designer row, or through Open With. The file is the document:
a gesture rewrites the header block at the top of the file and leaves the code
below it alone, Ctrl+Z is text undo, and save is the file's own save. Menus
draw as a menu bar, and a control from an OCX draws at its bounds under its
name.

### Run tests

Mark VBA procedures as project tests, then run them from XLIDE. Test execution
requires Microsoft Excel on Windows.

Guide:
[Testing VBA projects](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md)

### Use AI assistants safely

XLIDE gives compatible VS Code AI agents tools to inspect and edit project VBA,
run analysis, export modules, and read worksheet data with explicit workbook
context.

Guide:
[Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md)

---

## What XLIDE Adds To VS Code

- Project and module tree for office 365 files.
- Writable VBA editors backed by the project.
- VBA syntax highlighting and semantic coloring.
- IntelliSense for workbook symbols, VBA language features, and Excel host
  objects.
- Hover, signature help, Go to Definition, Find All References, and Rename
  Symbol.
- Smart Enter and block snippets for common VBA structures.
- Live diagnostics and project-wide analysis.
- Project-specific analysis tracking and rule controls.
- Previewable module import/export.
- Macro and workbook-test execution on Windows with Excel.
- Support bundle and diagnostics commands for troubleshooting.
- Optional performance snapshot command for debugging slow workflows.

---

## Requirements

Required for normal browsing, editing, analysis, import, and export:

- Visual Studio Code 1.95 or newer.

That is the whole list. XLIDE parses and rewrites the project container
(OLE compound file, VBA project, and OOXML package) natively in the extension,
so there is no runtime, interpreter, or library to install.

Required only for running VBA code from XLIDE:

- Windows.
- Microsoft Excel.
- Project macro settings that allow the code to run.

Reading, editing, analyzing, importing, and exporting project VBA do not require
Excel automation.

Setup and recovery guide:
[Safety, trust, and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md)

---

## User Guides

| Need | Guide |
|---|---|
| First setup and first project workflow | [Getting started](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md) |
| Diagnostics, analysis results, and ignored findings | [Analysis and ignores](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md) |
| Import/export and module sync | [Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md) |
| Writing and running project tests | [Testing VBA projects](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md) |
| XML documentation comments for IntelliSense | [VBA documentation comments](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/vba-doc-comments.md) |
| AI-agent and CI workflows | [Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md) |
| Trust model, Excel setup, support bundles, and recovery | [Safety, trust, and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md) |
| Full guide index | [User guides README](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/README.md) |

---

## Common Commands

Open the Command Palette and type `XLIDE` to find these commands:

| Command | Use it when you want to |
|---|---|
| `XLIDE: Analyze Project` | Review a file's issues in one report. |
| `XLIDE: Analyze Current Module` | Check only the module you are editing. |
| `XLIDE: Export All Modules to Folder` | Save a file's code as reviewable source files. |
| `XLIDE: Import Modules from Folder` | Bring reviewed module files back into the file. |
| `XLIDE: Unit Tests` | Run marked tests through the file's own application. |
| `XLIDE: Open Workbook in Excel` | Open the selected workbook in Excel. |
| `XLIDE: Copy Diagnostics` | Copy setup and environment details for support. |
| `XLIDE: Export Support Bundle` | Create a troubleshooting bundle. |
| `XLIDE: Copy Performance Snapshot` | Copy recent timing data when something feels slow. |

---

## Notes And Limits

- XLIDE reads and writes VBA modules through the project file. Keep normal
  backups for important projects, especially before large sync operations.
- Running macros or tests uses Excel automation and is Windows-only.
- UserForms are designed in XLIDE's own designer, which reads and writes the
  form's binary storage directly. Its markup is the editable face of that
  storage: a property XLIDE does not name is carried through untouched.
- VB6 forms are designed from their `.frm` text. A gesture rewrites the header
  in the designer's own layout, so a form the designer never touched saves
  byte for byte and a real gesture changes header lines only. A multi-line
  text goes to the form's `.frx` as a new record when the form is saved;
  pictures and every other sidecar record are read, never written.
- Exported `.bas` and `.cls` files are useful for review and source control, but
  the project remains the source of truth unless you explicitly import files
  back into it.

---

## For Developers And Contributors

Repository:
[https://github.com/WilliamSmithEdward/xlide_vscode](https://github.com/WilliamSmithEdward/xlide_vscode)

```bash
git clone https://github.com/WilliamSmithEdward/xlide_vscode.git
cd xlide_vscode

npm install
npm run compile
```

Press `F5` in VS Code to launch an Extension Development Host.

Useful development commands:

| Command | Purpose |
|---|---|
| `npm run compile` | Type-check and build the extension bundle. |
| `npm run watch` | Rebuild while developing. |
| `npm test` | Run the Vitest suite. |
| `npm run package` | Build a production bundle. |
| `npm run vsix` | Create a versioned `.vsix` in `dist/`. |
| `npm run test:oracle:vbe` | Optional Excel/VBE behavior checks. Run oracle checks sequentially. |

Architecture reference:
[XLIDE architecture](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/architecture.md)

---

## Support Open Source

XLIDE is open-source software. If it saves you time or helps your team keep VBA
projects maintainable, support helps keep the project moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
