# XLIDE: VBA for VS Code

XLIDE gives Excel VBA projects a modern VS Code workspace.

Add one or more `.xlsm`, `.xlsb`, or `.xlam` workbooks to your VS Code project
and XLIDE detects them automatically. Navigate each workbook's VBA project in
the XLIDE tree, open modules directly from the workbook, edit with
workbook-aware IntelliSense, and save changes back to the workbook with normal
VS Code save behavior. When you want a file-based workflow, use the full local
disk push/pull workflow for the selected workbook: preview detailed
side-by-side diffs, export workbook modules to `.bas` and `.cls` files, commit
them to your favorite version control platform, and import reviewed files back
into that workbook.

For new programmers, XLIDE makes the VBA language easier to approach.
It shows useful completion lists, explains symbols, highlights likely mistakes,
and keeps the workbook's project structure visible while you learn how Excel
automation fits together.

For experienced VBA developers, XLIDE brings serious engineering workflows to
existing workbooks: project-wide symbol navigation, rename and reference tools,
full static analysis, live diagnostics, workbook-level analysis reports,
module-qualified IntelliSense, import/export diff previews, source-control
friendly `.bas`/`.cls` sync, explicit workbook settings, performance
diagnostics, and workbook unit tests that can execute through Excel when you
need runtime confidence.

For everyone, XLIDE opens a new agentic AI surface for Excel. Compatible AI
assistants can inspect the real workbook, read and write VBA modules, analyze
code, run tests, inspect worksheet data, and sync modules with files through
explicit XLIDE tools. That means AI help can work with the workbook itself, not
just copied snippets or stale exports, making Excel automation more reviewable,
testable, and collaborative.

[Install XLIDE from the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)

---

## Why Use XLIDE

1. **Make VBA easier to learn and maintain**
   VBA is powerful, but its rules, project structure, and Excel object model can
   feel hidden. XLIDE makes them visible with workbook navigation,
   IntelliSense, hover explanations, signature help, semantic coloring, and
   immediate feedback while you type.

2. **Catch VBA mistakes earlier**
   XLIDE shows live red and yellow squiggles for many high-confidence VBA
   problems: missing block closers, duplicate names, undeclared variables,
   invalid parameter/property shapes, array misuse, bad `Set` usage, common
   `#If` pitfalls, and more.

3. **Get IntelliSense that understands your workbook**
   Completion and tooltips know about your modules, classes, functions,
   constants, enums, user-defined types, XML documentation comments, and a broad
   set of Excel object-model members.

4. **Edit workbook VBA in a real code editor**
   Work with workbook modules in VS Code using normal editor habits: tabs,
   search, save, rename, Go to Definition, Find All References, hover,
   signature help, formatting-friendly indentation, and semantic coloring.

5. **Keep workbook code reviewable**
   Export modules to `.bas` and `.cls` files, preview exactly what will change,
   and use source control workflows without guessing what is inside a binary
   workbook.

6. **Import and export safely**
   XLIDE previews create, update, overwrite, and delete actions before applying
   them. Workbook-specific settings live beside the workbook so each project can
   keep its own sync rules.

7. **Analyze the whole workbook**
   Run workbook-wide analysis and review findings in a dedicated report instead
   of hunting through modules one by one.

8. **Run macros and workbook tests when Excel is available**
   On Windows with Microsoft Excel installed, XLIDE can run macros and
   `@xlide-test` workbook tests through explicit Excel automation.

9. **Give AI assistants real workbook context**
   XLIDE exposes tools for workbook discovery, VBA reads/writes, analysis,
   tests, sheet/cell access, formulas, and module sync so agents can work from
   the actual workbook instead of stale exported files.

---

## Who It Is For

XLIDE is useful if you:

- Are learning VBA for a class, internship, first automation project, or career
  change.
- Use Excel heavily and want to start programming without losing sight of the
  workbook.
- Are a student, analyst, accountant, engineer, researcher, or operations user
  turning repeated spreadsheet work into reusable automation.
- Own business-critical Excel workbooks with VBA.
- Maintain shared macros for finance, operations, reporting, engineering, or
  internal tools.
- Want better visibility into old VBA projects before changing them.
- Need to review workbook code with teammates.
- Prefer VS Code editing, search, navigation, and source control.
- Want AI help that can inspect the workbook directly.

Excel remains where the workbook runs. XLIDE adds a better workspace around the
VBA project.

---

## Get Started

1. Install the extension:
   [XLIDE on the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)

2. Install Python 3.10 or newer if you do not already have it:
   [Download Python](https://www.python.org/downloads/)

3. Open a folder in VS Code that contains an `.xlsm`, `.xlsb`, or `.xlam`
   workbook.

4. Open the XLIDE view in the Activity Bar. If setup is incomplete, use the
   sidebar buttons to choose Python and install the required Python libraries.

5. Expand your workbook, open a module, and start editing. Press `Ctrl+S` to
   save code back to the workbook.

Start here for a guided first-hour walkthrough:
[Getting started with XLIDE](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md)

---

## Everyday Workflow

### Edit VBA modules

Open a workbook in the XLIDE tree, choose a module, edit it in VS Code, and save.
XLIDE writes the module back to the workbook.

### Fix red squiggles

Live diagnostics appear as you type. For a bigger pass, run **XLIDE: Analyze
Workbook** from the workbook menu and review the analysis report.

Guide:
[Analysis and ignores](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md)

### Sync code with files

Use **Export Modules to Folder** to create or update `.bas` and `.cls` files.
Use **Import Modules from Folder** to bring reviewed files back into the
selected workbook. XLIDE shows a workbook-scoped preview before applying
changes.

Guide:
[Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md)

### Run tests

Mark VBA procedures as workbook tests, then run them from XLIDE. Test execution
requires Microsoft Excel on Windows.

Guide:
[Testing VBA workbooks](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md)

### Use AI assistants safely

XLIDE gives compatible VS Code AI agents tools to inspect and edit workbook VBA,
run analysis, export modules, and read worksheet data with explicit workbook
context.

Guide:
[Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md)

---

## What XLIDE Adds To VS Code

- Workbook and module tree for `.xlsm`, `.xlsb`, and `.xlam` files.
- Writable VBA editors backed by the workbook.
- VBA syntax highlighting and semantic coloring.
- IntelliSense for workbook symbols, VBA language features, and Excel host
  objects.
- Hover, signature help, Go to Definition, Find All References, and Rename
  Symbol.
- Smart Enter and block snippets for common VBA structures.
- Live diagnostics and workbook-wide analysis.
- Workbook-specific analysis tracking and rule controls.
- Previewable module import/export.
- Macro and workbook-test execution on Windows with Excel.
- Support bundle and diagnostics commands for troubleshooting.
- Optional performance snapshot command for debugging slow workflows.

---

## Requirements

Required for normal browsing, editing, analysis, import, and export:

- Visual Studio Code 1.95 or newer.
- Python 3.10 or newer.
- Python packages installed by XLIDE setup: `pyOpenVBA` and `openpyxl`.

Required only for running VBA code from XLIDE:

- Windows.
- Microsoft Excel.
- Workbook macro settings that allow the code to run.

Reading, editing, analyzing, importing, and exporting workbook VBA do not require
Excel automation.

Setup and recovery guide:
[Safety, trust, and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md)

---

## User Guides

| Need | Guide |
|---|---|
| First setup and first workbook workflow | [Getting started](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md) |
| Diagnostics, analysis results, and ignored findings | [Analysis and ignores](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md) |
| Import/export and module sync | [Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md) |
| Writing and running workbook tests | [Testing VBA workbooks](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md) |
| XML documentation comments for IntelliSense | [VBA documentation comments](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/vba-doc-comments.md) |
| AI-agent and CI workflows | [Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md) |
| Trust model, Excel setup, support bundles, and recovery | [Safety, trust, and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md) |
| Full guide index | [User guides README](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/README.md) |

---

## Common Commands

Open the Command Palette and type `XLIDE` to find these commands:

| Command | Use it when you want to |
|---|---|
| `XLIDE: Analyze Workbook` | Review workbook-wide issues in one report. |
| `XLIDE: Analyze Current Module` | Check only the module you are editing. |
| `XLIDE: Export All Modules to Folder` | Save workbook code as reviewable source files. |
| `XLIDE: Import Modules from Folder` | Bring reviewed module files back into the workbook. |
| `XLIDE: Unit Tests` | Run marked workbook tests through Excel. |
| `XLIDE: Open Workbook in Excel` | Open the selected workbook in Excel. |
| `XLIDE: Copy Diagnostics` | Copy setup and environment details for support. |
| `XLIDE: Export Support Bundle` | Create a troubleshooting bundle. |
| `XLIDE: Copy Performance Snapshot` | Copy recent timing data when something feels slow. |

---

## Notes And Limits

- XLIDE reads and writes VBA modules through the workbook file. Keep normal
  backups for important workbooks, especially before large sync operations.
- Running macros or tests uses Excel automation and is Windows-only.
- UserForm designer files are not edited directly. XLIDE can work with the VBA
  code-behind where supported.
- Live Share guests can view and edit modules the host has already opened, but
  only the host can browse the XLIDE workbook tree.
- Exported `.bas` and `.cls` files are useful for review and source control, but
  the workbook remains the source of truth unless you explicitly import files
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

python -m venv .venv
.venv\Scripts\activate
pip install -r python/requirements.txt
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

Roadmaps:

- [Version 2.1.0 completed red-squiggle closeout](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/roadmap_version_2.1.0.md)
- [Version 2.2.0 static-analysis completeness roadmap](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/roadmap_version_2.2.0.md)
- [Version 3.0.0 deferred product/backlog roadmap](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/roadmap_version_3.0.0.md)

---

## Support Open Source

XLIDE is open-source software. If it saves you time or helps your team keep VBA
workbooks maintainable, support helps keep the project moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
