# XLIDE: VBA for VS Code

[![release](https://vsmarketplacebadges.dev/version-short/WilliamSmithE.xlide.svg?style=flat&color=orange&label=release)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)
[![installs](https://vsmarketplacebadges.dev/installs-short/WilliamSmithE.xlide.svg?style=flat&color=blue)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)
[![rating](https://vsmarketplacebadges.dev/rating-short/WilliamSmithE.xlide.svg?style=flat&color=blue)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide&ssr=false#review-details)
[![CI](https://img.shields.io/github/actions/workflow/status/WilliamSmithEdward/xlide_vscode/ci.yml?branch=main&label=CI)](https://github.com/WilliamSmithEdward/xlide_vscode/actions/workflows/ci.yml)
[![license](https://img.shields.io/badge/license-MIT-brightgreen)](LICENSE)
[![Hosts](https://img.shields.io/badge/Hosts-Excel%2C%20Word%2C%20PowerPoint%2C%20Access%2C%20VB6-blue)](#why-use-xlide)

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
(`.pptm`, `.potm`, `.ppsm`, `.ppam`, `.ppt`, `.ppa`), Access databases
(`.accdb`, `.mdb`, `.mda`), and Visual Basic 6 projects (`.vbp`, with the
`.bas`, `.cls`, `.frm`, `.ctl`, and `.pag` files they list). Navigate each
file's VBA project in the XLIDE tree, open modules directly from the file, edit
with host-aware IntelliSense (Excel modules complete against Excel's object
model, Word modules against Word's, and so on), and save changes back with
normal VS Code save behavior.
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
seven refactorings, a UserForm designer, full static analysis, live
diagnostics, file-level analysis reports, module-qualified IntelliSense,
import/export diff previews, source-control friendly `.bas`/`.cls` sync,
explicit per-file settings, performance diagnostics, and VBA unit tests that
execute through Excel, Word, PowerPoint, or Access when you need runtime
confidence.

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
   `#If` pitfalls, and more. XLIDE reads the project's own conditional
   compilation arguments, so `#If MY_FLAG Then` is decided the way the VBE
   decides it.

3. **Get IntelliSense that understands your file and its host**
   Completion and tooltips know about your modules, classes, functions,
   constants, enums, user-defined types, XML documentation comments, and the
   object model of the host the file belongs to - Excel members in a workbook,
   Word members in a document, PowerPoint and Access members in theirs.

4. **Edit VBA in a real code editor**
   Work with a file's modules in VS Code using normal editor habits: tabs,
   search, save, rename, Go to Definition, Find All References, hover,
   signature help, formatting-friendly indentation, and semantic coloring.

5. **Refactor without breaking VBA's rules**
   Extract Method, Extract Variable, Inline Variable, Encapsulate Field,
   Implement Interface, Move to Module, Introduce Parameter, and Rename. Each
   one either does the work or declines with a reason. Inline Variable, for
   example, refuses a compound value, because `Foo (x)` passes by value where
   `Foo x` passes by reference.

6. **Design UserForms in VS Code**
   A canvas, a properties pane, and the form's markup in one editor. Drag
   controls from the toolbox, resize them on the canvas, and double-click one
   to open its event handler. Every gesture is a text edit, so `Ctrl+Z` undoes
   it and nothing is written until you save. The designer reads and writes the
   form's binary storage itself, so no Office application needs to be running.

7. **Work with shapes and the macros they run**
   A worksheet's shapes sit under its module in the tree, a document's under
   ThisDocument, and a presentation's under Slides. Click one to change it in
   one form: position, text, font, fill, outline, rotation, stacking, and the
   Sub a click runs. XLIDE writes the file itself, so no Office application
   needs to be running.

8. **Edit Access databases, forms, and reports**
   Access modules can be edited, added, renamed, and deleted. Forms and reports
   open in the same designer as a UserForm, and in the code behind one `Me` is
   the Access form, with its controls as members. XLIDE writes the source and
   marks the compiled copy stale, so Access recompiles on the next open, the
   same thing its `/decompile` switch does.

9. **Open Visual Basic 6 projects**
   A `.vbp` shows in the tree like any other project, its modules get the same
   language services, and its forms open in the designer from their own `.frm`
   text. XLIDE does not build or run VB6 projects.

10. **Keep the project organized from the code**
   Put `'@Folder("Accounts.Ledger")` in a module and the **Folders** view
   groups the project by it, the Rubberduck convention. Annotations such as
   `'@PredeclaredId` and `'@Description("...")` write the hidden module
   attributes on save. The tree follows the editor, and the status bar names
   the procedure your cursor is in.

11. **Keep VBA code reviewable**
   Export modules to `.bas` and `.cls` files, preview exactly what will change,
   and use source control workflows without guessing what is inside a binary
   file.

12. **Import and export safely**
   XLIDE previews create, update, overwrite, and delete actions before applying
   them. Per-file settings live beside the file so each project can keep its
   own sync rules.

13. **Analyze the whole file**
   Run file-wide analysis over any container - workbook, document,
   presentation, or database - and review findings in a dedicated report
   instead of hunting through modules one by one.

14. **Run macros and VBA tests when Office is available**
   On Windows with Microsoft Office installed, XLIDE runs macros and
   `@xlide-test` unit tests through explicit automation of the file's own
   application - Excel, Word, PowerPoint, or Access.

15. **Give AI assistants real file context**
   XLIDE exposes tools for file discovery, VBA reads/writes, analysis,
   tests, sheet/cell access, formulas, shapes on any Office surface - a
   worksheet, a slide, a document body, a header - and the macros they run,
   and module sync so agents can work from the actual Office file instead of
   stale exported copies.

---

## Who It Is For

XLIDE is useful if you:

- Are learning VBA for a class, internship, first automation project, or career
  change.
- Use Excel, Word, PowerPoint, or Access heavily and want to start programming
  without losing sight of the file.
- Are a student, analyst, accountant, engineer, researcher, or operations user
  turning repeated Office work into reusable automation.
- Own business-critical Office files with VBA.
- Maintain a Visual Basic 6 application and want a modern editor for it.
- Maintain shared macros for finance, operations, reporting, engineering, or
  internal tools.
- Want better visibility into old VBA projects before changing them.
- Need to review VBA code with teammates.
- Prefer VS Code editing, search, navigation, and source control.
- Want AI help that can inspect the file directly.

The Office application remains where the code runs. XLIDE adds a better
workspace around the VBA project.

---

## Get Started

1. Install the extension:
   [XLIDE on the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)

2. Open a folder in VS Code that contains a macro-enabled Office file - a
   workbook, document, presentation, or database - or a VB6 `.vbp`.

3. Expand your file in the XLIDE view, open a module, and start editing.
   Press `Ctrl+S` to save code back to the file.

There is nothing else to install: XLIDE reads and writes the files itself, with
no external runtime and no setup step.

Start here for a guided first-hour walkthrough:
[Getting started with XLIDE](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md)

---

## Everyday Workflow

### Edit VBA modules

Open a file in the XLIDE tree, choose a module, edit it in VS Code, and save.
XLIDE writes the module back to the file.

### Fix red squiggles

Live diagnostics appear as you type. For a bigger pass, run **XLIDE: Analyze
Project** from the project menu and review the analysis report.

Guide:
[Analysis and ignores](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md)

### Refactor

Put the cursor on a symbol or select a block and use the lightbulb, or run a
refactoring from the Command Palette.

### Design a form

Open a UserForm, an Access form or report, or a VB6 `.frm` from the tree and
it opens in the designer. Save writes the form back to the file.

### Work with shapes

Open a worksheet's module in the tree and its shapes are in a Shapes folder
above its procedures. A Word document's are under ThisDocument, by story, and
a presentation's are under Slides. Click a shape to open it in the shape
editor. Its menu links it to a Sub, unlinks it, opens the Sub, or deletes
the shape. A Shapes folder, a slide, or a story adds one. Word cannot run a
macro from a shape, so its shapes offer no link.

### Open a VB6 project

Add a `.vbp` to the folder and it appears in the tree. Its modules get the
same editing, analysis, and navigation as an Office project.

Guide:
[Visual Basic 6](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/vb6.md)

### Sync code with files

Use **Export Modules to Folder** to create or update `.bas` and `.cls` files.
Use **Import Modules from Folder** to bring reviewed files back into the
selected file. XLIDE shows a file-scoped preview before applying changes.

Guide:
[Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md)

### Run tests

Mark VBA procedures as tests, then run them from XLIDE. Test execution
requires the file's own Office application on Windows.

Guide:
[Testing VBA projects](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md)

### Use AI assistants safely

XLIDE gives compatible VS Code AI agents tools to inspect and edit a file's VBA,
run analysis, export modules, and read and change worksheet cells and shapes
with explicit file context.

Guide:
[Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md)

---

## What XLIDE Adds To VS Code

- File and module tree for Office files.
- Writable VBA editors backed by the file.
- VBA syntax highlighting and semantic coloring.
- IntelliSense for project symbols, VBA language features, and the object
  model of the file's host: Excel, Word, PowerPoint, Access, or VB6.
- Hover, signature help, Go to Definition, Find All References, and Rename
  Symbol.
- Seven refactorings: Extract Method, Extract Variable, Inline Variable,
  Encapsulate Field, Implement Interface, Move to Module, and Introduce
  Parameter.
- A UserForm designer with a canvas, toolbox, properties pane, tab order, and
  the form's markup in one editor.
- Writable Access modules, and Access forms and reports in the same designer.
- Shapes in the tree for worksheets, slides, and documents, and a shape
  editor for every property and the Sub a click runs.
- Visual Basic 6 projects in the tree, with their forms in the designer.
- A Folders view driven by `'@Folder` annotations, and a status bar that
  names the current procedure.
- Annotations that write hidden module attributes on save.
- Conditional compilation arguments read from the project.
- Smart Enter and block snippets for common VBA structures.
- Format Document, Format Selection, and Format All Modules: indentation by
  block, keyword casing, and the spacing the VBE applies. Pasted code is
  formatted as it lands.
- Live diagnostics and file-wide analysis.
- Dead-code findings: unused variables, uncalled private procedures, and
  unreachable code, faded in the editor, with quick fixes.
- Per-file analysis tracking and rule controls.
- Previewable module import/export.
- Compare a module, or every module in a file, with any git commit without
  exporting. The tree marks the modules that changed since the last commit.
- Macro and unit-test execution on Windows, in the file's own application.
- Saves that work while the file is open in its application: XLIDE can close
  it there, save, and reopen it where you were.
- A read-only workbook in Excel that follows your saves, keeping its sheet,
  selection and scroll.
- A save refused because the file is held names what holds it: Excel, Word,
  OneDrive, a backup agent.
- A VBA project added to a macro-enabled file that has none yet, the way its
  application would add one.
- Support bundle and diagnostics commands for troubleshooting.
- Optional performance snapshot command for debugging slow workflows.

---

## Outside VS Code

XLIDE puts this surface in an editor. If you want an agent working the
same way without one, [xlide-mcp](https://github.com/WilliamSmithEdward/xlide_mcp)
is an MCP server over the same ground: the VBA, the UserForms and the
Power Query inside Excel, Word, PowerPoint and Access files, reachable
by anything that speaks the Model Context Protocol. It carries the
content-token guard and the analysis build gate this extension uses, so
an agent editing a workbook is held to the same rules you are. Run it
beside XLIDE and its edits show in the XLIDE tree with a before/after
diff and Keep / Revert, like the edits of XLIDE's own agent tools.

```bash
uvx --from "xlide-mcp[live]" xlide-mcp --root /path/to/your/files
```

## Requirements

Required for normal browsing, editing, analysis, import, and export:

- Visual Studio Code 1.95 or newer.

That is the whole list. XLIDE parses and rewrites each container (OLE compound
file, VBA project, OOXML package, and Access database) natively in the extension,
so there is no runtime, interpreter, or library to install.

Required only for running VBA code from XLIDE:

- Windows.
- The file's own Office application: Excel, Word, PowerPoint, or Access.
- Macro security settings that allow the code to run.

Reading, editing, analyzing, importing, and exporting VBA do not require any
Office application.

Setup and recovery guide:
[Safety, trust, and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md)

---

## User Guides

| Need | Guide |
|---|---|
| First setup and first workflow | [Getting started](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md) |
| Diagnostics, analysis results, and ignored findings | [Analysis and ignores](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md) |
| Import/export and module sync | [Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md) |
| Writing and running VBA tests | [Testing VBA projects](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md) |
| XML documentation comments for IntelliSense | [VBA documentation comments](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/vba-doc-comments.md) |
| AI-agent and CI workflows | [Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md) |
| Trust model, Office setup, support bundles, and recovery | [Safety, trust, and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md) |
| VB6 projects, forms, and limits | [Visual Basic 6](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/vb6.md) |
| Full guide index | [User guides README](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/README.md) |

---

## Common Commands

Open the Command Palette and type `XLIDE` to find these commands:

| Command | Use it when you want to |
|---|---|
| `XLIDE: Analyze Project` | Review a file's issues in one report. |
| `XLIDE: Analyze Current Module` | Check only the module you are editing. |
| `XLIDE: Open Designer` | Open a form in the designer. |
| `XLIDE: Extract Method` | Refactor the selection. The other six refactorings sit beside it. |
| `XLIDE: Export All Modules to Folder` | Save a file's code as reviewable source files. |
| `XLIDE: Import Modules from Folder` | Bring reviewed module files back into the file. |
| `XLIDE: Compare Module with Git HEAD` | See what changed in a module since the last commit. |
| `XLIDE: Compare File with Git HEAD` | List the modules that changed since the last commit. |
| `XLIDE: Restore Module from Git HEAD` | Put one module back to its committed text. |
| `XLIDE: Show Module History` | List the commits that changed one module, and see what each did. |
| `Format Document` (Shift+Alt+F) | Re-indent a module and fix keyword casing and spacing. |
| `XLIDE: Format All Modules` | Do the same for every module in a file, after a count. |
| `XLIDE: Unit Tests` | Run marked tests through the file's own application. |
| `XLIDE: Open in Office Application` (Ctrl+Alt+O) | Open the selected file in its own application. Ctrl+Alt+Shift+O opens it read-only. |
| `XLIDE: Add VBA Project` | Give a macro-enabled file with no VBA in it yet its first project. The file's row in the tree offers it. |
| `XLIDE: Copy Diagnostics` | Copy setup and environment details for support. |
| `XLIDE: Export Support Bundle` | Create a troubleshooting bundle. |
| `XLIDE: Copy Performance Snapshot` | Copy recent timing data when something feels slow. |

---

## Notes And Limits

- XLIDE reads and writes VBA modules through the Office file itself. Keep
  normal backups for important files, especially before large sync operations.
- Running macros or tests uses Office automation and is Windows-only.
- Ctrl+Alt+O has no tree row to go on, so it works out which file you mean:
  the one picked in the XLIDE sidebar, or the one whose module is in front,
  or the one you last had a module open from, or the only one in the
  workspace. With several files and none of those, it does nothing rather
  than guess. The status bar names the file it will open, or says
  "ambiguous". Rebind it under File > Preferences > Keyboard Shortcuts.
- Add VBA Project works on the macro-enabled Office Open XML formats
  (.xlsm, .xlsb, .docm, .pptm and their templates and add-ins). A legacy
  .xls, .doc or .ppt file gets its first macro in its application. Excel
  keeps a project only while it holds code, so a workbook saved in Excel
  before any is written loses it again.
- The UserForm designer keeps every property it does not name itself, so a
  form you have not edited saves back unchanged. Access designs work the same
  way.
- The tree lists a workbook's sheets in a Sheets folder, in tab order, with
  each module named the way the VBA editor names it: Sheet1 (Budget). Excel
  gives a worksheet its module only once its VBA editor is opened after the
  sheet was added; a sheet with no module and no shapes sits in Sheets With
  No Modules or Shapes. ActiveX controls are listed but not edited, and
  pictures and charts can be changed but not added. The shape tools need the
  Office Open XML formats; a binary .xlsb, .xls, .doc, or .ppt file lists
  its sheets but shows no shapes.
- After XLIDE writes to an Access database, Access recompiles it on the next
  open, so all the VBA in it has to compile, including code XLIDE did not
  touch.
- A VB6 designer edit rewrites the header block at the top of the `.frm` and
  leaves the code below alone. Pictures and other `.frx` records are read,
  never written. XLIDE does not build or run VB6 projects.
- Exported `.bas` and `.cls` files are useful for review and source control, but
  the Office file remains the source of truth unless you explicitly import files
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
| `npm run test:integration` | Drive a real VS Code against a copy of the fixture workbook. The first run downloads a VS Code build into `.vscode-test/`. |
| `npm run test:office` | Optional checks against real Excel, Word, PowerPoint and Access on Windows, on scratch copies of the fixtures. Excel runs in instances the suite starts; Word, PowerPoint and Access are skipped while you have them open. Nothing you have running is attached to. |
| `npm run package` | Build a production bundle. |
| `npm run vsix` | Create a versioned `.vsix` in `dist/`. |
| `npm run test:oracle:vbe` | Optional Excel/VBE behavior checks. Run oracle checks sequentially. |
| `npm run test:oracle:twinbasic` | Optional VB6 parity checks against twinBASIC. |

Architecture reference:
[XLIDE architecture](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/architecture.md)

---

## Support Open Source

XLIDE is open-source software. If it saves you time or helps your team keep VBA
projects maintainable, support helps keep the project moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
