# XLIDE: VBA for VS Code

[![release](https://vsmarketplacebadges.dev/version-short/WilliamSmithE.xlide.svg?style=flat&color=orange&label=release)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)
[![installs](https://vsmarketplacebadges.dev/installs-short/WilliamSmithE.xlide.svg?style=flat&color=blue)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)
[![rating](https://vsmarketplacebadges.dev/rating-short/WilliamSmithE.xlide.svg?style=flat&color=blue)](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide&ssr=false#review-details)
[![CI](https://img.shields.io/github/actions/workflow/status/WilliamSmithEdward/xlide_vscode/ci.yml?branch=main&label=CI)](https://github.com/WilliamSmithEdward/xlide_vscode/actions/workflows/ci.yml)
[![license](https://img.shields.io/badge/license-MIT-brightgreen)](LICENSE)
[![Hosts](https://img.shields.io/badge/Hosts-Excel%2C%20Word%2C%20PowerPoint%2C%20Access%2C%20VB6-blue)](#what-it-does)

[Install from the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide)\
\
[XLIDE's sister project puts XLIDE inside the VBA editor itself](https://github.com/WilliamSmithEdward/xlide_vbide)

---

Edit the VBA inside your Office files from VS Code. Drop an `.xlsm` in a folder
and its modules appear in a tree; open one, edit it, press `Ctrl+S`, and the
workbook has your code.

XLIDE reads and writes the project container itself - the compound file, the
VBA project, the OOXML package - so nothing else has to be installed and Excel
never has to be open. It runs on a build server. The module you open is the one
in the file, not an export of it.

## Files it opens

| Host | Extensions |
|---|---|
| Excel | `.xlsm` `.xlsb` `.xlam` `.xltm` `.xls` `.xlt` `.xla` |
| Word | `.docm` `.dotm` `.doc` `.dot` |
| PowerPoint | `.pptm` `.potm` `.ppsm` `.ppam` `.ppt` `.ppa` |
| Access | `.accdb` `.mdb` `.mda` |
| Visual Basic 6 | `.vbp`, with the `.bas` `.cls` `.frm` `.ctl` `.pag` files it names |

---

## What it does

**Editing that knows the host.** Completion, hover and signature help resolve
against your own modules and against the object model of the application the
file belongs to - Excel members in a workbook, Word members in a document,
Access members in a database. `Me` in `ThisDocument` is a `Word.Document`.
Go to Definition, Find All References and Rename Symbol work across the
project.

**Diagnostics while you type.** Missing block closers, duplicate declarations,
undeclared variables, bad `Set` usage, array misuse, argument shapes that will
not compile, and the `#If` traps. XLIDE reads a project's own conditional
compilation arguments, so `#If MY_FLAG Then` is decided the way the VBE decides
it rather than analyzed both ways. **XLIDE: Analyze Project** puts the whole
file's findings in one report.

**Refactorings.** Extract Method, Extract Variable, Inline Variable,
Encapsulate Field, Implement Interface, Move to Module, Introduce Parameter,
and Rename. Each either does the work or declines with a reason, and the
reasons matter in VBA: Inline Variable will not bracket a compound value,
because `Foo (x)` passes by value where `Foo x` passes by reference.

**A designer for UserForms.** Canvas, properties pane and markup in one tab.
Drag from the toolbox, resize on the canvas, double-click for the event
handler; multi-select aligns, moves and deletes as one. Every gesture is a text
edit of the markup, so `Ctrl+Z` undoes it and nothing is written until you
save. It reads and writes the form's binary storage directly and never needs
Excel running. VB6 forms open in the same designer from their own `.frm` text.

**Access databases, natively.** An Access database keeps its VBA in rows of a
system table and runs a compiled copy of it rather than the source, so XLIDE
writes the source and marks that copy stale. Access recompiles on the next
open, which is what its own `/decompile` switch does. Modules can be edited,
added, renamed and deleted, and so can forms and reports: they open in the same
designer as a UserForm, with the same canvas, toolbox, tab order and property
pane. Renaming one moves the design, its catalog row, the navigation pane's row
and the module Access binds its code to, all four together. Access never has to
be running.

**Folders, from the code.** Put `'@Folder("Accounts.Ledger")` in a module and
the **Folders** view groups the project by it - the Rubberduck convention,
accepted however it is spaced, cased or quoted. The tree follows the editor:
folders on the way to the module you are editing open, and the status bar names
the procedure your cursor is in.

**Annotations that write the hidden attributes.** A VBA module carries
attributes the code pane never shows and the editor gives no way to set:
`VB_PredeclaredId`, `VB_Description`, `VB_UserMemId`. Write `'@PredeclaredId`
or `'@Description("...")` in the code, where it can be reviewed like anything
else, and XLIDE writes the attribute on save.

**Source control that sees inside the binary.** Export modules to `.bas` and
`.cls`, review the diff, import the reviewed files back. Every import previews
what it will create, update, overwrite and delete before it does any of it.

**Tests and macros, when Office is there.** Mark procedures with
`@xlide-test` and run them through the file's own application. Windows with
Office installed; everything above works without it.

**Tools for AI assistants.** Compatible agents get explicit tools to list and
read modules, write them, run analysis, read cells and formulas, and sync with
files - working against the real document instead of a pasted copy.

---

## Get started

1. [Install XLIDE](https://marketplace.visualstudio.com/items?itemName=WilliamSmithE.xlide).
2. Open a folder containing a macro-enabled Office file or a `.vbp`.
3. Expand the project in the XLIDE view, open a module, edit, and save.

There is no runtime or library to add. For a walk through the first hour,
see
[Getting started](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md).

**Requirements:** Visual Studio Code 1.95 or newer. Running macros and tests
additionally needs Windows and the relevant Office application.

---

## Commands

Open the Command Palette and type `XLIDE`.

| Command | For |
|---|---|
| `XLIDE: Analyze Project` | Every finding in the file, in one report |
| `XLIDE: Analyze Current Module` | Just the module you are in |
| `XLIDE: Export All Modules to Folder` | Code as reviewable source files |
| `XLIDE: Import Modules from Folder` | Reviewed files back into the project |
| `XLIDE: Unit Tests` | Run marked tests through the file's own application |
| `XLIDE: Open Workbook in Excel` | Hand the workbook to Excel |
| `XLIDE: Export Support Bundle` | A redacted bundle for troubleshooting |
| `XLIDE: Copy Performance Snapshot` | Recent timings, when something feels slow |

---

## Guides

| Need | Guide |
|---|---|
| First setup and first project | [Getting started](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/getting_started.md) |
| Diagnostics, analysis and ignores | [Analysis](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/analysis.md) |
| Import, export and module sync | [Import and export](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/sync.md) |
| Writing and running tests | [Testing](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/testing.md) |
| Doc comments for IntelliSense | [Documentation comments](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/vba-doc-comments.md) |
| VB6 projects, forms and limits | [Visual Basic 6](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/vb6.md) |
| AI-agent and CI workflows | [Automation and CI](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/automation.md) |
| Trust, Excel setup, recovery | [Safety and support](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/user_guides/support.md) |

---

## Limits worth knowing

- XLIDE writes to your project file. Keep backups of anything important,
  particularly before a large import.
- Access recompiles a database the first time it opens after XLIDE has
  written to it, so all the VBA in it has to compile - including code XLIDE
  did not touch.
- Macros and tests need Windows and Office. Editing, analysis, import and
  export do not.
- The UserForm designer keeps every property it does not name itself, so a
  form you have not edited saves back unchanged. An Access design is the same:
  the properties it shows are the ones its own type's schema names, and every
  other record is carried through untouched.
- A VB6 gesture rewrites the header block at the top of the `.frm` and leaves
  the code below alone. Multi-line text goes to the `.frx` as a new record on
  save; pictures and other sidecar records are read, never written. XLIDE does
  not build or run VB6 projects.
- Exported `.bas` and `.cls` files are for review. The project file stays the
  source of truth until you import them back.

---

## Contributing

```bash
git clone https://github.com/WilliamSmithEdward/xlide_vscode.git
cd xlide_vscode
npm install
npm run compile
```

`F5` launches an Extension Development Host.

| Command | Purpose |
|---|---|
| `npm run compile` | Type-check and build the bundle |
| `npm run watch` | Rebuild while developing |
| `npm test` | Run the Vitest suite |
| `npm run vsix` | Build a versioned `.vsix` in `dist/` |
| `npm run test:oracle:vbe` | Optional Excel/VBE behavior checks, run sequentially |

[Architecture reference](https://github.com/WilliamSmithEdward/xlide_vscode/blob/main/docs/architecture.md)

---

## Support the project

XLIDE is MIT-licensed and open source. If it saves you time, support helps keep
it moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
