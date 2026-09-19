# Getting Started With XLIDE

XLIDE lets you browse, edit, analyze, run, test, and sync VBA projects from VS
Code while keeping file-affecting actions explicit. This guide covers the
first-hour workflow for a developer opening a real macro-enabled Office file:
an Excel workbook, a Word document, a PowerPoint presentation, or an Access
database.

## Install Requirements

XLIDE requires:

- VS Code 1.95 or newer.

That is the whole list. XLIDE reads and writes the files itself, so there is no
runtime to install and no setup step to complete: open the XLIDE Activity Bar
view and it is ready.

No Office application is required for reading, editing, exporting, importing,
or analyzing VBA modules - XLIDE parses every container natively. Office COM
is required only for workflows that execute VBA: running a macro or running
unit tests, each in the file's own application (Excel, Word, PowerPoint, or
Access, chosen by the file's format).

## Open A File

1. Open a folder that contains macro-enabled Office files: Excel (`.xlsm`,
   `.xlsb`, `.xlam`, `.xltm`, `.xls`, `.xlt`, `.xla`), Word (`.docm`,
   `.dotm`, `.doc`, `.dot`), PowerPoint (`.pptm`, `.potm`, `.ppsm`,
   `.ppam`, `.ppt`, `.ppa`), or Access (`.accdb`, `.mdb`, `.mda`).
2. Open the XLIDE Activity Bar view.
3. Pick the target file from the file selector.
4. Use **Refresh** if a file was added after VS Code opened the folder.
5. Open modules from the XLIDE tree.

Modules open as writable `xlide-vba` documents and save with normal VS Code
save behavior. Access modules open read-only: Access executes compiled
p-code, so source edits there could not take effect, and XLIDE says so
rather than pretending to save.

## Group Modules Into Folders

The **Folders** button above the XLIDE tree groups a project's modules by a
comment in each module's declarations section, the Rubberduck convention:

```vba
'@Folder("Accounts.Ledger")
Option Explicit
```

Dots nest, so that module sits in `Ledger` inside `Accounts`. The annotation
is read leniently - `'@Folder Accounts.Ledger`, `'@Folder(Accounts.Ledger)`
and `'@folder("accounts.ledger")` all mean the same folder - and a module
with no annotation stays at the project's root. **Tree** switches back to the
flat list. The buttons and the `xlide.explorer.view` setting are the same
choice.

With **Auto Expand And Collapse Explorer Tree** on, the folders follow the
editor: the folders on the way to the module you are editing open, the rest
fold when you move to another folder, and a folder you open or shut yourself
stays that way until you move to a module somewhere else. Turn the setting off
to arrange the tree yourself.

The status bar shows where the cursor is: the file, the module, and the
procedure - `Book.xlsm | Module1 | Sub Recalculate`, or `(Declarations)` above
the first procedure. The tree highlights that procedure's row at the same time,
so you can see where you are in the project without leaving the editor. Both
follow the same auto expand and collapse setting.

## Edit VBA

XLIDE provides VBA syntax highlighting, canonical keyword casing, Smart Enter
for common block structures, keyword snippets, symbol navigation, hover,
signature help, and completions for project symbols, VBA runtime symbols, and
the object model of the file's own host: Excel members in a workbook, Word
members in a document, PowerPoint and Access members in theirs. `Me` in
Word's `ThisDocument` is a `Word.Document`, `wd*` constants resolve in Word
files and not in projects, and `ThisDocument` offers Word's `Document_*`
event stubs. In the code behind an Access form or report, `Me` is an
`Access.Form` or `Access.Report`, and the design's sections and controls are
members of it, so `Me.Lines.AddItem` and a bare `Requery` both resolve. A
bound form also has a member for every field of its record source, which only
the running database knows, so XLIDE never reports a name as undeclared there.
At module level it offers the design's own event stubs (`Form_Load`,
`Report_Open`), each section's (`Detail_Click`) and each control's
(`Qty_AfterUpdate`), with the parameter lists Access writes. A control named
`Order Date` is `Me.Order_Date` in code, as it is in the VBE.

**Format Document** (Shift+Alt+F) re-indents a module by its block structure,
gives every keyword its canonical casing, cases identifiers the way they are
declared, and inserts the spaces the VBE inserts around `=`, after commas, and
around operators. It never removes a space between two tokens, so aligned
comments and aligned `Const` blocks stay as they are, and it never touches a
string, a comment, or an `Attribute` line. **Format Selection** does the same
for the selected lines only, and pasted code is formatted as it lands, so a
block copied from a forum arrives indented. To format on every save, turn on
`editor.formatOnSave` for the `[xlide-vba]` language.

To bring a whole file up to date, right-click it in the XLIDE tree and choose
**Format All Modules**. XLIDE counts the modules that would change and asks
before it writes. A module with unsaved edits is formatted in its editor and
left for you to save.

## Refactor

Seven refactorings sit on the lightbulb where they apply, and in the command
palette under **XLIDE**:

| Refactoring | What it does |
| --- | --- |
| Extract Method | Selected statements become a Private procedure below the caller |
| Extract Variable | A selected expression becomes a declared local |
| Inline Variable | A local is replaced by the value it was assigned |
| Encapsulate Field | A Public variable becomes a property pair with the same name |
| Implement Interface | Stubs for every member an `Implements` promises |
| Move to Module | A procedure moves to another standard module, its doc comment with it |
| Introduce Parameter | A local becomes a parameter, and callers pass its value |

Each one either does the work or tells you why it will not, and the reasons
are worth reading: they are the cases where the rewrite would change what your
code does. Inline Variable will not bracket a compound value, because in VBA
`Foo (x)` passes by value where `Foo x` passes by reference. Extract Method
needs `Option Explicit`, because without it an undeclared name would become a
second, separate variable in the new procedure. Move to Module tells you which
Private member would be left behind.

Useful editor flows:

- Type `If ... Then` and press Enter to let Smart Enter create the matching
  block.
- Type after `As`, `As New`, or `New` for type-name completion.
- Type after `object.` for member completion when XLIDE can resolve the object
  type.
- Use **Go to Definition**, **Find All References**, and **Rename Symbol** for
  source-backed project symbols where the analyzer can bind the reference.

XLIDE is conservative. When it cannot prove a type, binding, or host member, it
prefers no hard diagnostic over a guessed error.

## Analyze Code

Use one of these actions:

- **Analyze Current Module** from the editor context menu or command palette.
- **Analyze Project** from the file tree or XLIDE sidebar.
- Live diagnostics in open VBA editors.

Analysis results use stable diagnostic codes. Red diagnostics represent
compile-equivalent errors or deterministic runtime failures that XLIDE can
prove. Warnings and information diagnostics are guidance, tracking, or softer
quality signals.

Analysis settings can be adjusted per file from the Analysis
results/settings UI. Global defaults live in XLIDE Global Settings; per-file
overrides live beside the file in `<file>.xlide_settings.json`.
See [analysis.md](analysis.md) for diagnostic tracking, severity controls,
untracked rules, and source suppressions.

## Run A Macro

To run a macro at the editor cursor, use **Run Macro at Cursor**.

Macro execution runs through COM in the file's own application - Excel, Word,
PowerPoint, or Access - and is currently Windows-only. It can attach to a
running instance of that application depending on
`xlide.officeIntegration.attachToRunning`. Treat macro execution as a
file-affecting action: save or back up important work before running code that
mutates the file, other files, external systems, or the VBA project.
**Open in Office Application** opens any of these files in its own
application, and **Open in Office Application (Read Only)** does the same
without locking it for editing (Access has no read-only open).

If the file is open in its application when you save, XLIDE cannot write it:
the application holds the file locked. `xlide.officeIntegration.coordinationMode`
decides what happens then. The default refuses and tells you which application
has it. A read-only copy XLIDE itself opened, which is how F5 leaves the file,
is closed and reopened around the save automatically.

A change saved to the file outside XLIDE, from the VBE or by a git checkout,
reaches the modules you have open. A module with no unsaved edits reloads, and
problems in the modules that call into it update. Saving unsaved edits over a
change to the same module stops at VS Code's usual prompt, where you compare
or overwrite.

## Run VBA Tests

Use **Unit Tests** from the selected file's actions or the file tree. Tests
are ordinary zero-argument `Sub` procedures in standard modules marked with
`' @xlide-test` comments, and they run in Excel, Word, PowerPoint, and Access
files alike. See [testing.md](testing.md) for the full test contract.

Test runs execute against a temporary copy of the file in an XLIDE-owned
read-only instance of the file's own application. The original file is not
saved by the test run. The bundled assertion support module must be installed
into the file before tests can run.

## Sync Modules With Files

Use export/import actions when you want normal `.bas` and `.cls` files on disk
for review, source control, or external tooling.

- **Export All Modules to Folder** writes the file's modules to a chosen folder.
- **Import Modules from Folder** previews changes before applying them back
  (available for writable containers; Access is export-only).
- **Export/Sync Current Module** exports only the active module.

The import/export diff GUI is the safest place to configure per-file
sync folders and modes. XLIDE stores those choices in
`<file>.xlide_settings.json`.
See [sync.md](sync.md) for export/import modes, true-up behavior, skipped import
cases, and sync settings.

## Documentation Comments And Metadata

Inline `'''` XML documentation comments enrich hovers, signature help, and
completion detail for your own procedures and classes.

```vba
''' <summary>Calculates an invoice total after tax.</summary>
''' <param name="Subtotal" type="Currency">Pre-tax amount.</param>
''' <param name="TaxRate" type="Double">Tax rate as a decimal.</param>
''' <returns type="Currency">Final invoice total.</returns>
Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency
End Function
```

External `.vbref.xml` metadata can describe APIs whose source XLIDE cannot see.
Use `xlide.docs.metadataGlob` to control discovery. See
[vba-doc-comments.md](vba-doc-comments.md) for module-header docs, supported XML
tags, and metadata-file examples.

## Troubleshooting Setup

If the file tree, tests, or Office actions are unavailable:

- Open the XLIDE Activity Bar view and check setup health.
- Use **Copy Diagnostics** for a quick local setup summary.
- Use **Export XLIDE Support Bundle** when you need a redacted support snapshot.
- See [support.md](support.md) for Trust Center, macro security, Office COM, and
  recovery notes.
