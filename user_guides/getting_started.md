# Getting Started With XLIDE

XLIDE lets you browse, edit, analyze, run, test, and sync VBA projects from VS
Code while keeping workbook-specific actions explicit. This guide covers the
first-hour workflow for a developer opening a real Excel macro workbook.

## Install Requirements

XLIDE requires:

- VS Code 1.95 or newer.
- Python 3.10 or newer.
- Python packages used by the XLIDE backend.

Open the XLIDE Activity Bar view and use the setup actions shown there. If
Python is not detected, set `xlide.pythonPath` in XLIDE Global Settings or use
the setup action to install the required Python packages.

Excel is not required for reading, editing, exporting, importing, or analyzing
VBA modules. Microsoft Excel COM is required only for workflows that execute
VBA, such as running a macro or running workbook tests.

## Open A Workbook

1. Open a folder that contains `.xlsm`, `.xlsb`, or `.xlam` files.
2. Open the XLIDE Activity Bar view.
3. Pick the target workbook from the workbook selector.
4. Use **Refresh** if a workbook was added after VS Code opened the folder.
5. Open modules from the XLIDE workbook tree.

Workbook modules open as writable `xlide-vba` documents. Save with normal VS
Code save behavior after editing.

## Edit VBA

XLIDE provides VBA syntax highlighting, canonical keyword casing, Smart Enter
for common block structures, keyword snippets, symbol navigation, hover,
signature help, and completions for project symbols, VBA runtime symbols, and
known Excel object-model members.

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
- **Analyze Workbook** from the workbook tree or XLIDE sidebar.
- Live diagnostics in open VBA editors.

Analysis results use stable diagnostic codes. Red diagnostics represent
compile-equivalent errors or deterministic runtime failures that XLIDE can
prove. Warnings and information diagnostics are guidance, tracking, or softer
quality signals.

Workbook analysis settings can be adjusted per workbook from the Analysis
results/settings UI. Global defaults live in XLIDE Global Settings; workbook
overrides live beside the workbook in `<workbook>.xlide_settings.json`.

## Run A Macro

To run a macro at the editor cursor, use **Run Macro at Cursor**.

Macro execution requires Microsoft Excel COM and is currently Windows-only. It
can attach to a running Excel instance depending on `xlide.attachToRunningExcel`.
Treat macro execution as a workbook-affecting action: save or back up important
work before running code that mutates sheets, files, external systems, or the
VBA project.

## Run VBA Tests

Use **Unit Tests** from the selected workbook actions or workbook tree. Tests are
ordinary zero-argument `Sub` procedures in standard modules marked with
`' @xlide-test` comments. See [testing.md](testing.md) for the full test
contract.

Test runs execute against a temporary copy of the workbook in an XLIDE-owned
read-only Excel instance. The original workbook is not saved by the test run.
The bundled assertion support module must be installed into the workbook before
tests can run.

## Sync Modules With Files

Use export/import actions when you want normal `.bas` and `.cls` files on disk
for review, source control, or external tooling.

- **Export All Modules to Folder** writes workbook modules to a chosen folder.
- **Import Modules From Folder** previews changes before applying them back to
  the workbook.
- **Export/Sync Current Module** exports only the active workbook module.

The import/export diff GUI is the safest place to configure workbook-specific
sync folders and modes. XLIDE stores those choices in
`<workbook>.xlide_settings.json`.

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
Use `xlide.docs.metadataGlob` to control discovery.

## Troubleshooting Setup

If the workbook tree, tests, or Excel actions are unavailable:

- Open the XLIDE Activity Bar view and check setup health.
- Use **Copy Diagnostics** for a quick local setup summary.
- Use **Export XLIDE Support Bundle** when you need a redacted support snapshot.
- See [support.md](support.md) for Trust Center, macro security, Excel COM, and
  recovery notes.
