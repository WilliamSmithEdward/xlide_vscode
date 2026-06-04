# XLIDE Safety, Trust, And Support

XLIDE works with real macro-enabled workbooks. This guide explains what XLIDE can
read, write, or execute, which actions need Excel, and how to collect support
details without exposing workbook source by default.

## Trust Model

XLIDE separates actions into four broad categories.

| Category | Examples | User expectation |
|---|---|---|
| Read-only workbook inspection | list workbooks, list modules, read modules, read cell values, analyze code | Does not save the workbook. |
| Workbook mutation | save a module, create/rename/delete a module, import modules, write cells, create a workbook | Requires an explicit command, save, UI action, or approved agent tool call. |
| VBA execution | run macro at cursor, run workbook tests | Requires an explicit run action and Microsoft Excel COM. |
| Local support/export artifacts | export support bundle, export modules, write test artifacts | Writes files to a selected or configured local folder. |

XLIDE should not infer consent from merely opening a workbook or viewing a
module. Mutating actions are tied to visible commands, normal editor save
operations, workbook GUIs, or agent tool calls that are expected to request user
approval before writes.

## Workbook Mutation

Common workbook-mutating actions include:

- Saving an edited `xlide-vba` document.
- Adding, renaming, or deleting a VBA module.
- Importing modules from a folder.
- Writing cells through XLIDE tools.
- Creating a new macro-enabled workbook.
- Installing or updating the `XlideAssert.bas` support module for tests.

Before applying bulk import/export changes, prefer the diff preview. It shows
which modules will change, be skipped, or be removed before you apply the
operation.

For high-value workbooks, keep your own backup or source-control checkpoint
before large imports, module deletion, or workbook-structure changes. XLIDE
records in-session command outcomes and change summaries for support, but a
backup is still the cleanest recovery path after a bad edit.

## Macro And Test Execution

Macro and test execution require Microsoft Excel COM and are Windows-only.

Macro execution can affect the workbook, the file system, external data sources,
and any systems the macro touches. Run macros only when you trust the workbook
code.

Workbook tests run against a temporary copy in an XLIDE-owned Excel instance and
close without saving the original workbook. Installing the assertion support
module does modify the original workbook because it adds or updates
`XlideAssert.bas`.

The production test runner does not require Excel's "Trust access to the VBA
project object model" setting for normal test execution. XLIDE's internal
developer VBE oracle is different and may require that setting because it creates
disposable VBE modules to verify language behavior.

## Excel Trust Center And Macro Security

Excel macro security still applies. If Excel blocks macros for a workbook,
XLIDE cannot make that workbook's VBA run safely around Excel's policy.

Check these Excel settings when macro or test execution is blocked:

- The workbook is in a trusted location or is unblocked according to your
  organization's macro policy.
- Macros are allowed for the workbook.
- Protected View is not blocking execution.
- The workbook is not waiting on a modal prompt.
- Excel COM is registered on the machine.

Only enable "Trust access to the VBA project object model" when a workflow
explicitly requires VBE automation. Normal XLIDE editing uses workbook module I/O
and normal test execution should not require that setting.

## Workbook Open State

Some workflows can behave differently when the workbook is already open in
Excel. XLIDE has a setting, `xlide.attachToRunningExcel`, that controls whether
Windows Excel workflows try to reuse an already-running Excel instance.

If a write, run, or test operation appears blocked:

- Close other Excel windows that have the same workbook open.
- Save pending Excel changes before running XLIDE operations.
- Check whether the workbook is read-only or locked by another process.
- Retry from a clean Excel state if Excel was left at a prompt or in edit mode.

## Support Bundle

Use **Export XLIDE Support Bundle** when you need a local diagnostic snapshot.
The bundle is designed to help debug setup, COM availability, settings, recent
XLIDE command outcomes, and analysis/test/sync state without including workbook
source by default.

By default, XLIDE redacts workbook paths and path-like settings. Optional
sections may include selected logs or anonymized workbook analysis summaries
when you explicitly opt in.

Use **Copy Diagnostics** for a smaller setup-focused summary when the issue is
basic environment health.

## Recovery Checklist

If something goes wrong:

1. Stop and save a copy of the workbook before retrying destructive actions.
2. Close Excel instances that may still hold the workbook.
3. Reopen VS Code and refresh the XLIDE workbook tree.
4. Use the import/export diff preview to inspect workbook-vs-file differences.
5. Use the support bundle if the failure involves setup, COM, test host, or
   unclear command outcomes.
6. Restore from your own workbook backup or source-controlled module export when
   the desired state is known.
