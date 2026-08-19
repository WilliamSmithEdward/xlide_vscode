# XLIDE Safety, Trust, And Support

XLIDE works with real macro-enabled Office files. This guide explains what
XLIDE can read, write, or execute, which actions need an Office application,
and how to collect support details without exposing VBA source by default.

## Trust Model

XLIDE separates actions into four broad categories.

| Category | Examples | User expectation |
|---|---|---|
| Read-only inspection | list files, list modules, read modules, read cell values, analyze code | Does not save the file. |
| File mutation | save a module, create/rename/delete a module, import modules, write cells, create a file | Requires an explicit command, save, UI action, or approved agent tool call. Access files never mutate: Access executes compiled p-code, so source writes there could not take effect, and XLIDE opens them read-only. |
| VBA execution | run macro at cursor, run unit tests | Requires an explicit run action and Office COM (Excel for macros; Excel, Word, or PowerPoint for tests). |
| Local support/export artifacts | export support bundle, export modules, write test artifacts | Writes files to a selected or configured local folder. |

XLIDE should not infer consent from merely opening a file or viewing a
module. Mutating actions are tied to visible commands, normal editor save
operations, GUIs, or agent tool calls that are expected to request user
approval before writes.

## File Mutation

Common file-mutating actions include:

- Saving an edited `xlide-vba` document.
- Adding, renaming, or deleting a VBA module.
- Importing modules from a folder.
- Writing cells through XLIDE tools.
- Creating a new macro-enabled file.
- Installing or updating the `XlideAssert.bas` support module for tests.

Before applying bulk import/export changes, prefer the diff preview. It shows
which modules will change, be skipped, or be removed before you apply the
operation.

For high-value files, keep your own backup or source-control checkpoint
before large imports, module deletion, or structural changes. XLIDE
records in-session command outcomes and change summaries for support, but a
backup is still the cleanest recovery path after a bad edit.

## Macro And Test Execution

Macro and test execution require Office COM and are Windows-only. Macros run
through Excel; unit tests run through the file's own application - Excel,
Word, or PowerPoint.

Macro execution can affect the file, the file system, external data sources,
and any systems the macro touches. Run macros only when you trust the file's
code.

Tests run against a temporary copy in an XLIDE-owned application instance and
close without saving the original file. Ownership is proven before the host
touches any application setting, so a PowerPoint you already have open is
refused rather than borrowed. Installing the assertion support module does
modify the original file because it adds or updates `XlideAssert.bas`.

The production test runner does not require Excel's "Trust access to the VBA
project object model" setting for normal test execution. XLIDE's internal
developer VBE oracle is different and may require that setting because it creates
disposable VBE modules to verify language behavior.

## Trust Center And Macro Security

Office macro security still applies. If the application blocks macros for a
file, XLIDE cannot make that file's VBA run safely around the application's
policy.

Check these settings when macro or test execution is blocked:

- The file is in a trusted location or is unblocked according to your
  organization's macro policy.
- Macros are allowed for the file.
- Protected View is not blocking execution.
- The file is not waiting on a modal prompt.
- The application's COM is registered on the machine.

Only enable "Trust access to the VBA project object model" when a workflow
explicitly requires VBE automation. Normal XLIDE editing uses native module
I/O and normal test execution should not require that setting.

## File Open State

Some workflows can behave differently when the file is already open in its
application. XLIDE has a setting, `xlide.attachToRunningExcel`, that controls
whether Windows Excel workflows try to reuse an already-running Excel
instance.

If a write, run, or test operation appears blocked:

- Close other Office windows that have the same file open (XLIDE's lock
  messages name the application holding it).
- Save pending application changes before running XLIDE operations.
- Check whether the file is read-only or locked by another process.
- Retry from a clean state if the application was left at a prompt or in
  edit mode.

## Support Bundle

Use **Export XLIDE Support Bundle** when you need a local diagnostic snapshot.
The bundle is designed to help debug setup, COM availability, settings, recent
XLIDE command outcomes, and analysis/test/sync state without including VBA
source by default.

By default, XLIDE redacts file paths and path-like settings. Optional
sections may include selected logs or anonymized analysis summaries
when you explicitly opt in.

Use **Copy Diagnostics** for a smaller setup-focused summary when the issue is
basic environment health.

## Recovery Checklist

If something goes wrong:

1. Stop and save a copy of the file before retrying destructive actions.
2. Close Office instances that may still hold the file.
3. Reopen VS Code and refresh the XLIDE tree.
4. Use the import/export diff preview to inspect file-vs-export differences.
5. Use the support bundle if the failure involves setup, COM, test host, or
   unclear command outcomes.
6. Restore from your own workbook backup or source-controlled module export when
   the desired state is known.
