# Testing VBA Projects

XLIDE can discover and run VBA tests that you write in normal modules.
Tests run through COM in the file's own application - Excel for workbook
formats, Word for documents, PowerPoint for presentations - so they exercise
the same runtime your file uses. Access files cannot run tests: Access
executes compiled p-code, so the staged runner module cannot execute there,
and XLIDE says so instead of failing obscurely.

## Write Tests

Create a standard module for tests, such as `InvoiceTests`, and add zero-argument
`Sub` procedures. Mark each test with an XLIDE comment directive immediately
above the procedure.

```vba
' @xlide-test tags=smoke,invoice owner=finance requirement=INV-104
Public Sub InvoiceTotal_AddsTax()
    XlideAssert.AreEqual 108, InvoiceTotal(100, 0.08)
End Sub
```

Rules:

- The marker must be an apostrophe comment line immediately above the test.
- Blank lines or non-comment lines break the marker block.
- Tests must be in standard modules.
- Tests must be `Sub` procedures with no parameters.
- Test names are not inferred; a procedure runs as a test only when marked.

## Marker Syntax

Supported markers:

- `@xlide-test`
- `@xlide-test-skip reason="Requires external workbook"`
- `@xlide-test-xfail reason="Known issue pending fix"`

Supported metadata:

- `tags=smoke,fast`
- `owner=finance`
- `requirement=INV-104` or `req=INV-104`
- `timeout=10s` or `timeout=2500ms`
- `expected-error=13`
- `expected-error` or `expected-error=any`
- `reason="..."` on skip and expected-failure markers

`expected-error=13` makes the test pass only when the test procedure raises VBA
error 13. If the procedure raises no error, or raises a different VBA error
number, XLIDE records the test as failed. `expected-error` and
`expected-error=any` make the test pass when the procedure raises any caught VBA
error, and fail when it completes normally. Prefer a specific error number when
the expected failure path has a deterministic `Err.Number`. Use
`XlideAssert.Throws` inside a test when you want a single assertion to verify
that another procedure raises a specific error.

XLIDE applies the same contract to analysis. In a valid `@xlide-test`
procedure, `expected-error` suppresses only analyzer diagnostics that represent
the intentional deterministic runtime error path. It does not suppress the
whole procedure: syntax errors, compile-equivalent diagnostics, invalid test
marker syntax, style diagnostics, and unrelated findings remain visible.

XLIDE reports invalid marker syntax as a `vba-test-directive` warning in live
diagnostics and workbook analysis. This catches typos, unsupported metadata
keys, invalid timeouts, detached markers, markers in class/document modules,
Functions, Properties, and parameterized Subs.

## Assertions And Output

Install `XlideAssert.bas` from the Tests GUI before running tests. The module is
also used for IntelliSense when it is present in the workbook.

Core assertions:

- `XlideAssert.AreEqual expected, actual[, message]`
- `XlideAssert.AreNotEqual expected, actual[, message]`
- `XlideAssert.IsTrue condition[, message]`
- `XlideAssert.IsFalse condition[, message]`
- `XlideAssert.Fail [message]`

Object and value-state assertions:

- `XlideAssert.AreSame expectedObject, actualObject[, message]`
- `XlideAssert.AreNotSame expectedObject, actualObject[, message]`
- `XlideAssert.IsNothing actualObject[, message]`
- `XlideAssert.IsNotNothing actualObject[, message]`
- `XlideAssert.IsNullValue actual[, message]`
- `XlideAssert.IsNotNullValue actual[, message]`
- `XlideAssert.IsEmptyValue actual[, message]`
- `XlideAssert.IsNotEmptyValue actual[, message]`

String and expected-output assertions use binary, case-sensitive comparison:

- `XlideAssert.Contains actual, expectedSubstring[, message]`
- `XlideAssert.DoesNotContain actual, unexpectedSubstring[, message]`
- `XlideAssert.StartsWith actual, expectedPrefix[, message]`
- `XlideAssert.EndsWith actual, expectedSuffix[, message]`

Expected-error helpers:

- `XlideAssert.Throws expectedErrorNumber, macroName[, message]`
- `XlideAssert.DoesNotThrow macroName[, message]`

Use `Throws` or `DoesNotThrow` when one call inside a larger test should raise,
or not raise, a deterministic error. Use marker metadata such as
`expected-error=13` when the whole test procedure is expected to fall through
with that VBA error.

`XlideAssert.WriteLine value` records deterministic per-test output. Output
appears in the results view Details column, in `summary.json`, and in
`output.log`.

```vba
' @xlide-test tags=smoke
Public Sub InvoiceTotal_WritesUsefulOutput()
    Dim total As Currency
    total = InvoiceTotal(100, 0.08)

    XlideAssert.WriteLine "total=" & CStr(total)
    XlideAssert.AreEqual 108, total
    XlideAssert.Contains "invoice-total:108", "108"
End Sub
```

## Setup And Test Data

XLIDE tests are ordinary VBA. Put setup, teardown, factories, and fixture data
in helper procedures or modules, then call them explicitly from the test. This
keeps order deterministic and avoids hidden naming conventions.

```vba
' @xlide-test tags=invoice
Public Sub InvoiceTotal_UsesFixture()
    On Error GoTo Cleanup

    ResetInvoiceFixture
    SeedInvoiceFixture 100, 0.08

    XlideAssert.AreEqual 108, ActiveInvoiceTotal()

Cleanup:
    ClearInvoiceFixture
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub
```

Dedicated setup/teardown hook directives are not part of the public contract
yet. Prefer explicit helper calls until hook ordering, failure reporting, and
cleanup semantics are fully specified.

## Run Tests

Use the XLIDE Unit Tests action or the file-tree Unit Tests context menu to
open the Tests GUI for the selected file. Install or update the bundled
`XlideAssert.bas` support module from that GUI first; run buttons stay
disabled until the installed module matches the version bundled with XLIDE.
The GUI also checks for Office COM registration before enabling runs. That
check does not launch any Office application.

Runs execute against a temporary copy of the file in an XLIDE-owned instance
of the file's own application: Excel for workbook formats, Word for
documents, PowerPoint for presentations. XLIDE injects a transient test
dispatcher into that copy so ordinary VBA runtime errors are reported as
failed tests without requiring "Trust access to the VBA project object model"
through COM - including in Word, which surfaces a run target's unhandled
error as a modal dialog rather than propagating it; the dispatcher runs
targets as direct calls so every host behaves alike. The original file is not
modified by the run.

The Tests GUI runs all discovered tests, checked tests from the discovered test
list, current-module/current-test scopes from the active workbook editor,
selected include/exclude tag filters from discovered tag checkboxes, fail-fast
mode, and reruns failed/timed-out/host-error/unexpected-pass tests from the
previous run. It refreshes with the XLIDE workbook tree, so installing/removing
modules or refreshing the tree updates the support gate without reopening the
panel.

Command palette and automation surfaces use the same runner:

- `XLIDE: Run VBA Tests` opens the Tests GUI for the selected file.
- `XLIDE: Run VBA Tests With Options` prompts for include/exclude tag filters.
- `XLIDE: Run VBA Tests In Current Module` runs marked tests in the active
  module.
- `XLIDE: Run VBA Test At Cursor` runs the marked test under the editor cursor.
- The `xlide_runVbaTests` AI-agent tool runs headless and writes the same
  artifacts as the GUI.

By default XLIDE creates one XLIDE-owned application instance, opens the file
read-only, runs the selected tests, closes without saving, and cleans up the
owned process. It never attaches to your running Office session: ownership is
proven before any setting is touched, and a single-instance host (PowerPoint)
that hands back your own running application is refused with its settings
untouched. PowerPoint cannot hide its application window, so its runs show an
empty application frame; the presentation itself opens windowless. Test
execution uses COM to run macros, but it does not require the "Trust access
to the VBA project object model" setting. Support-module install and
temporary test-host injection use XLIDE's native module I/O rather than COM
`VBProject` automation.

> Note: XLIDE's internal developer-only Excel/VBE oracle is different from the
> production unit-test runner. The internal oracle can require Excel's
> "Trust access to the VBA project object model" setting because it creates
> disposable VBE modules to verify language behavior. Running unit tests
> through XLIDE should not require users to enable that setting.

## Results And Artifacts

Results include passed, failed, skipped, expected-failure, unexpected-pass,
timeout, and host-error outcomes. Assertion failures from the bundled
`XlideAssert.bas` module show concise assertion details in the results view
instead of raw automation stack output. Each run writes artifacts beside
the file under the default `tests` folder:

```text
tests/
  file_name_yyyy-mm-dd_hhmmss/
    summary.json
    host-trace.json
    output.log
  status_for_ci.json
```

`summary.json` is the complete report. `host-trace.json` captures the host
lifecycle. `output.log` is a readable transcript. `status_for_ci.json` is
overwritten on each run with compact latest-run metadata for downstream CI.
`summary.json` includes every discovered test result, status, duration, failure
message, metadata, and `XlideAssert.WriteLine` output. `status_for_ci.json`
keeps only deterministic CI-oriented metadata: schema version, pass/fail/error
status, reason, run id, generated timestamp, file name, relative artifact
paths, counts, duration, host summary, and failed test identities. It omits line
and column numbers until exact failure locations are deterministic.

To change the artifact folder or retention policy for one file, edit its
sidecar `<file>.xlide_settings.json`:

```json
{
  "tests": {
    "artifactFolder": "tests",
    "artifactRetention": 20
  }
}
```

`artifactFolder` may be relative to the file's directory or absolute.
`artifactRetention` keeps the newest matching XLIDE run directories for that
file. Cleanup only targets generated run folders that contain `summary.json`;
unrelated folders in the output directory are left alone.

The `xlide_runVbaTests` AI agent tool uses the same artifact writer and per-file
test settings as the Tests GUI. When an agent run is not blocked by setup or
application availability, its JSON response includes an `artifacts` object with
the run directory, `status_for_ci.json` path, effective artifact settings, and
the latest CI status payload.

## Current Limitations

Office COM execution is Windows-only. Access files cannot host test runs
(compiled p-code, stated above). The public test contract currently runs
explicitly marked zero-argument `Sub` procedures in standard modules.
Hook-style setup/teardown directives, deterministic compile-error preflight,
and richer suite-level lifecycle controls remain planned hardening work.
