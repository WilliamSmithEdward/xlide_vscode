# Testing VBA Workbooks

XLIDE can discover and run workbook tests that you write in normal VBA modules.
Tests run in Excel through COM so they exercise the same runtime your workbook
uses.

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
- `reason="..."` on skip and expected-failure markers

XLIDE reports invalid marker syntax as a `vba-test-directive` warning in live
diagnostics and workbook analysis. This catches typos, unsupported metadata
keys, invalid timeouts, detached markers, markers in class/document modules,
Functions, Properties, and parameterized Subs.

## Run Tests

Use the XLIDE Unit Tests workbook action or command palette test commands to run
tests for the selected workbook. The runner supports all discovered tests,
current module/current test selection, tag include/exclude filters, and fail-fast
mode.

By default XLIDE creates one XLIDE-owned Excel instance, opens the workbook
read-only, runs the selected tests, closes without saving, and cleans up the
owned Excel process. It does not attach to your normal Excel session by default.

## Results And Artifacts

Results include passed, failed, skipped, expected-failure, unexpected-pass,
timeout, and host-error outcomes. Each run writes artifacts beside the workbook
under the default `tests` folder:

```text
tests/
  workbook_name_yyyy-mm-dd_hhmmss/
    summary.json
    host-trace.json
    output.log
  status_for_ci.json
```

`summary.json` is the complete report. `host-trace.json` captures the Excel host
lifecycle. `output.log` is a readable transcript. `status_for_ci.json` is
overwritten on each run with compact latest-run metadata for downstream CI.

## Current Limitations

Excel COM execution is Windows-only. The first shipped path uses the selected
workbook and standard-module VBA tests; setup/teardown helpers, richer expected
output/state assertions, configurable artifact folders, retention policies, and
a full test runner GUI are planned follow-ups.
