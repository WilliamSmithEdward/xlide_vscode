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

Use the XLIDE Unit Tests workbook action or the workbook-tree Unit Tests context
menu to open the Tests GUI for the selected workbook. Install or update the
bundled `XlideAssert.bas` support module from that GUI first; run buttons stay
disabled until the installed module matches the version bundled with XLIDE.
The GUI also checks for Microsoft Excel COM registration before enabling runs.
That check does not launch Excel.

The Tests GUI currently runs all discovered tests or opens the filtered-run
prompt for tag include/exclude filters and fail-fast mode. It refreshes with the
XLIDE workbook tree, so installing/removing modules or refreshing the tree
updates the support gate without reopening the panel.

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

To change the artifact folder or retention policy for one workbook, edit the
workbook sidecar file `<workbook>.xlide_settings.json`:

```json
{
  "tests": {
    "artifactFolder": "tests",
    "artifactRetention": 20
  }
}
```

`artifactFolder` may be relative to the workbook directory or absolute.
`artifactRetention` keeps the newest matching XLIDE run directories for that
workbook. Cleanup only targets generated run folders that contain `summary.json`;
unrelated folders in the output directory are left alone.

## Current Limitations

Excel COM execution is Windows-only. The first shipped path uses the selected
workbook and standard-module VBA tests; setup/teardown helpers, richer expected
output/state assertions, and module/current-test run controls inside the Tests
GUI are planned follow-ups.
