# XLIDE Automation And CI Guide

XLIDE exposes file analysis and VBA test execution to AI agents through VS
Code language-model tools. Use these tools when an automated workflow needs to
inspect, edit, analyze, and test a macro-enabled Office file - Excel, Word,
PowerPoint, or Access - without driving the XLIDE panels by hand.

## Recommended Agent Flow

1. Discover the target file with `xlide_listWorkbooks` (it lists every macro
   container) or confirm structure with `xlide_getWorkbookInfo`.
2. Read the file's VBA with `xlide_readModule`.
3. Write changes with `xlide_writeModule` or the other module tools. Access
   files refuse writes with the reason (compiled p-code).
4. Run `xlide_analyzeWorkbook` and treat an empty `problems` array as analysis
   pass.
5. Run `xlide_runVbaTests` to execute discovered `@xlide-test` procedures
   through the production read-only test host of the file's own application
   (Excel, Word, or PowerPoint).

`xlide_runVbaTests` supports `moduleName`, `procedureName`, `testIds`,
`includeTags`, `excludeTags`, and `failFast` so an agent can start narrow while
editing and finish with a full run.

## Test Artifacts

Agent-driven test runs write the same artifact surface as the Tests GUI:

```text
tests/
  workbook_name_yyyy-mm-dd_hhmmss/
    summary.json
    host-trace.json
    output.log
  status_for_ci.json
```

The returned JSON includes `{ ok, summary, artifacts, report }`.
`artifacts.ciStatus` is the same payload written to `status_for_ci.json`.
Downstream CI should prefer `status_for_ci.json` for the latest run and
`summary.json` when it needs the full result detail.

If setup is incomplete, the tool returns `blocked: true` with `reason:
"test-support"` or `reason: "excel-com"` instead of attempting to run Excel.

## Analysis Contract

`xlide_analyzeWorkbook` uses the same analyzer that powers editor diagnostics.
Problems include module, line, column, severity, code, and message. Test
directives such as `expected-error` suppress only the intentionally expected
deterministic runtime diagnostic path inside that test procedure; unrelated
syntax, compile, style, or structural diagnostics remain visible.
