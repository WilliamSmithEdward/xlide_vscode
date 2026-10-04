# Editor performance checks

`editorSurfacePerformance.test.ts` checks work counts and behavior for completion,
hover, local context reuse, and the loop synchronization typing listener. Timing
is opt-in so ordinary CI checks do not depend on the host's load.

Run the synthetic benchmark from the repository root in PowerShell:

```powershell
New-Item -ItemType Directory -Force out | Out-Null
$env:XLIDE_EDITOR_BENCHMARK_OUTPUT = 'out/editor-latency.json'
npx vitest run tests/editorSurfacePerformance.test.ts -t 'measures editor latency'
Remove-Item Env:XLIDE_EDITOR_BENCHMARK_OUTPUT
Get-Content out/editor-latency.json
```

The report records median CPU time over 21 samples for a 1,200-procedure module
(about 47 KB), with a warm-up before each measurement. It covers the full
completion provider after a source edit and on repeated requests, the hover
resolver over `ThisWorkbook.Sheets(1).Name`, type-reference collection, and the
loop synchronization listener on an ordinary edit. The VS Code API is mocked;
the document text getter joins a line array to model text materialization.
These are component measurements, not visible suggestion-menu or mouse-hover
latency. Run before and after changes on an otherwise idle machine; do not run
the full suite concurrently.
