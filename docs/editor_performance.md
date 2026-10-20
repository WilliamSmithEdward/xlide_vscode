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

The broader surface benchmark covers document highlights in a 102-module
project, caret tracking during ordinary body edits in a 1,200-procedure module,
and repeated semantic-token requests for the same document version:

```powershell
New-Item -ItemType Directory -Force out | Out-Null
$env:XLIDE_SURFACE_BENCHMARK_OUTPUT = 'out/surface-latency.json'
npx vitest run tests/editorSurfaceFollowupPerformance.test.ts -t 'measures highlighting and caret surface work'
Remove-Item Env:XLIDE_SURFACE_BENCHMARK_OUTPUT
Get-Content out/surface-latency.json
```

It also uses 21 samples per median and a mocked VS Code API. Reference binding
still uses the whole project; the highlight occurrence search is limited to the
active document. Deterministic tests compare those results against filtering
the project-wide reference search and verify that structural edits invalidate
the caret range cache.

Exercise the registered providers and real editor events in the integration
harness (the workspace is a disposable copy of the fixture workbook):

```powershell
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Editor surfaces|Canonical casing|Diagnostics across modules|Project references|Doc comment diagnostics'
```

The editor suite checks dot completion on `ThisWorkbook.Sheets(1).`, hover
after an edit, signature help, read/write highlights, caret body edits and header
renames, and typing followed by completion/hover/semantic tokens in a
1,201-procedure module. It prints warm provider-command medians and edit-command
timings without machine-dependent pass thresholds. These include extension-host
command dispatch; they still exclude renderer repaint, mouse-hover delay, and
suggestion-widget presentation.
