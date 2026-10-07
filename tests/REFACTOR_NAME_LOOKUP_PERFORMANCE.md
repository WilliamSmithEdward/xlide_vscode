# Refactor identifier lookup on the typing path

Related: https://github.com/WilliamSmithEdward/xlide_vscode/issues/985

The actual extension-host CPU profile showed code actions running alongside
large-class typing. Inline Variable's identifier lookup used an unanchored
Unicode regular expression over the entire source before the caret. It spent
11.33 ms across the profiled typing sequence scanning earlier identifiers.

Identifiers cannot cross CR/LF boundaries. The shared refactor lookup now
bounds both regex inputs to the caret's physical line. It preserves the
existing Unicode identifier rules and String.slice offset normalization,
including negative, fractional, NaN and infinite offsets. No waits or debounce
delays are added. Inline Variable and Introduce Parameter share this helper.

## Reproduce

~~~powershell
npx vitest run tests/refactorNameLookupPerformance.test.ts
$env:XLIDE_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npx vitest run tests/refactorNameWorkbookPerformance.test.ts
npm run compile
npx vscode-test --code-version 1.139.1 --grep 'Actual large class latency'
~~~

Portable parity tests compare every caret boundary against the preceding lookup
with CR, LF and CRLF, accented/combining/non-Latin/supplementary-plane letters,
numbers, bracketed names, comments, strings and Unicode line separators. A
work-count regression proves that only caret-line slices reach the lookup for
a document with 40,000 surrounding lines. The read-only workbook probe
alternates old/new timing order over 21 batches of ten lookups and checks the
returned identifier. The editor harness uses a disposable workbook copy and
checks code actions immediately after an actual edit, including retry attempts
in elapsed time if automatic casing overtakes the edit.

## Observations and limits

On ROneCOne, the matched lookup median fell from 5.7877 ms to 0.00171 ms.
Maximum batch-average samples were 8.077 and 0.1793 ms. This measures the
identifier lookup only, not an entire refactoring or UI response.

The editor probe observed 36.58 ms warm code actions and 6.53 ms from fresh edit
to returned code actions. Completion samples in that run were 49–60 ms;
fresh-edit hover was 39.24 ms. These are returned command results, not menu,
tooltip or lightbulb painting, and not a matched end-to-end before/after test.
An existing Type lookback test rejected an editor edit during automatic casing
on the first harness run; an unchanged rerun passed all 18 checks. A later
rebased run hit the same race in Immediate completion. Typing loops now retry
only rejected edits, at most three times, using fresh caret positions without
waits. All attempts stay inside measured latency and are recorded for the
completion probes. All 18 checks passed after that harness change; code actions
with two edit attempts took 33.57 ms in that run. The component improvement does
not eliminate scheduling or editor-edit outliers.

Completion outliers, other semantic collectors and background diagnostic work
still require investigation under #985. Long individual physical lines retain
line-sized lookup work.
