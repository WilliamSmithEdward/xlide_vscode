# Intermittent typing and Backspace stalls

The private ROneCOne class supplied for issue #964 contains 969,296 characters. The editor probe copies its workbook, appends a small procedure, and edits only that disposable copy. Original source and profiles are not committed.

## Reproduction and measurements

Use VS Code 1.139.1, set `XLIDE_PERF_WORKBOOK` to the original workbook path, run `npm run compile`, then run `npx vscode-test --code-version 1.139.1 --grep "Actual large class latency"`. `XLIDE_PERF_CPU_PROFILE=1` additionally records extension-host CPU profiles in the disposable workspace. Avoid competing harness or unit runs during measurements.

The suite measures five edits with completion, 48 keyboard typing/smart Backspace pairs at two locations with 0/30/250/350 ms idle cadences, eight `.cez` -> `.ce` recovery cycles with completion/hover checks, and a declaration edit with updated hover. It verifies each keyboard insertion before deleting and verifies restoration afterward. Explicit editor focus prevents a notification or another focused control from swallowing the `type` command while native deletion still targets the editor. The heartbeat starts after CPU profiler setup; setup alone otherwise inflated its maximum delay.

These are command roundtrip/provider timings and extension-host heartbeat delays. They do not measure menu or tooltip painting. Focus commands and idle setup occur outside the typing measurement. A failed insertion invalidates a run; do not reinterpret it as a dropped Backspace.

For differential parser validation, set `XLIDE_INCREMENTAL_PARSE_CORPUS` to a private exported source file and run `npx vitest run tests/incrementalModuleParse.test.ts`. That optional test checks the entire AST and diagnostics against a fresh parse for tail and middle-procedure edits, without logging source.

## Changes

- Reparse an eligible same-line procedure-body edit and immutably rebase the unchanged tail. Header, newline, directive and malformed-procedure edits retain main's safe prefix/full parsing. Only the most recent compatible cached source is probed, so a failed attempt does not scan the class against eight historical versions.
- Retain the line-bounded caret name lookup now on main. Update procedure ownership from cached headers/lead-in lines for same-line nonstructural edits, including Backspace, and refresh the current range even when its display name did not change.
- Use shared current-module symbols for a standalone identifier statement while project context warms. Preserve the module-name/documentation row; member, type, argument and continued-statement positions retain their complete local context. The partial list remains refreshable.
- Yield between semantic-token collector passes in large documents and discard requests invalidated by typing, close or cancellation before running the next pass.

Actual-class parser comparisons with warm lexing measured fresh parses at 31-41 ms and tail-body incremental parses at 1.2-6.4 ms. A middle edit, including rebasing the tail, took 24.2 ms. These are component measurements rather than an editor speedup claim. Completion comparisons verify unchanged rows, documentation and insertion data for local declarations, Unicode names, module names, host functions and continued statements.

## Results and remaining work

The final CPU-profiled editor run passed all four actual-workbook cases. All 48 typing/Backspace pairs restored the document; typical typing and Backspace were approximately 2 ms. One Backspace sample still took 100.4 ms and the heartbeat maximum was 92.1 ms. The eight member-recovery cycles returned Cells and a hover every time: completion after deletion took 9.1-29.5 ms, and warm hover took about 1.2-1.5 ms. Updating the declaration produced the new hover in 37.8 ms.

Unprofiled repeat runs before the last bounded fallback recorded valid Backspace outliers up to 266.7 ms. Two subsequent repeats recorded maxima of 117.6 and 77.2 ms. A third repeat failed the keyboard insertion precondition and was excluded. Startup, profiling and scheduling differ between runs, so these observations do not establish a controlled overall speedup or absence of future outliers.

After merging main at f006e788, three unprofiled runs passed all five actual-workbook cases (including main's code-action probe): 144 valid typing/Backspace pairs and 24 member-recovery/hover cycles. Typing medians were 2.26/3.06/5.17 ms; Backspace medians were 1.89/4.83/7.37 ms and maxima were 25.28/105.07/304.80 ms. Member completion after deletion reached 82.77 ms; the slowest member-recovery Backspace was 142.11 ms. Fresh declaration hover took 33.46/61.32/53.41 ms. The merged full suite passed 14,378 tests across 721 files, with 32 tests and seven files intentionally skipped; 25 editor integration cases and the private-corpus differential parser check passed too. The 304.80 ms outlier remains material.

The persistent worker-queue and stale editing failures were fixed separately in #921 and #996. Issues #964 and #985 remain open for the remaining intermittent latency. The current profile still includes index rebuilds, semantic work and allocations; the goal is not complete while significant outliers remain.
