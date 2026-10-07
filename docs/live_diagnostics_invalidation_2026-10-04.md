# ROneCOne diagnostic invalidation performance

Investigated the 969,296-character, 1,552-procedure ROneCOne class in
`F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm`.
The original workbook was read only. The editor integration harness edited a
disposable copy, using the pinned VS Code 1.139.1 installation.

## Reproduction and changes

At the end of Function Value, the screenshot's `Debug wb As Worksheet` followed
by `Set wb = ThisWorkbook` produced three undeclared-variable errors. Repairing
`Debug` to `Dim` replaced them with an object assignment mismatch. Repairing
`Worksheet` to `Workbook` cleared the errors and restored the unused-read
information finding. Before these changes, each transition selected full
analysis and took approximately 2.0-2.4 seconds in the analyzer alone.

The dependency scan treated local variables, unrelated objects' members and
words in prose strings as module procedure references. Its repeated transitive
scan also allocated growing callee-name arrays. Binding now respects locals,
parameters, named argument labels and simple typed receivers; dynamic targets
and unknown receivers remain conservative. A reverse graph replaces repeated
whole-module dependency scans.

Direct callers are rechecked when callee effects change. Transitive invalidation
uses the argument-preservation and module-variable facts consumed by the rules.
Collection/dictionary member replay facts are compared too, including literal
output edits. Tests cover activation and clearing through ByRef chains, module
variable chains, dialogs and collection replay changes.

Parsed procedure interiors include their first/last statements and blank or
comment-only bodies. Stable procedure identities allow additions, removals and
reordering to retain unrelated findings. Ordinary signature and module variable
or constant edits recheck affected consumers. Header defaults and type names
remain dependencies even when a local or parameter shares their name.

Compiler directives, indirect declaration dependencies, implicit default-member
contracts and external context changes retain conservative full analysis.
Findings with structured offset data are rechecked instead of shifted blindly.
Module-wide facts and suppression are recomputed. All diagnostic severities
use the shared analysis and publication path.

## Full-analysis parity

The read-only benchmark's `--verify` option compares every timed snapshot with
a new worker seeded with the same project and request context. It compares
sorted public diagnostics and suppressed diagnostics, including structured
data. Diagnostic ordering may differ between full and incremental results.

Final verification passed 43 snapshot comparisons: 13 closer snapshots,
19 declaration/procedure/effect snapshots and 11 screenshot snapshots.

| Analyzer scenario | Median milliseconds |
| --- | ---: |
| Screenshot declaration typo | 396 |
| Screenshot declaration repair | 460 |
| Screenshot assignment repair | 411 |
| Mismatched closer | 386 |
| Missing closer | 454 |
| Closer repair | 351 |
| Module declaration addition / removal | 448 / 349 |
| Procedure addition / removal | 481 / 360 |
| Callee body effect edit / removal | 474 / 352 |

These are analyzer timings, excluding debounce and editor publication. Cold
initial analysis still measured approximately 2.4-2.5 seconds. Broad context
changes can still require comparable full scans. These results do not establish
an individual latency bound for every diagnostic rule, platform or machine load.

## Editor and regression verification

The expanded real-editor run passed the screenshot transitions in 499-639 ms;
the original trailing-minus appearances in 468-717 ms; eight semantic error
appearances in 543-683 ms and clears in 434-453 ms; and warning/information
appearances in 527-574 ms and clears in 420-451 ms. Procedure/declaration/closer
transitions took 436-691 ms. Cached line-exit reveal took 61 ms; the maximum
extension-host heartbeat gap was 142 ms.

Timing guards remain below 1,000 ms for measured appearance and clearing, and
below 500 ms for cached line-exit reveal. One loaded editor run measured
1,028 ms for its first trailing-minus edit; subsequent isolated runs passed.
Timing runs are serialized with regression runs.

Compilation passed. The final expanded regression selection passed 4,699 tests in
320 files, including parser, diagnostic, incremental parity, active-line,
cross-module, worker snapshot, scheduling, suppression and web bundle coverage.

## Superseded work and final editor measurements

An additional rapid-edit probe found a separate queue delay. An obsolete
snapshot correctly withheld its results but kept using the worker, delaying
the latest snapshot by 1,502 ms. Live requests now signal cancellation through
a shared atomic flag. The analyzer checks between phases, rules and procedures;
superseded work exits without publishing or caching partial findings. Closing
the document cancels its live work too. Explicit analysis requests and other
documents remain independent.

The final real-editor run with cancellation passed every timing guard:

| Editor transition | Milliseconds |
| --- | ---: |
| Trailing minus, appearance / clear | 496-732 / 450-483 |
| Screenshot typo and repairs | 501-655 |
| Eight semantic errors, appearance / clear | 619-771 / 450-591 |
| Five warning/information families, appearance / clear | 725-823 / 558-575 |
| Procedure/declaration/closer activation and deactivation | 547-838 |
| Latest snapshot after rapid edits / information clear | 761 / 637 |
| Cached syntax reveal on leaving the line | 62 |
| Maximum extension-host heartbeat gap | 117 |

The rapid-edit check also verifies that the superseded undeclared-variable
error never appears in the published results. Cooperative cancellation is
control flow, rather than a recovered analyzer failure. Worker tests verify
running and queued replacement, document closure, unrelated/explicit work,
cache integrity and timeout recovery in a fresh worker.

All 43 full-analysis snapshot comparisons were repeated with cancellation
checks enabled and passed. The three scenario diagnostic digests were unchanged.

With inactive cancellation checks enabled, the read-only warm screenshot
benchmark measured medians of 425 ms (typo), 473 ms (declaration repair) and
386 ms (assignment repair). Timing varies with machine load and warm-up.
Cold analysis in the later verification runs measured 2.5-2.9 seconds; broad
context changes still retain conservative full analysis.

Reproduce analyzer parity with:

```powershell
node scripts/benchmark-live-diagnostics.mjs F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm --closures --cancellable --verify --summary
node scripts/benchmark-live-diagnostics.mjs F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm --surface --cancellable --verify --summary
node scripts/benchmark-live-diagnostics.mjs F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm --screenshot --cancellable --verify --summary
```

Reproduce editor measurements with:

```powershell
npm run compile
$env:XLIDE_LIVE_DIAGNOSTICS_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
.\node_modules\.bin\vscode-test.cmd --code-version 1.139.1 --grep 'Live diagnostics performance'
```
