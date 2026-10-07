# Live hard-error diagnostics performance â€” 2026-10-04

Investigated the 969,296-character `ROneCOne` class in
`F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm`.
The original workbook was only read; the editor probe used a disposable copy.
This checkout already contained other analyzer and worker-client changes;
those were preserved. Measurements used Node v24.18.0 and VS Code 1.139.1.

## Why the squiggle waits

The edit-time local pass starts after a 90 ms debounce, but it runs the entire
module analyzer. Incremental analysis avoids statement/expression walks for
unchanged procedures, while eager module rules still rerun. A hard syntax
finding is published only when that complete result returns.

The full pass follows after 450 ms on the healthy worker path and often
requests the same source and context. Before this change, even an identical
follow-up reran eager module checks. Both requests use one worker queue, so
that redundant work also delayed subsequent edits. Trace duration includes
queue waiting and must not be interpreted as independent CPU cost per pass.

Current-line syntax suppression can reveal cached findings immediately on a
cursor move only when the cache's document version matches the current edit.
With a newer document version, the editor must await fresh analysis. The
in-host fallback additionally scales its debounce up to 2 seconds on large
modules; that cap is a debounce cap, not a total-latency guarantee.

## Changes

The worker now retains one completed exact snapshot per document. Reuse
requires identical source, module/context fingerprint and active incomplete
expression offset. A changed owning project seed invalidates its completed
results; closing/forgetting a document drops its entry. Results containing
recovered analysis failures are not cached. A repeated response carries the
new request ID. Existing procedure-level incremental reuse remains in place.

`checkUnusedDeclarations` now collects references only to declarations it
actually checks. Constants need mention existence, so their uses bypass
read/write classification. Member-name exclusion, named-argument exclusion,
shadowing and inactive-branch mention behavior remain covered by tests.

## Measurements and limits

The native probe preserves worker state across five sequences of valid body
edits, an unmatched-parenthesis edit, and an unchanged follow-up. Every broken
snapshot must produce `unbalanced-parens` at the edited statement. Sources
are identical in the original/current comparisons; only the two changed
implementation files are replaced with their saved versions for the original
bundle. The complete sequence's diagnostic digest matches:
`2c096600d947756ffb4750d1920cafaea804bb29a1695c8da067710316bb96c4`.

| Worker operation, five-round median | Original | Current |
|---|---:|---:|
| Valid edit | 1,478.83 ms | 1,446.28 ms |
| Hard syntax error | 1,278.65 ms | 1,286.91 ms |
| Identical follow-up | 1,194.92 ms | 0.02 ms |

These sequential runs do not establish a first-result speedup. They do
establish removal of the identical follow-up analysis: a regression test
asserts the analyzer runs only once for an exact repeated request. These
native times omit worker transport and editor publication.

The complete unused-declaration rule, after three warm-ups and across seven
rounds, measured median 36.11 ms original versus 17.97 ms current. Its findings
matched (digest `de5f9b3657dbaf264cccd466fce3038f808387796033d553eec7a88629e83501`).
This is a component measurement, not an end-to-end latency claim.

The freshly compiled extension's actual VS Code probe measured appearance
times of 1,639 / 1,434 / 1,634 ms and clearing times of 1,443 / 1,632 /
1,431 ms across three edits. Polling is every 200 ms, so these are approximate
publication waits. The largest 25 ms heartbeat gap was 109 ms. The remaining
first-squiggle delay is still the complete semantic pass before publication;
this change reduces repeat work and queue pressure but does not give syntax
errors a separate early publication path.

## Validation and reproduction

Compilation, type checking, focused cache/analyzer/scheduling/active-line
tests and the real editor probe passed. A broader worker-client test has an
existing contradictory timeout assertion: it expects `available === false`
after timeout, while this checkout's pre-existing implementation restarts
the worker. That unrelated implementation/test pair was left intact.

```powershell
node scripts/benchmark-live-diagnostics.mjs 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
node scripts/benchmark-live-diagnostics.mjs 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm' --baseline=35ec3574bb7501bb47aea29f7336f8683285b4ab
# Add --unused-only to isolate the complete unused-declaration rule.
# Add --profile to write live-diagnostics.cpuprofile in the repository.

$env:XLIDE_LIVE_DIAGNOSTICS_PERF_WORKBOOK = 'F:\GitHub\xlide\xlide_vscode_testing\ROneCOne_Delegates_Demo.xlsm'
npm run compile
.\node_modules\.bin\vscode-test.cmd --code-version 1.139.1 --grep 'Live diagnostics performance'
```

The opt-in integration test checks hard-error publication and removal with
the cursor moved off the edited line, and guards against a two-second
extension-host stall. It does not enforce a cross-machine squiggle deadline.

## Follow-up: exact MsgBox trailing-minus edit

The owner clarified the reproduction as `MsgBox "hello world"-`, followed by
Enter. The existing analyzer reports `string-arithmetic-coercion` for that
example; its publication previously waited behind the complete semantic pass.

The desktop editor now runs structural, lexical, expression-syntax and
string-arithmetic checks on an independent analysis worker. This uses the
same rule implementations, suppression directives, conditional activity and
project metadata as full analysis. Its own incremental state skips unchanged
procedure walks, and its mode participates in the reuse fingerprint so a
partial result can never satisfy a full analysis request. Ordinary project
analysis continues on the existing semantic worker. Loose files and the web
build retain their existing complete-analysis path.

Early results replace only their own diagnostic codes, retaining other
semantic findings. Current early results remain authoritative for those codes
when the full pass arrives, including syntax findings the full pass omitted
while an expression was being typed. They are cached unfiltered and held by
the editor's active-line filter, so cursor movement can reveal a completed
syntax finding without requesting another analysis. Stale document generations
are still rejected. Failed semantic-worker fallback keeps its large-module
backoff even when the early worker remains healthy.

The integration probe now inserts the exact MsgBox text, positions the caret
after the minus and executes the editor's real `type` command with a newline.
It observes diagnostic events instead of polling for measured appearance and
clearing. In one freshly built run, appearance took 688 / 382 / 389 ms;
clearing took 271 / 265 / 248 ms. Enter-command time was 9 / 5 / 18 ms, and the
largest host timer gap was 165 ms. Staying on an unmatched-parenthesis line
for four seconds, then moving off it, revealed the cached error in 85 ms.

An earlier run under heavier machine load recorded a 1,233 ms appearance and
failed the new one-second guard. The independent queue removes the dependency
on completing semantic analysis; it does not guarantee an absolute latency
under every load or eliminate the whole-module lexical/structural work. The
earlier first-result measurements above describe the superseded shared-worker
path. The probe also checks retention of existing semantic findings, one
trailing-minus finding, and a cached line-exit reveal below 500 ms.

Final verification passed compilation/type checking, 99 focused tests across
10 files (including the web bundle), and the exact editor probe. That run
measured appearance at 590 / 374 / 387 ms, clearing at 266 / 250 / 249 ms,
Enter-command time at 10 / 4 / 3 ms, cached syntax reveal at 61 ms and a
123 ms maximum host heartbeat gap. Existing semantic finding codes remained
present during early publication, and the trailing-minus finding was not
duplicated. These are local observations, not a released-version claim.


## All error-severity findings

The initial narrow syntax/string pass above is superseded by `errorsOnly`:
all error-producing rules now run in the independent early worker. Selection
and editor merging use effective severity, so downgraded runtime errors remain
warnings in the full pass. There is no diagnostic-code whitelist restricting
which red findings can arrive early. Desktop workbook virtual documents use
this early worker; other surfaces retain their existing execution routes.

Twenty-five eager rule families now accept a procedure filter while retaining
the complete module AST and symbols for facts such as function results, types,
array return shapes and module declarations. Their procedure findings share
the existing incremental cache and shift with untouched procedures. Changes to
module-variable writes, directives, envelopes or external context invalidate
reuse. Changed callees also invalidate their direct and indirect callers.
Literal-only output statements are excluded from the dependency fingerprint
because they cannot write a return value or ByRef argument; their own diagnostics
still rerun. The early/full modes are isolated in the incremental state itself.

Runtime dead-code suppression now builds flow facts only for procedures with
runtime findings. A stale straight-line snapshot is discarded before formatting
its large constant/default map. The full pass yields to a pending early pass,
then waits another 450 ms of quiet before dispatch, reducing competition with
follow-up edits prompted by the red finding. This can defer warning refresh.
Internal failure/settings diagnostics remain visible when merging early errors.

The actual editor probe copies the original workbook, uses the real Enter
command for the exact MsgBox repro, and inserts/clears eight semantic cases in
an isolated private procedure in the same 26,000-line ROneCOne class. Semantic
cases use editor edits and line exit; they are not all real Enter keystrokes.
The probe warms initial analysis before measuring edits, checks warning
retention and single findings, and still exercises cached syntax line exit.

A successful built run on 2026-10-04 recorded:

| Finding | Appearance (ms) | Clearing (ms) |
| --- | ---: | ---: |
| Exact MsgBox trailing minus, three rounds | 669 / 433 / 497 | 449 / 407 / 501 |
| Undeclared variable | 653 | 451 |
| Unknown call | 622 | 436 |
| Argument count | 668 | 422 |
| Missing member | 545 | 418 |
| Assignment type mismatch | 558 | 403 |
| Array subscript bounds | 574 | 388 |
| Object variable not set | 604 | 404 |
| Division by zero | 557 | 403 |

Enter itself took 10 / 7 / 7 ms. Cached syntax reveal took 62 ms, and the
largest host heartbeat gap was 152 ms. Initial whole-module analysis and adding
an entire procedure still took seconds. Edits to widely used function results
or module state can require many callers to be checked again. These are local
observations, not a release or a universal subsecond guarantee.

Earlier expanded runs failed the one-second guard: one had type/array findings
at 1,071 / 1,085 ms before restricting runtime dead-code flow scans; an earlier
MsgBox run fell back to full analysis because the dependency classifier omitted
Call nodes. Both defects were fixed before the successful run above. The original
workbook is only read/copied throughout this investigation.


Final verification: compilation/type checking passed; the diagnostic regression
selection passed 4,598 tests across 316 files. The focused final cache checks
also passed, including shadowed MsgBox/Debug names and string-based callers.
A second final editor run passed all one-second appearance guards: exact
MsgBox 686 / 496 / 452 ms, the eight semantic findings 544–701 ms, clearing
404–466 ms, cached syntax reveal 85 ms, and maximum host timer gap 125 ms.
No change was committed, published or released as part of this investigation.


## All severity publication follow-up

The early editor pass now uses the complete analyzer, including warnings and
information. This supersedes the errors-only editor routing described above;
the errorsOnly analyzer option remains supported for callers/tests that need it.
No diagnostic-code whitelist decides what gets fast publication. Both workers
share the same suppression, severity override and project-context logic.

Baseline editor measurements for information findings were 1,723 ms (unused
variable), 1,572 ms (never read), and 1,601 ms (unreachable code). The early
snapshot now replaces all ordinary findings together, so removed warnings/info
clear with errors. It retains unfiltered syntax for cursor-only reveal and
includes settings findings. Successful complete early results avoid another
full-worker scan for the same scheduler generation and text version. Project
or configuration invalidations create a new generation and still recheck
unchanged text; failed early analysis leaves the full pass available.

Unused local declaration and unreachable-code rules now accept the dirty
procedure filter. Module-private references and shadowing still inspect the
whole module; private procedure reachability and doc-comment rules remain
module passes because their dependencies extend beyond body edits.

The extended ROneCOne harness passed twice. After duplicate-pass elimination,
the exact MsgBox/Enter appearances were 778 / 545 / 515 ms. Eight semantic
errors appeared in 648-716 ms. Five warning/information families appeared in
774-828 ms and cleared in 546-606 ms: unused variables, never-read variables,
unreachable code, impossible TypeOf and malformed formula strings. Cached
syntax reveal was 62 ms; maximum host heartbeat gap was 100 ms. The preceding
run measured warning/info appearances at 544-652 ms. These are measurements,
not guarantees for every diagnostic code or machine.

Cold analysis, adding a procedure, changing module declarations or changing
shared effects can still require full analysis lasting seconds. The harness
covers body edits and actual Enter for the original MsgBox repro. Broad rule
regressions cover other severities, suppression, dead branches, dependency
changes, severity overrides, loose files and web bundle compatibility; they do
not establish individual latency bounds for every rule or platform.

Compilation passed. The expanded diagnostic regression selection passed 4,555
tests in 312 files before two added regression cases. Both added cases passed:
full parity for information changes/module reference changes, and all-severity
early publication without a repeated scan while unchanged text rechecks after
a project invalidation. The original workbook remains read-only.

Final expanded regression rerun: 4,557 tests passed across 312 files; targeted whitespace checks passed.


## Continued hunt: full scans and shared overhead

Added a read-only --surface benchmark covering module declaration additions
and removals, whole procedure additions/removals, and edits to Parameter that
can change caller effects. The initial source measured 2,671 ms; additions
and callee-effect edits measured 1,915-2,313 ms; restoring the original source
measured 1,454-1,825 ms. These are analyzer-only measurements, not editor
publication times. All these scenarios conservatively selected full analysis.

A CPU profile found allocation/GC, source stripping, start-map serialization,
and whole-module name collection alongside eager diagnostic rules. Further
changes avoid character-array allocation for lines containing no quote,
apostrophe or Rem token; build dependency-name sets only when effects changed;
reuse those sets when procedure body text is unchanged; and retain private
procedure mentions only for actual candidates. No rule or severity is omitted.
The ordinary read-only benchmark retained the identical diagnostic digest
c8d918734d345fbd6b558fca2adbe62f5091922562ca03cd0bc6824df22871fe.

Compilation and 4,579 tests in 315 files passed, including corpus comparison
of the old source stripping behavior and incremental/full diagnostic parity.
A latency run launched alongside that CPU-intensive regression suite failed
its timing bounds (including 915 ms cached reveal). It is a load-contended
measurement; a separate editor run is required for comparable idle latency.
The generated CPU profile was removed after analysis.

The isolated editor rerun passed: exact MsgBox appearances 702 / 500 / 496 ms; eight semantic errors 656-732 ms; warning/information appearances 665-826 ms, clearing 500-667 ms; cached reveal 63 ms; maximum heartbeat gap 113 ms. Cold/module-envelope scans remain multi-second.


## Shared procedure starts

Procedure starts now share the module constant table through a read-only
layer, with local constants/defaults in a small overlay and parameters/other
shadowing declarations hidden from the module layer. Analyzer-owned immutable
starts explicitly opt into sharing the base on subsequent reaching-value
updates; ordinary mutable input maps still detach. Existing scopes remain
independent, including defaults, unknown local constant values and duplicate
name handling. A bounded-work regression verifies 1,000 constants are not
enumerated on repeated writes; the existing mutable-map detachment test remains.

Added --copied-starts and --summary benchmark switches for a controlled
comparison. --copied-starts restores the two per-procedure Map copies and
ordinary start detachment, keeping the remaining analyzer changes. Across
module declarations, new procedures and Parameter effect edits, both variants
produced the identical diagnostic digest:
a000b7d68aed4705e3aacc9c45d0fba53c26a5e95749126b987344d75fc16a03.

Shared-start medians were 2,003 ms (declaration addition), 1,443 ms (removal),
1,938 ms (procedure addition), 1,493 ms (removal), 2,035 ms (callee effect edit),
and 1,548 ms (removal). A repeated copied-start control measured 1,982 / 1,475 /
1,943 / 1,586 / 2,081 / 1,502 ms respectively. Initial analyses were 2,436 and
2,458 ms. An earlier copied-start run was substantially noisier/slower. The
allocation reduction is verified; the end-to-end timing differences are within
run-to-run variation, so this is not evidence that full scans are now fast.

Final shared-start verification: compilation and 4,575 tests across 315 files passed. The isolated editor harness passed: exact MsgBox 669 / 481 / 450 ms; semantic errors 619-728 ms; warning/info 656-717 ms; warning/info clear 452-467 ms; cached reveal 62 ms; maximum heartbeat gap 120 ms.
