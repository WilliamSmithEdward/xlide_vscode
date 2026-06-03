# XLIDE VBA COM Test Runner

Status: first implementation slice exists. This document tracks the current
downstream-developer workflow plus remaining shipped-runner gaps.

Purpose: describe the intended workflow for writing and running VBA project tests
from XLIDE through Excel COM. This document should become the user-facing guide
before the test runner is considered shipped.

## Goals

- Let workbook developers author tests in VBA.
- Run those tests in the real Excel/VBA runtime through COM.
- Capture deterministic pass/fail results, runtime errors, assertion failures,
  explicit test output, timeouts, and teardown failures.
- Keep this product feature separate from XLIDE's internal Excel/VBE oracle.

## Intended Developer Workflow

1. Open the workbook/project in XLIDE.
2. Open the workbook-scoped Unit Tests GUI from the sidebar or workbook-tree
   context menu.
3. Install or update the bundled `XlideAssert.bas` support module from the
   Tests GUI. Test runs are blocked until that module is installed.
4. Create VBA test procedures using an explicit XLIDE test marker.
5. Run all tests, a filtered tag run, or rerun the last failed tests from the
   Tests GUI. Module/current-test controls should live in this GUI as the runner
   matures.
6. Review test results in VS Code.
7. Open failing tests, assertion messages, runtime errors, and logs from the
   result view.

Discovery must be explicit. XLIDE should not discover tests by guessing from
procedure names alone.

## Current Test Marker

XLIDE discovers tests from a VBA-comment-compatible annotation block immediately
above a zero-argument `Sub` in a standard module. Discovery is explicit; names
such as `TestSomething` do not matter unless the marker is present.

```vba
' @xlide-test
Public Sub InvoiceTotal_AddsTax()
    XlideAssert.AreEqual 108, InvoiceTotal(100, 0.08)
End Sub
```

Supported first-slice metadata is explicit and deterministic:

```vba
' @xlide-test tags=invoice,smoke owner=finance requirement=INV-104
Public Sub InvoiceTotal_AddsTax()
    XlideAssert.AreEqual 108, InvoiceTotal(100, 0.08)
End Sub

' @xlide-test tags=invoice,known-bug
' @xlide-test-xfail reason="Rounding fix pending in INV-231"
Public Sub InvoiceTotal_RoundsHalfCents()
    XlideAssert.AreEqual 108.01, InvoiceTotal(100.005, 0.08)
End Sub

' @xlide-test-skip reason="Requires staging ERP workbook"
Public Sub ImportsFromErp()
    XlideAssert.Fail "Should be skipped before execution"
End Sub
```

Current marker rules:

- Annotations live in apostrophe comment lines immediately above the procedure.
- Blank lines or non-comment source lines break the annotation block.
- `@xlide-test`, `@xlide-test-skip`, and `@xlide-test-xfail` all mark the
  procedure as discoverable.
- Tests must be standard-module `Sub` procedures with no parameters.
- Supported key/value metadata: `tags`, `owner`, `requirement`, `timeout`,
  `expected-error`, and `reason` on skip/xfail markers.
- Malformed test markers, malformed metadata, unsupported metadata keys,
  detached markers, non-standard-module markers, Functions, Properties, and
  parameterized Subs are reported through the `vba-test-directive` warning in
  live diagnostics and workbook analysis.
- Command execution supports run-all, current-module, current-test, tag
  include/exclude filters, and fail-fast mode.

## Expected Failures, Skips, and Tags

The runner should support common test-suite metadata without guessing.

Expected failure:

```vba
' @xlide-test
' @xlide-test-xfail reason="Known VBE rounding behavior under investigation"
Public Sub KnownFailingCase()
    XlideAssert.AreEqual 1, BrokenFunction()
End Sub
```

Semantics:

- If the test fails and is marked expected-failure, record `xfail`.
- If the test passes and is marked expected-failure, record `xpass`.
- `xpass` should be visible in the result view and configurable as either a
  suite failure or warning.
- Expected-failure markers should include a reason; if omitted, XLIDE records a
  default expected-failure reason.

Skip:

```vba
' @xlide-test-skip reason="Requires external workbook not checked into repo"
Public Sub ExternalWorkbookScenario()
End Sub
```

Semantics:

- Skipped tests are discovered but not executed.
- Skip markers should include a reason; if omitted, XLIDE records a default skip
  reason.
- Skips appear in result JSON and the Test Results view.

Tags:

```vba
' @xlide-test tags=smoke,slow,requires-com
Public Sub SlowEndToEndScenario()
End Sub
```

Semantics:

- Tags are explicit metadata, not inferred from names.
- The command UI supports include/exclude tag filters. Automation-oriented
  filtering remains a planned workflow surface.

## Current Execution Features

The current runner supports:

- run all discovered tests in the target workbook
- run all discovered tests that match include/exclude tag filters
- rerun failed, timed-out, host-error, and unexpected-pass tests from the last
  run using stable discovered test ids
- run all discovered tests in the current module from editor context
- run the discovered test at the cursor from editor context
- fail fast after the first failure or unexpected pass
- one dedicated XLIDE-owned Excel instance per test run
- read-only workbook open, disabled link updates, suppressed Excel alerts, and
  close-without-saving cleanup
- per-test timeout metadata enforced by the host runner watchdog
- bounded startup/open/run/cleanup stages so popups, hangs, and automation-busy
  states become timeout or host-error outcomes instead of indefinite runs
- pass/fail/skip/xfail/xpass accounting
- timeout and host-error accounting
- clean assertion details from the bundled `XlideAssert` support module, with
  terse fallback text for generic Excel automation `Run` failures
- JSON report logging in the XLIDE Output channel
- persisted default run artifacts under `tests` beside the workbook:
  `summary.json`, `host-trace.json`, `output.log`, and latest
  `status_for_ci.json`
- live/workbook diagnostics for invalid `@xlide-test` marker syntax and markers
  that cannot discover a runnable test
- a concise VS Code test results panel
- command palette and workbook-tree entry points
- Tests GUI runtime gating for Excel COM registration without launching Excel
- non-Windows discovery with skipped execution because Excel COM is Windows-only

## Locked Default Excel Host Contract

The default Excel host runs each selected test suite in one dedicated
XLIDE-owned Excel instance. It opens the target workbook read-only, runs all
selected tests in that same hidden instance, closes the workbook without saving,
and quits or kills only the Excel instance XLIDE created.

By default, the test runner does not attach to the user's normal Excel session.
Attaching to an existing Excel instance can be a future explicit opt-in mode,
but it is not the default safety posture.

The default host also accounts for Excel automation rough edges that can
otherwise hang the window or the extension:

- every test macro carries a positive timeout
- startup, workbook-open, macro-run, and cleanup stages are watchdog-bounded
- link updates are disabled when opening the workbook
- read-only recommendation prompts are bypassed
- Excel alerts that can block automation are suppressed
- known modal blockers should be prevented or handled with low-level hooks
  where possible
- unknown modals, Trust Center/trust-access failures, protected-view/security
  prompts, automation-busy states, and cleanup failures become `host-error` or
  timeout outcomes with host trace context
- normal cleanup closes without saving and quits the owned Excel instance
- timeout/hang cleanup kills the owned Excel instance and records the run as a
  host/runtime failure, not a normal assertion failure

The runner must not depend on `SendKeys`, foreground focus, or
timing-dependent keystroke playback for correctness. Excel/VBA automation should
use COM, Win32 handles/messages, UI Automation, or other inspectable low-level
hooks. If a blocker cannot yet be handled deterministically, the run should
finish as a bounded host failure with diagnostic metadata rather than waiting
forever.

`src/vbaTestHostOracle.ts` is the unit-test oracle surface for this contract,
and `src/vbaTestExcelHost.ts` owns the current PowerShell host script builder
and event parser. Together they validate simple lifecycle traces so COM-host
changes can prove the expected behavior without needing live Excel in routine
unit tests.

## Test Host Oracle Completion Plan

The test-host oracle should become the CI-safe proof surface for the Excel test
runner. It should validate deterministic traces and generated-host behavior
without launching live Excel in routine unit tests. Live Excel checks remain
optional smoke tests for environment-specific behavior, not the default oracle.

The completion plan:

1. Freeze the trace schema. Every important stage should have an inspectable
   event: host start, Excel creation with owned PID, workbook-open start/result,
   macro start/result, modal/blocker detection, timeout/watchdog decisions,
   cleanup start/result, workbook close, Excel quit, and owned-process kill.
   Events should carry deterministic metadata only; no workbook source text,
   absolute paths by default, or user workbook contents.
2. Expand contract invariants in `src/vbaTestHostOracle.ts`. The validator
   should reject traces that attach to user Excel, create multiple default
   Excel instances, omit read-only open flags, allow link/update/read-only
   prompts, save on close, run macros after close/kill, miss per-test timeout or
   suite timeout coverage, classify host failures as ordinary test failures, or
   depend on `SendKeys`/focus-driven keystroke automation.
3. Build a popup/blocker fixture matrix. Each blocker needs one synthetic trace
   that proves the expected classification: prevented, handled by low-level
   hook, timeout, or `host-error`. Include workbook-open prompts, link/update
   prompts, read-only-recommended prompts, password/protected-view/security
   prompts, Trust Center/trust-access failures, compile/runtime/modal VBA
   dialogs, add-in prompts, repair/recovery prompts, automation-busy states,
   unknown modal dialogs, macro hangs, startup failures, and cleanup failures.
4. Cover the full result taxonomy. Oracle fixtures should prove `passed`,
   `failed`, `skipped`, `xfail`, `xpass`, `timeout`, and `host-error`, including
   fail-fast behavior, expected-error behavior, zero discovered tests, and no
   further macro execution after timeout/kill.
5. Validate host-script generation. Unit tests should assert that the generated
   host creates one owned Excel instance, never calls `GetActiveObject` in the
   default mode, never uses `SendKeys`, opens read-only with link updates and
   alerts suppressed, emits the required trace events, and executes bounded
   cleanup paths.
6. Validate artifact and CI-status mapping. Synthetic result/trace fixtures
   should prove `summary.json`, `host-trace.json`, `output.log`, and
   `status_for_ci.json` classification, including failed-test identity,
   bounded/sanitized messages, relative artifact paths, and omission of
   nondeterministic line/column data.
7. Add performance gates after correctness gates. The default path should be
   fast because it opens one owned read-only Excel instance, compiles once per
   run, runs selected tests in that same host, emits lightweight trace events,
   and does not re-open/re-import workbooks per test. Record duration by stage
   and by test so slow startup, open, compile, macro execution, modal handling,
   and cleanup are visible. Set regression thresholds after a baseline is
   measured; never trade away read-only safety, cleanup guarantees, or blocker
   classification just to make a run faster.
8. Add optional live Excel canaries. Keep a small manually invoked smoke suite
   for one passing test, one assertion failure, one expected failure, one
   timeout/hang, and one trusted-access failure where possible. These canaries
   should confirm the host emits traces matching the unit-test oracle, but they
   should not be required for routine CI.

Feature-complete means every runner state transition and every user-facing
outcome has a unit oracle fixture, every known blocker has an explicit
classification, generated host scripts are checked for the safety contract,
performance is measured by stage with regression gates, and unknown blockers are
bounded host failures with owned-process cleanup rather than indefinite runs.

## Planned Execution Features

Remaining shipped-runner execution features include:

- GUI selected-test checkbox/list execution
- suite timeout
- comprehensive popup/blocker matrix with deterministic handling or host-error
  classification
- stage/test duration metrics and performance regression baselines
- setup and teardown hooks
- explicit test logs
- GUI controls for artifact output folder and retention policy
- explicit headless/automation runner mode
- stable result ids for editor navigation and CI artifacts

Setup/teardown design should be explicit and documented before implementation.
Candidate scopes:

- before each test
- after each test
- before all tests in a module
- after all tests in a module
- before workbook test session
- after workbook test session

## Current Assertion Surface

Use the workbook-scoped Unit Tests GUI to install/update the standard
`XlideAssert` module. The current API is:

```vba
XlideAssert.AreEqual expected, actual
XlideAssert.AreNotEqual unexpected, actual
XlideAssert.IsTrue condition
XlideAssert.IsFalse condition
XlideAssert.AreSame expectedObject, actualObject
XlideAssert.IsNothing value
XlideAssert.IsNotNothing value
XlideAssert.Throws expectedErrorNumber, "ModuleName.ProcedureName"
XlideAssert.DoesNotThrow "ModuleName.ProcedureName"
XlideAssert.Fail "message"
XlideAssert.AssertionErrorNumber()
```

Assertion failures are recorded inside the bundled support module instead of
being raised as normal VBA runtime errors. For each run, XLIDE copies the
workbook to a temporary file, injects an XLIDE-owned direct-call dispatcher into
that copy, and opens the copy in the owned Excel host. The dispatcher calls each
selected test directly under `On Error GoTo Caught`, so assertion failures and
ordinary VBA runtime errors can return as structured failed-test results before
Excel Automation can collapse them into generic `Run` HRESULTs. This temporary
runner is for XLIDE integration, not the documented authoring API, and it does
not require COM access to the VBA project object model.

## Planned Assertion Surface

The shipped documentation should describe every supported assertion with
examples. Candidate future assertions:

```vba
XlideAssert.AreEqual expected, actual
XlideAssert.AreNotEqual unexpected, actual
XlideAssert.IsTrue condition
XlideAssert.IsFalse condition
XlideAssert.IsEmpty value
XlideAssert.IsNotEmpty value
XlideAssert.IsNothing value
XlideAssert.IsNotNothing value
XlideAssert.Contains expectedSubstring, actualText
XlideAssert.Matches expectedPattern, actualText
XlideAssert.Throws expectedErrorNumber, "ProcedureName"
XlideAssert.DoesNotThrow "ProcedureName"
XlideAssert.Fail "message"
XlideTest.Log "message"
```

The assertion API must be deterministic and implemented by an auditable VBA
support module or equivalent controlled test runtime.

## Planned Result Data

Each test result should include:

- workbook path or temporary run identity
- module name
- procedure name
- outcome: pass, fail, skip, xfail, xpass, error, timeout
- tags and metadata
- skip or expected-failure reason
- assertion message
- runtime error number, source, and description
- compile failure details when available
- explicit `XlideTest.Log` output
- elapsed time
- teardown/cleanup status

Machine-readable JSON output should distinguish persisted artifacts from an
explicit headless/automation runner mode. `status_for_ci.json` is useful for
automation consumers, but it does not by itself make XLIDE a headless CI runner.

## Run Artifacts and CI Status

The runner persists run artifacts into a workbook-specific output folder. The
default folder is `tests`, resolved relative to the workbook directory.
Workbook-specific sidecar settings can override both the folder and retention
count:

```json
{
  "tests": {
    "artifactFolder": "tests",
    "artifactRetention": 20
  }
}
```

`artifactFolder` may be relative to the workbook directory or absolute.
`artifactRetention` keeps the newest matching XLIDE run directories for the same
workbook, including the current run, and defaults to `20`. Cleanup only targets
run directories that match XLIDE's generated `<workbook>_yyyy-mm-dd_hhmmss`
pattern and contain `summary.json`; unrelated folders in the output directory
are left alone.

Each run creates a timestamped run directory using
`yyyy-mm-dd_hhmmss`, preferably prefixed with a sanitized workbook stem:

```text
tests/
  live_test_2026-06-03_142233/
    summary.json
    host-trace.json
    output.log
  status_for_ci.json
```

`summary.json` remains the complete machine-readable report for the run.
`host-trace.json` records the Excel host lifecycle trace for timeout and cleanup
diagnostics. `output.log` is a human-readable transcript for local debugging.

`status_for_ci.json` is overwritten on every test run and points at the most
recent run. It is compact, stable, and deterministic enough for CI to
consume without parsing the full report:

```json
{
  "schemaVersion": 1,
  "status": "fail",
  "reason": "test-failures",
  "generatedAt": "2026-06-03T21:22:33.000Z",
  "runId": "live_test_2026-06-03_142233",
  "workbook": {
    "name": "live_test.xlsm"
  },
  "paths": {
    "runDirectory": "tests/live_test_2026-06-03_142233",
    "summary": "tests/live_test_2026-06-03_142233/summary.json",
    "hostTrace": "tests/live_test_2026-06-03_142233/host-trace.json",
    "outputLog": "tests/live_test_2026-06-03_142233/output.log"
  },
  "counts": {
    "total": 12,
    "passed": 9,
    "failed": 2,
    "timeout": 1,
    "hostError": 0,
    "skipped": 0,
    "xfail": 0,
    "xpass": 0
  },
  "failedTests": [
    {
      "id": "InvoiceTests.CalculatesTax",
      "module": "InvoiceTests",
      "procedure": "CalculatesTax",
      "status": "failed",
      "durationMs": 124,
      "message": "Expected 108 but got 107.99"
    }
  ],
  "durationMs": 18420
}
```

`status` should be one of `pass`, `fail`, or `error`. `reason` should be one of
`passed`, `test-failures`, `timeouts`, `host-errors`, `unexpected-pass`,
`no-tests`, or `runner-error`. No discovered tests should produce `error` with
`reason: "no-tests"` by default, because silently treating zero tests as green is
dangerous.

`fail` covers failed tests, unexpected passes, and timeouts. `error` is reserved
for host or runner failures. `failedTests` should include `failed`, `timeout`,
`host-error`, and `xpass` outcomes because all of those should fail CI. Messages
should be bounded, with line breaks and control characters replaced, so CI logs
stay readable.

Do not put source text, workbook contents, or misleading location metadata in
`status_for_ci.json`. Procedure declaration line/column can stay in
`summary.json`; exact assertion failure line/column should be added only if the
assertion/runtime layer can provide it deterministically.

Prefer relative artifact paths where possible, and include workbook name only by
default. Avoid absolute workbook paths in CI status unless an explicit setting
enables them. Retention is a bounded cleanup policy that keeps the last N
matching run directories with a default around 20. Use `status_for_ci.json` as
the latest-run pointer instead of filesystem symlinks or junctions.

## Excel COM Session Policy

The runner should:

- run in an opt-in command, never during normal analysis;
- use the locked single owned read-only Excel host by default;
- offer disposable workbook/session isolation later as a stronger safety mode;
- avoid silently mutating a workbook already open outside XLIDE's control;
- close/reset Excel after each run where practical;
- warn clearly when Excel cannot be reset or the workbook cannot be reopened;
- use timeouts as a fallback, not as the primary success/failure signal.

## Documentation Required Before Shipping

Before the feature is called shipped, the public-facing guide
`user_guides/testing.md` should exist and cover:

- quickstart
- complete annotation or manifest syntax
- complete assertion API reference
- examples for expected errors and expected output
- command palette usage
- result view walkthrough
- JSON schema
- COM/Excel trust and security prerequisites
- tags, skips, expected failures, rerun failed, and fail-fast examples
- setup/teardown examples
- troubleshooting
- limitations and host-version notes

Keep this `docs/` file as the internal engineering contract and roadmap notes.
Keep `user_guides/testing.md` as the publish-ready end-user guide for the
broader testing topic.

## Non-Goals

- The test runner is not XLIDE's internal VBE oracle.
- The test runner should not run on every analysis or compile.
- The runner should not infer test intent from names alone.
- The runner should not hide Excel trust-center or COM failures.
