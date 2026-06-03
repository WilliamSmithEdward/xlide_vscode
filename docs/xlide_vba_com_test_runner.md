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
2. Run `XLIDE: Install VBA Test Support Module` if you want the built-in
   `XlideAssert` helpers.
3. Create VBA test procedures using an explicit XLIDE test marker.
4. Run all tests, filtered tests, the current module's tests, or the test at
   the cursor.
5. Review test results in VS Code.
6. Open failing tests, assertion messages, runtime errors, and logs from the
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
- run all discovered tests in the current module from editor context
- run the discovered test at the cursor from editor context
- fail fast after the first failure or unexpected pass
- pass/fail/skip/xfail/xpass accounting
- JSON report logging in the XLIDE Output channel
- a concise VS Code test results panel
- command palette and workbook-tree entry points
- non-Windows discovery with skipped execution because Excel COM is Windows-only

## Planned Execution Features

Remaining shipped-runner execution features include:

- run selected tests by manifest id
- rerun failed tests
- timeout per test and per suite
- setup and teardown hooks
- explicit test logs
- JSON output for automation
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

Run `XLIDE: Install VBA Test Support Module` on a workbook to install/update the
standard `XlideAssert` module. The current API is:

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

Assertion failures raise `vbObjectError + 513` with source `XLIDE.Assert`, so
the COM runner records them as failed tests.

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

Machine-readable JSON output should be documented for automation.

## Excel COM Session Policy

The runner should:

- run in an opt-in command, never during normal analysis;
- use a disposable workbook/session by default;
- avoid silently mutating a workbook already open outside XLIDE's control;
- close/reset Excel after each run where practical;
- warn clearly when Excel cannot be reset or the workbook cannot be reopened;
- use timeouts as a fallback, not as the primary success/failure signal.

## Documentation Required Before Shipping

Before the feature is called shipped, this document should include:

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

## Non-Goals

- The test runner is not XLIDE's internal VBE oracle.
- The test runner should not run on every analysis or compile.
- The runner should not infer test intent from names alone.
- The runner should not hide Excel trust-center or COM failures.
