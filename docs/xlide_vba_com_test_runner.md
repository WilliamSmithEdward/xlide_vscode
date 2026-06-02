# XLIDE VBA COM Test Runner

Status: planned downstream-developer documentation, not implemented behavior.

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

1. Create VBA test procedures using an explicit XLIDE test marker or manifest.
2. Open the workbook/project in XLIDE.
3. Run `XLIDE: Run VBA Tests`.
4. Review test results in VS Code.
5. Open failing tests, assertion messages, runtime errors, and logs from the
   result view.

Discovery must be explicit. XLIDE should not discover tests by guessing from
procedure names alone.

## Planned Test Marker

Exact syntax is not final. The first implementation should prefer a
VBA-comment-compatible annotation such as:

```vba
' @xlide-test
Public Sub InvoiceTotal_AddsTax()
    XlideAssert.AreEqual 108, InvoiceTotal(100, 0.08)
End Sub
```

Feature-rich metadata should stay explicit and deterministic:

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

Rules to decide before implementation:

- Whether annotations live in comments, a manifest file, or both.
- The exact key/value grammar for metadata.
- Whether setup/teardown hooks live in annotations, a manifest, or convention
  backed by explicit manifest entries.
- Whether tests must be `Public Sub` procedures with no required parameters.
- How to select tests by module, procedure, tag, or workbook.

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
- Expected-failure markers require a reason.

Skip:

```vba
' @xlide-test-skip reason="Requires external workbook not checked into repo"
Public Sub ExternalWorkbookScenario()
End Sub
```

Semantics:

- Skipped tests are discovered but not executed.
- Skips require a reason.
- Skips appear in result JSON and the Test Results view.

Tags:

```vba
' @xlide-test tags=smoke,slow,requires-com
Public Sub SlowEndToEndScenario()
End Sub
```

Semantics:

- Tags are explicit metadata, not inferred from names.
- The command UI and automation mode should support include/exclude tag filters.

## Planned Execution Features

The shipped runner should support enough workflow surface for real projects:

- run all tests
- run the current test at cursor
- run all tests in the current module
- run selected tests by manifest id
- include or exclude tags
- rerun failed tests
- fail fast
- timeout per test and per suite
- setup and teardown hooks
- expected failure and skip accounting
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

## Planned Assertion Surface

The shipped documentation should describe every supported assertion with
examples. Candidate assertions:

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
