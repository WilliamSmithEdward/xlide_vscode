import type { VbaTestCase } from './vbaTestRunner';
import { VBA_IDENTIFIER_NAME_RE } from './vbaSourceScan';
import { XLIDE_VBA_JSON_ESCAPE_FUNCTION_LINES } from './vbaTestSupportModule';

/**
 * VBA codegen for the temporary test dispatcher module written into the staged
 * workbook copy: a `RunTest(testId)` Select Case dispatcher that calls each
 * discovered test directly so VBA runtime errors are caught inside VBA and
 * reported as JSON. The shared JsonEscape helper lines come from
 * vbaTestSupportModule.ts so the VBA JSON template lives once.
 */

export const XLIDE_TEST_RUNNER_MODULE_NAME = 'XlideTestRuntime';

export function buildVbaTestDirectRunnerModule(
    tests: readonly VbaTestCase[],
    moduleName = XLIDE_TEST_RUNNER_MODULE_NAME,
): string {
    validateVbaTestDispatcherIdentifiers(tests, moduleName);
    const cases = tests.map((test) => [
        `        Case ${vbaStringLiteral(test.qualifiedName)}`,
        `            Call ${test.moduleName}.${test.procedureName}`,
    ].join('\n'));
    return [
        `Attribute VB_Name = "${moduleName.replace(/"/g, '')}"`,
        'Option Explicit',
        '',
        'Public Function RunTest(ByVal testId As String) As String',
        '    XlideAssert.ResetTestState',
        '    On Error GoTo Caught',
        '    Select Case testId',
        ...cases,
        '        Case Else',
        '            RunTest = FailureJson(5, "XLIDE.TestRunner", "Unknown XLIDE test: " & testId)',
        '            Exit Function',
        '    End Select',
        '    On Error GoTo 0',
        '    If Len(XlideAssert.LastFailureMessage()) > 0 Then',
        '        RunTest = FailureJson(XlideAssert.AssertionErrorNumber(), "XLIDE.Assert", XlideAssert.LastFailureMessage())',
        '    Else',
        '        RunTest = "{""outcome"":""passed"",""output"":" & XlideAssert.OutputJson() & "}"',
        '    End If',
        '    Exit Function',
        'Caught:',
        '    Dim actualNumber As Long',
        '    Dim actualSource As String',
        '    Dim actualDescription As String',
        '    actualNumber = Err.Number',
        '    actualSource = Err.Source',
        '    actualDescription = Err.Description',
        '    On Error GoTo 0',
        '    RunTest = FailureJson(actualNumber, actualSource, actualDescription)',
        'End Function',
        '',
        'Private Function FailureJson(ByVal number As Long, ByVal source As String, ByVal message As String) As String',
        '    FailureJson = "{""outcome"":""failed"",""number"":" & CStr(number) & ",""source"":""" & JsonEscape(source) & """,""message"":""" & JsonEscape(message) & """,""output"":" & XlideAssert.OutputJson() & "}"',
        'End Function',
        '',
        ...XLIDE_VBA_JSON_ESCAPE_FUNCTION_LINES,
        '',
    ].join('\r\n');
}

export function validateVbaTestDispatcherIdentifiers(
    tests: readonly VbaTestCase[],
    runnerModuleName = XLIDE_TEST_RUNNER_MODULE_NAME,
): void {
    if (!VBA_IDENTIFIER_NAME_RE.test(runnerModuleName)) {
        throw new Error(`XLIDE test runner module name is not a valid plain VBA identifier: ${runnerModuleName}`);
    }
    for (const test of tests) {
        if (!VBA_IDENTIFIER_NAME_RE.test(test.moduleName)) {
            throw new Error(`XLIDE test module name is not a valid plain VBA identifier: ${test.moduleName}`);
        }
        if (!VBA_IDENTIFIER_NAME_RE.test(test.procedureName)) {
            throw new Error(`XLIDE test procedure name is not a valid plain VBA identifier: ${test.procedureName}`);
        }
    }
}

function vbaStringLiteral(value: string): string {
    return `"${value.replace(/"/g, '""')}"`;
}
