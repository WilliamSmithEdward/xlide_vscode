import type { VbaTestCase } from './vbaTestRunner';
import { parseModule } from './analyzer';
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

export const XLIDE_TEST_DISPATCH_MODULE_NAME = 'XlideTestDispatch';

export interface VbaTestDispatchModuleInput {
    name: string;
    type?: string;
    source?: string;
}

/**
 * Generates the staged XlideTestDispatch module: a by-name dispatcher that
 * DIRECT-CALLS every public zero-parameter Sub in the project's standard
 * modules and records the outcome through XlideAssert.RecordTargetOutcome.
 *
 * XlideAssert.Throws/DoesNotThrow prefer this over `Application.Run target`
 * because Word surfaces a Run-target's unhandled error as a VBE modal instead
 * of propagating it to the calling VBA (measured live; Excel propagates). A
 * direct call inside the dispatcher keeps the error in one execution context,
 * so every host behaves alike. Targets resolve by `Module.Proc` always, and
 * by bare `Proc` when that name is unambiguous project-wide.
 */
export function buildVbaTestDispatchModule(modules: readonly VbaTestDispatchModuleInput[]): string {
    interface DispatchTarget {
        moduleName: string;
        procedureName: string;
    }
    const targets: DispatchTarget[] = [];
    for (const module of modules) {
        if ((module.type ?? 'standard') !== 'standard' || module.source === undefined) {
            continue;
        }
        if (
            module.name === XLIDE_TEST_DISPATCH_MODULE_NAME ||
            !VBA_IDENTIFIER_NAME_RE.test(module.name)
        ) {
            continue;
        }
        const parsed = parseModule(module.source);
        for (const member of parsed.members) {
            if (member.kind !== 'Procedure' || member.procKind !== 'Sub' || member.params.length > 0) {
                continue;
            }
            const modifiers = member.modifiers.map((modifier) => modifier.toLowerCase());
            if (modifiers.includes('private') || modifiers.includes('friend')) {
                continue;
            }
            if (!VBA_IDENTIFIER_NAME_RE.test(member.name)) {
                continue;
            }
            targets.push({ moduleName: module.name, procedureName: member.name });
        }
    }

    const bareCounts = new Map<string, number>();
    for (const target of targets) {
        const key = target.procedureName.toLowerCase();
        bareCounts.set(key, (bareCounts.get(key) ?? 0) + 1);
    }

    const cases: string[] = [];
    for (const target of targets) {
        const qualified = `${target.moduleName}.${target.procedureName}`.toLowerCase();
        const bare = target.procedureName.toLowerCase();
        const keys = [vbaStringLiteral(qualified)];
        if (bareCounts.get(bare) === 1) {
            keys.push(vbaStringLiteral(bare));
        }
        cases.push(`        Case ${keys.join(', ')}`);
        cases.push(`            ${target.moduleName}.${target.procedureName}`);
    }

    return [
        `Attribute VB_Name = "${XLIDE_TEST_DISPATCH_MODULE_NAME}"`,
        'Option Explicit',
        '',
        "' Generated for one XLIDE test run and staged into the temporary copy",
        "' next to XlideAssert; never written into the user's file.",
        'Public Sub XlideInvokeTarget(ByVal macroName As String)',
        '    On Error GoTo Caught',
        '    XlideAssert.RecordTargetOutcome 0, "", ""',
        '    Select Case LCase$(Trim$(macroName))',
        ...cases,
        '        Case Else',
        '            Err.Raise 5, "XLIDE.TestDispatch", "Unknown test target: " & macroName',
        '    End Select',
        '    Exit Sub',
        'Caught:',
        '    XlideAssert.RecordTargetOutcome Err.Number, Err.Source, Err.Description',
        'End Sub',
        '',
    ].join('\r\n');
}
