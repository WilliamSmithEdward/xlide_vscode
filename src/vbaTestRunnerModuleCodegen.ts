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
        // The parameters carry the canonical casing of the host members they
        // shadow (Err.Number/.Source, a user's Message): a lowercase spelling
        // here re-cases every one of them project-wide (issue #38). The JSON
        // KEYS in the literal are wire protocol and stay lowercase.
        'Private Function FailureJson(ByVal Number As Long, ByVal Source As String, ByVal Message As String) As String',
        '    FailureJson = "{""outcome"":""failed"",""number"":" & CStr(Number) & ",""source"":""" & JsonEscape(Source) & """,""message"":""" & JsonEscape(Message) & """,""output"":" & XlideAssert.OutputJson() & "}"',
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
 *
 * Matching runs under `Option Compare Text` with keys in their ORIGINAL
 * casing: both sides of the comparison then use VBA's own case-insensitive
 * rule. Lowercasing the keys here with JS toLowerCase() and the incoming
 * name with LCase$ at runtime mixed two languages' case mappings, which
 * disagree for names like Turkish dotted I - reachable now that module
 * names are full unicode.
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

    // Grouping for bare-name uniqueness approximates VBA's text comparison
    // with JS toLowerCase(); the qualified `Module.Proc` key is always
    // emitted and always authoritative, so a divergence here can only cost
    // a bare shortcut, never a dispatch.
    const bareCounts = new Map<string, number>();
    for (const target of targets) {
        const key = target.procedureName.toLowerCase();
        bareCounts.set(key, (bareCounts.get(key) ?? 0) + 1);
    }

    // VBA caps a compiled procedure at 64KB, so the Select Case is chunked
    // into bounded helpers; targets stay direct calls in one execution
    // context either way (the error handler in XlideInvokeTarget catches
    // through the whole direct-call chain).
    const CHUNK_SIZE = 100;
    const chunks: DispatchTarget[][] = [];
    for (let start = 0; start < targets.length; start += CHUNK_SIZE) {
        chunks.push(targets.slice(start, start + CHUNK_SIZE));
    }

    const helperLines: string[] = [];
    chunks.forEach((chunk, index) => {
        helperLines.push('');
        helperLines.push(`Private Function XlideDispatch${index}(ByVal targetKey As String) As Boolean`);
        helperLines.push(`    XlideDispatch${index} = True`);
        helperLines.push('    Select Case targetKey');
        for (const target of chunk) {
            const qualified = `${target.moduleName}.${target.procedureName}`;
            const keys = [vbaStringLiteral(qualified)];
            if (bareCounts.get(target.procedureName.toLowerCase()) === 1) {
                keys.push(vbaStringLiteral(target.procedureName));
            }
            helperLines.push(`        Case ${keys.join(', ')}`);
            helperLines.push(`            ${target.moduleName}.${target.procedureName}`);
        }
        helperLines.push('        Case Else');
        helperLines.push(`            XlideDispatch${index} = False`);
        helperLines.push('    End Select');
        helperLines.push('End Function');
    });

    const dispatchCalls = chunks.map(
        (_chunk, index) => `    If XlideDispatch${index}(targetKey) Then Exit Sub`,
    );

    return [
        `Attribute VB_Name = "${XLIDE_TEST_DISPATCH_MODULE_NAME}"`,
        'Option Explicit',
        // Case-insensitive Select Case matching by VBA's own comparison
        // rule, so target names never pass through a JS case mapping.
        'Option Compare Text',
        '',
        "' Generated for one XLIDE test run and staged into the temporary copy",
        "' next to XlideAssert; never written into the user's file.",
        // MacroName, not macroName: the lowercase spelling would re-case every
        // MacroName in the project the moment this module is staged (#38).
        'Public Sub XlideInvokeTarget(ByVal MacroName As String)',
        '    On Error GoTo Caught',
        '    XlideAssert.RecordTargetOutcome 0, "", ""',
        '    Dim targetKey As String',
        '    targetKey = Trim$(MacroName)',
        ...dispatchCalls,
        '    Err.Raise 5, "XLIDE.TestDispatch", "Unknown test target: " & MacroName',
        '    Exit Sub',
        'Caught:',
        '    XlideAssert.RecordTargetOutcome Err.Number, Err.Source, Err.Description',
        'End Sub',
        ...helperLines,
        '',
    ].join('\r\n');
}
