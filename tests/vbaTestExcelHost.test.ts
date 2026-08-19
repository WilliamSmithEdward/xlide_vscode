import { describe, expect, it } from 'vitest';
import {
    buildOwnedReadOnlyExcelTestHostScript,
    vbaTestHostPlanItems,
} from '../src/vbaTestExcelHost';
import {
    parseVbaTestHostEventLine,
    XLIDE_TEST_HOST_EVENT_PREFIX,
} from '../src/vbaTestHostOracle';
import {
    buildVbaTestDirectRunnerModule,
    buildVbaTestDispatchModule,
    XLIDE_TEST_DISPATCH_MODULE_NAME,
    XLIDE_TEST_RUNNER_MODULE_NAME,
} from '../src/vbaTestRunnerModuleCodegen';
import { XLIDE_ASSERT_MODULE_SOURCE } from '../src/vbaTestSupportModule';
import type { VbaTestCase } from '../src/vbaTestRunner';

describe('VBA test Excel host script', () => {
    it('builds a single owned read-only Excel host script without attaching to user Excel', () => {
        const script = buildOwnedReadOnlyExcelTestHostScript('C:/work/Book.xlsm', [
            { qualifiedName: 'Tests.Pass', timeoutMs: 5000, expectedFailure: false },
            { qualifiedName: 'Tests.KnownFailure', timeoutMs: 7000, expectedFailure: true },
        ], { failFast: true });

        expect(script).toContain("$hostProgId = 'Excel.Application'");
        expect(script).toContain('New-Object -ComObject $hostProgId');
        expect(script).not.toContain('GetActiveObject');
        expect(script).toContain('$app.DisplayAlerts = $false');
        expect(script).toContain('Set-XlideHostAlertsOff $excel');
        expect(script).toContain('$excel.AskToUpdateLinks = $false');
        expect(script).toContain('$excel.Visible = $false');
        expect(script).toContain('$excel.ScreenUpdating = $false');
        expect(script).not.toContain('$excel.Visible = $true');
        expect(script).toContain('$excel.Workbooks.Open($targetPath, 0, $true');
        expect(script).toContain(`$runnerModuleName = '${XLIDE_TEST_RUNNER_MODULE_NAME}'`);
        expect(script).toContain('$testRunnerRef = "\'" + ($workbook.Name -replace "\'", "\'\'") + "\'!" + $runnerModuleName + ".RunTest"');
        expect(script).toContain('$excel.Run($testRunnerRef, $macroName)');
        expect(script).not.toContain('$excel.Run($macroRef)');
        expect(script).not.toContain('VBProject');
        expect(script).not.toContain('VBComponents');
        expect(script).not.toContain('$excel.VBE');
        expect(script).not.toContain('Trust access');
        expect(script).toContain('[Type]::Missing, [Type]::Missing, [Type]::Missing, $true');
        expect(script).toContain('$workbook.Close($false)');
        expect(script).toContain('$excel.Quit()');
        expect(script).toContain('ReleaseComObject($workbook)');
        expect(script).toContain('ReleaseComObject($excel)');
        expect(script).toContain('[GC]::WaitForPendingFinalizers()');
        expect(script).toContain('exit 0');
        expect(script).toContain('GetWindowThreadProcessId');
        expect(script).toContain('visible = $false');
        expect(script).not.toContain('MainWindowHandle');
        expect(script).toContain('Emit-XlideTestHostEvent "macro-started"');
        expect(script).toContain('XlideTestModalWatcher');
        expect(script).toContain('EnumWindows');
        expect(script).toContain('GetDlgCtrlID');
        expect(script).toContain('BM_CLICK');
        expect(script).toContain('IDOK = 1');
        expect(script).toContain('IDYES = 6');
        expect(script).toContain('IDNO = 7');
        expect(script).toContain('compile-error-blocks-runner');
        expect(script).toContain('modal-detected');
        expect(script).toContain('modal-dismissed');
        expect(script).toContain('modal-blocked');
        expect(script).toContain('safeToDismiss');
        expect(script).toContain('buttonIds');
        expect(script).toContain('function Format-XlideHResult');
        expect(script).toContain('function Convert-XlideVbaRunResult');
        expect(script).toContain('function Format-XlideVbaRunResult');
        expect(script).toContain('function Convert-XlideNullableInt');
        expect(script).toContain('function Convert-XlideOutputLines');
        expect(script).toContain('function Format-XlideRunException');
        expect(script).toContain('$hex = Format-XlideHResult $exception.HResult');
        expect(script).not.toContain('[uint32]$exception.HResult');
        expect(script).toContain('XlideTestRuntime');
        expect(script).toContain('HRESULT: ');
        // Friendly wording for known HRESULTs lives in vbaTestFailureMessages.ts only.
        expect(script).not.toContain('0x800A9C68');
        expect(script).not.toContain('0x800706BE');
        expect(script).not.toContain('0x800706BA');
        expect(script).not.toContain('Excel could not run the test macro');
        expect(script).not.toContain('Excel automation became unavailable while running the test');
        expect(script).not.toContain('InvocationInfo.PositionMessage');
        expect(script).toContain('$message = "RUN_FAILED|" + (Format-XlideVbaRunResult $vbaRunResult)');
        expect(script).toContain('$errorNumber = Convert-XlideNullableInt $vbaRunResult.number');
        expect(script).toContain('$payload["errorNumber"] = $errorNumber');
        expect(script).toContain('$testOutput = Convert-XlideOutputLines $vbaRunResult.output');
        expect(script).toContain('$payload["output"] = @($testOutput)');
        expect(script).toContain('$message = "RUN_FAILED|" + (Format-XlideRunException $_)');
        expect(script).toContain('outcome = "runner-error"');
        // A per-test runner-error only aborts the whole run under fail-fast;
        // otherwise the loop continues so each remaining test still emits a result.
        expect(script).toContain('if ($failFast) { break }');
        expect(script).toContain('[XlideTestModalWatcher]::Start');
        expect(script).toContain('[XlideTestModalWatcher]::Stop');
        expect(script).toContain('Emit-XlideTestHostEvent "host-phase"');
        expect(script).toContain('Emit-XlideHostPhase "excel-create" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "workbook-open" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "workbook-open" "failed"');
        expect(script).toContain('OPEN_FAILED|XLIDE could not open the file read-only for tests');
        expect(script).toContain('Emit-XlideHostPhase "workbook-close" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "excel-quit" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "com-release" "passed"');
        expect(script).toContain('durationMs = [int]$phaseSw.ElapsedMilliseconds');
        expect(script).not.toContain('SendKeys');
        expect(script).toContain('if ($failFast -and -not $expectedFailure) { break }');
    });

    it('proves it owns its Excel instead of assuming it', () => {
        // The host quits its Excel on the way out, so attaching to a user's
        // instance would put their unsaved work in the blast radius. Snapshot
        // the running PIDs, and refuse if the new Application resolves to one.
        const script = buildOwnedReadOnlyExcelTestHostScript('C:/work/Book.xlsm', []);

        expect(script).toContain('$preExistingExcelPids');
        expect(script).toContain("$hostProcessName = 'EXCEL'");
        expect(script).toContain('Get-Process -Name $hostProcessName -ErrorAction SilentlyContinue');
        expect(script).toContain('$preExistingExcelPids -contains $excelPid');
        expect(script).toContain('XLIDE refused to run tests');
        expect(script).toContain('Emit-XlideHostPhase "excel-create" "failed"');
    });

    it('suppresses the prompts an owned instance cannot answer', () => {
        const script = buildOwnedReadOnlyExcelTestHostScript('C:/work/Book.xlsm', []);

        // msoAutomationSecurityLow: the macro-security prompt is Excel-owned and
        // carries no Win32 buttons, so no watcher could dismiss it.
        expect(script).toContain('$excel.AutomationSecurity = 1');
        // Test VBA can leave DisplayAlerts = True behind, so teardown re-asserts
        // it before both Close and Quit rather than trusting the initial set.
        const suppressions = script.split('Set-XlideHostAlertsOff $excel').length - 1;
        expect(suppressions).toBeGreaterThanOrEqual(3);
    });

    it('ties the owned Excel lifetime to the host process', () => {
        const script = buildOwnedReadOnlyExcelTestHostScript('C:/work/Book.xlsm', []);

        // A kill-on-close job means a crashed or force-killed host cannot orphan
        // a hidden EXCEL.EXE still holding the workbook open.
        expect(script).toContain('CreateKillOnCloseJob');
        expect(script).toContain('AssignProcessToJobObject');
        expect(script).toContain('killOnClose = $jobActive');
    });

    it('watches for modals across open and teardown, not only macro execution', () => {
        const script = buildOwnedReadOnlyExcelTestHostScript('C:/work/Book.xlsm', [
            { qualifiedName: 'Tests.Pass', timeoutMs: 5000, expectedFailure: false },
        ]);

        // Opening a workbook and closing it can both prompt. An unwatched dialog
        // wedges the host until its timeout instead of being reported.
        expect(script).toContain('$excelId, "workbook-open"');
        expect(script).toContain('$excelId, "host-teardown"');
        // The watcher stops after COM teardown, not between tests.
        const stopIndex = script.lastIndexOf('[XlideTestModalWatcher]::Stop()');
        expect(stopIndex).toBeGreaterThan(script.indexOf('$excel.Quit()'));
    });

    it('builds a direct-call VBA dispatcher so runtime errors are caught inside VBA', () => {
        const source = buildVbaTestDirectRunnerModule([
            testCase('Tests.Pass', {}),
            testCase('MoreTests.RaisesRuntimeError', {}),
        ]);

        expect(source).toContain('Attribute VB_Name = "XlideTestRuntime"');
        expect(source).toContain('Public Function RunTest(ByVal testId As String) As String');
        expect(source).toContain('XlideAssert.ResetTestState');
        expect(source).toContain('On Error GoTo Caught');
        expect(source).toContain('Select Case testId');
        expect(source).toContain('Case "Tests.Pass"');
        expect(source).toContain('Call Tests.Pass');
        expect(source).toContain('Case "MoreTests.RaisesRuntimeError"');
        expect(source).toContain('Call MoreTests.RaisesRuntimeError');
        expect(source).toContain('RunTest = FailureJson(actualNumber, actualSource, actualDescription)');
        expect(source).toContain('XlideAssert.LastFailureMessage()');
        expect(source).toContain('XlideAssert.OutputJson()');
        expect(source).not.toContain('Application.Run');
    });

    it('rejects unsafe dispatcher identifiers before generating VBA source', () => {
        expect(() => buildVbaTestDirectRunnerModule([
            { ...testCase('Tests.Pass', {}), moduleName: 'Tests;Kill' },
        ])).toThrow(/module name is not a valid plain VBA identifier/);
        expect(() => buildVbaTestDirectRunnerModule([
            { ...testCase('Tests.Pass', {}), procedureName: 'Pass;Kill' },
        ])).toThrow(/procedure name is not a valid plain VBA identifier/);
        expect(() => buildVbaTestDirectRunnerModule([
            testCase('Tests.Pass', {}),
        ], 'Bad-Runner')).toThrow(/runner module name is not a valid plain VBA identifier/);
    });

    it('uses per-test timeout metadata when building host plan items', () => {
        const tests: VbaTestCase[] = [
            testCase('Tests.DefaultTimeout', {}),
            testCase('Tests.CustomTimeout', { timeoutMs: 2500, xfailReason: 'Known issue' }),
        ];

        expect(vbaTestHostPlanItems(tests)).toEqual([
            { qualifiedName: 'Tests.DefaultTimeout', timeoutMs: 30000, expectedFailure: false },
            { qualifiedName: 'Tests.CustomTimeout', timeoutMs: 2500, expectedFailure: true },
        ]);
    });

    it('parameterizes the host per Office application, semantics measured live', () => {
        const word = buildOwnedReadOnlyExcelTestHostScript('C:/work/Doc.docm', [], { hostApp: 'word' });
        expect(word).toContain("$hostKind = 'word'");
        expect(word).toContain("$hostProgId = 'Word.Application'");
        expect(word).toContain("$hostProcessName = 'WINWORD'");
        // Word: Documents.Open, Module.Proc run refs, ByRef [ref] argument,
        // wdDoNotSaveChanges close; no window handle, so ownership resolves
        // through the process diff.
        expect(word).toContain('$excel.Documents.Open($targetPath, $false, $true, $false)');
        expect(word).toContain('$testRunnerRef = $runnerModuleName + ".RunTest"');
        expect(word).toContain('$excel.Run($testRunnerRef, [ref]$macroArg)');
        expect(word).toContain('$workbook.Close(0)');

        const powerpoint = buildOwnedReadOnlyExcelTestHostScript('C:/work/Deck.pptm', [], { hostApp: 'powerpoint' });
        expect(powerpoint).toContain("$hostProgId = 'PowerPoint.Application'");
        expect(powerpoint).toContain("$hostProcessName = 'POWERPNT'");
        // PowerPoint: cannot hide, opens windowless read-only, runs through
        // reflection (its ParamArray rejects ByRef marshaling), and uses
        // file-qualified run refs.
        expect(powerpoint).toContain('if ($hostKind -ne "powerpoint") { try { $excel.Visible = $false } catch { } }');
        expect(powerpoint).toContain('$excel.Presentations.Open($targetPath, -1, 0, 0)');
        expect(powerpoint).toContain('$testRunnerRef = $workbook.Name + "!" + $runnerModuleName + ".RunTest"');
        expect(powerpoint).toContain('InvokeMember("Run"');

        // The default stays Excel, unchanged.
        const excel = buildOwnedReadOnlyExcelTestHostScript('C:/work/Book.xlsm', []);
        expect(excel).toContain("$hostKind = 'excel'");
    });

    it('parses host event lines and ignores ordinary output', () => {
        expect(parseVbaTestHostEventLine('not an event')).toBeUndefined();
        expect(parseVbaTestHostEventLine(`${XLIDE_TEST_HOST_EVENT_PREFIX}${JSON.stringify({
            kind: 'macro-finished',
            excelId: 'xlide-1',
            qualifiedName: 'Tests.Pass',
            outcome: 'passed',
            durationMs: 12,
            errorNumber: 13,
            output: ['hello'],
        })}`)).toEqual({
            kind: 'macro-finished',
            excelId: 'xlide-1',
            qualifiedName: 'Tests.Pass',
            outcome: 'passed',
            durationMs: 12,
            errorNumber: 13,
            output: ['hello'],
        });
    });
});

function testCase(qualifiedName: string, metadata: Partial<VbaTestCase['metadata']>): VbaTestCase {
    const [moduleName, procedureName] = qualifiedName.split('.');
    return {
        id: qualifiedName,
        moduleName,
        moduleType: 'standard',
        procedureName,
        qualifiedName,
        line: 1,
        column: 1,
        annotationLine: 1,
        metadata: {
            tags: [],
            ...metadata,
        },
    };
}

describe('the staged Throws dispatcher', () => {
    const modules = [
        {
            name: 'Tests',
            type: 'standard',
            source: [
                'Public Sub PlainTarget()',
                'End Sub',
                '',
                'Private Sub Hidden()',
                'End Sub',
                '',
                'Friend Sub AlsoHidden()',
                'End Sub',
                '',
                'Public Sub NeedsArg(ByVal value As Long)',
                'End Sub',
                '',
                'Public Function NotASub() As Long',
                'End Function',
                '',
                'Public Sub Shared()',
                'End Sub',
            ].join('\r\n'),
        },
        {
            name: 'MoreTests',
            type: 'standard',
            source: 'Public Sub Shared()\r\nEnd Sub\r\n',
        },
        {
            name: 'CHelper',
            type: 'class',
            source: 'Public Sub ClassTarget()\r\nEnd Sub\r\n',
        },
    ];

    it('direct-calls every public zero-parameter Sub in standard modules', () => {
        const source = buildVbaTestDispatchModule(modules);
        expect(source).toContain(`Attribute VB_Name = "${XLIDE_TEST_DISPATCH_MODULE_NAME}"`);
        expect(source).toContain('Case "tests.plaintarget", "plaintarget"');
        expect(source).toContain('            Tests.PlainTarget');
        // Errors come back as recorded state, never across a Run boundary.
        expect(source).toContain('XlideAssert.RecordTargetOutcome 0, "", ""');
        expect(source).toContain('XlideAssert.RecordTargetOutcome Err.Number, Err.Source, Err.Description');
        expect(source).toContain('Err.Raise 5, "XLIDE.TestDispatch", "Unknown test target: " & macroName');
        expect(source).not.toContain('Application.Run');
    });

    it('excludes what a direct call could not compile or should not reach', () => {
        const source = buildVbaTestDispatchModule(modules);
        expect(source).not.toContain('Hidden');
        expect(source).not.toContain('NeedsArg');
        expect(source).not.toContain('NotASub');
        expect(source).not.toContain('ClassTarget');
    });

    it('keeps ambiguous bare names qualified-only', () => {
        const source = buildVbaTestDispatchModule(modules);
        expect(source).toContain('Case "tests.shared"');
        expect(source).toContain('Case "moretests.shared"');
        expect(source).not.toContain('Case "tests.shared", "shared"');
        expect(source).not.toContain('Case "moretests.shared", "shared"');
    });

    it('chunks large projects under the 64KB compiled-procedure cap', () => {
        const bigModule = {
            name: 'Bulk',
            type: 'standard',
            source: Array.from({ length: 250 }, (_v, index) =>
                `Public Sub Target${index}()\r\nEnd Sub`).join('\r\n'),
        };
        const source = buildVbaTestDispatchModule([bigModule]);
        expect(source).toContain('Private Function XlideDispatch0(');
        expect(source).toContain('Private Function XlideDispatch1(');
        expect(source).toContain('Private Function XlideDispatch2(');
        expect(source).not.toContain('XlideDispatch3(');
        expect(source).toContain('If XlideDispatch2(targetKey) Then Exit Sub');
        expect(source).toContain('Bulk.Target249');
    });

    it('XlideAssert prefers the dispatcher and keeps the classic Run fallback', () => {
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Application.Run "XlideTestDispatch.XlideInvokeTarget", macroName');
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Public Sub RecordTargetOutcome');
        // Editing-time installs have no dispatcher; Excel propagates through
        // the classic path exactly as before.
        expect(XLIDE_ASSERT_MODULE_SOURCE).toContain('Application.Run macroName');
    });
});
