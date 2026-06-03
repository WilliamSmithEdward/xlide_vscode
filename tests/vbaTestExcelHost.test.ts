import { describe, expect, it } from 'vitest';
import {
    buildVbaTestDirectRunnerModule,
    buildOwnedReadOnlyExcelTestHostScript,
    parseVbaTestHostEventLine,
    vbaTestHostPlanItems,
    XLIDE_TEST_RUNNER_MODULE_NAME,
    XLIDE_TEST_HOST_EVENT_PREFIX,
} from '../src/vbaTestExcelHost';
import type { VbaTestCase } from '../src/vbaTestRunner';

describe('VBA test Excel host script', () => {
    it('builds a single owned read-only Excel host script without attaching to user Excel', () => {
        const script = buildOwnedReadOnlyExcelTestHostScript('C:/work/Book.xlsm', [
            { qualifiedName: 'Tests.Pass', timeoutMs: 5000, expectedFailure: false },
            { qualifiedName: 'Tests.KnownFailure', timeoutMs: 7000, expectedFailure: true },
        ], { failFast: true });

        expect(script).toContain('New-Object -ComObject Excel.Application');
        expect(script).not.toContain('GetActiveObject');
        expect(script).toContain('$excel.DisplayAlerts = $false');
        expect(script).toContain('$excel.AskToUpdateLinks = $false');
        expect(script).toContain('$excel.Visible = $false');
        expect(script).toContain('$excel.ScreenUpdating = $false');
        expect(script).not.toContain('$excel.Visible = $true');
        expect(script).toContain('$excel.Workbooks.Open($targetPath, 0, $true');
        expect(script).toContain(`$runnerModuleName = '${XLIDE_TEST_RUNNER_MODULE_NAME}'`);
        expect(script).toContain('$testRunnerRef = "\'" + $workbook.Name + "\'!" + $runnerModuleName + ".RunTest"');
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
        expect(script).toContain('0x800A9C68');
        expect(script).toContain('0x800706BE');
        expect(script).toContain('0x800706BA');
        expect(script).toContain('Excel could not run the test macro');
        expect(script).toContain('Excel automation became unavailable while running the test');
        expect(script).not.toContain('InvocationInfo.PositionMessage');
        expect(script).toContain('$message = "RUN_FAILED|" + (Format-XlideVbaRunResult $vbaRunResult)');
        expect(script).toContain('$errorNumber = Convert-XlideNullableInt $vbaRunResult.number');
        expect(script).toContain('$payload["errorNumber"] = $errorNumber');
        expect(script).toContain('$testOutput = Convert-XlideOutputLines $vbaRunResult.output');
        expect(script).toContain('$payload["output"] = @($testOutput)');
        expect(script).toContain('$message = "RUN_FAILED|" + (Format-XlideRunException $_)');
        expect(script).toContain('outcome = "runner-error"');
        expect(script).toContain('      break');
        expect(script).toContain('[XlideTestModalWatcher]::Start');
        expect(script).toContain('[XlideTestModalWatcher]::Stop');
        expect(script).toContain('Emit-XlideTestHostEvent "host-phase"');
        expect(script).toContain('Emit-XlideHostPhase "excel-create" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "workbook-open" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "workbook-open" "failed"');
        expect(script).toContain('OPEN_FAILED|XLIDE could not open the workbook read-only for tests');
        expect(script).toContain('Emit-XlideHostPhase "workbook-close" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "excel-quit" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "com-release" "passed"');
        expect(script).toContain('durationMs = [int]$phaseSw.ElapsedMilliseconds');
        expect(script).not.toContain('SendKeys');
        expect(script).toContain('if ($failFast -and -not $expectedFailure) { break }');
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
