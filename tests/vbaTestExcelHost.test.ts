import { describe, expect, it } from 'vitest';
import {
    buildOwnedReadOnlyExcelTestHostScript,
    parseVbaTestHostEventLine,
    vbaTestHostPlanItems,
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
        expect(script).toContain('$excel.Visible = $false');
        expect(script).toContain('$excel.ScreenUpdating = $false');
        expect(script).not.toContain('$excel.Visible = $true');
        expect(script).toContain('$excel.Workbooks.Open($targetPath, 0, $true');
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
        expect(script).toContain('BM_CLICK');
        expect(script).toContain('modal-detected');
        expect(script).toContain('modal-dismissed');
        expect(script).toContain('modal-blocked');
        expect(script).toContain('safeToDismiss');
        expect(script).toContain('[XlideTestModalWatcher]::Start');
        expect(script).toContain('[XlideTestModalWatcher]::Stop');
        expect(script).toContain('Emit-XlideTestHostEvent "host-phase"');
        expect(script).toContain('Emit-XlideHostPhase "excel-create" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "workbook-open" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "workbook-close" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "excel-quit" "passed"');
        expect(script).toContain('Emit-XlideHostPhase "com-release" "passed"');
        expect(script).toContain('durationMs = [int]$phaseSw.ElapsedMilliseconds');
        expect(script).not.toContain('SendKeys');
        expect(script).toContain('if ($failFast -and -not $expectedFailure) { break }');
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
        })}`)).toEqual({
            kind: 'macro-finished',
            excelId: 'xlide-1',
            qualifiedName: 'Tests.Pass',
            outcome: 'passed',
            durationMs: 12,
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
