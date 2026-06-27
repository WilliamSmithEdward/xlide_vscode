import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import {
    buildCloseWorkbookScript,
    buildRefreshReadOnlyScript,
    forgetWorkbookOpenedByXlide,
    markWorkbookOpenedByXlide,
    resolveExcelCoordinationSettings,
    resolveReopenReadOnly,
    shouldAttemptClose,
    wasWorkbookOpenedByXlide,
    type ExcelCoordinationSettings,
} from '../src/excelWorkbookCoordinator';

const settings = (over: Partial<ExcelCoordinationSettings> = {}): ExcelCoordinationSettings => ({
    mode: 'block',
    trackOpenedWorkbooks: true,
    reopenAfterClose: true,
    reopenMode: 'lastState',
    reopenReadOnlyAfterSave: false,
    ...over,
});

describe('excelWorkbookCoordinator', () => {
    describe('buildCloseWorkbookScript', () => {
        const script = buildCloseWorkbookScript('C:\\work\\Book.xlsm', false);

        it('targets the workbook by full path and basename', () => {
            expect(script).toContain("$targetPath = 'C:\\work\\Book.xlsm'");
            expect(script).toContain("$targetName = 'Book.xlsm'");
        });

        it('attaches to the running Excel and closes without saving Excel\'s copy', () => {
            expect(script).toContain('GetActiveObject("Excel.Application")');
            expect(script).toContain('$wb.Close($false)');
        });

        it('reports the close, lock, and prior read-only state through the sentinel', () => {
            expect(script).toContain('XLIDE_CLOSE|closed=');
            expect(script).toContain('|locked=');
            expect(script).toContain('|wasReadOnly=');
            // captures the workbook's read-only state before closing it
            expect(script).toContain('$wasReadOnly = [bool]$wb.ReadOnly');
        });

        it('passes the force flag and only kills Excel under force', () => {
            expect(script).toContain('$force = $false');
            const forceScript = buildCloseWorkbookScript('C:\\work\\Book.xlsm', true);
            expect(forceScript).toContain('$force = $true');
            // The kill is runtime-gated by `if ($locked -and $force)`.
            expect(forceScript).toContain('if ($locked -and $force)');
            expect(forceScript).toContain('Stop-Process -Force');
        });

        it('escapes single quotes in the interpolated path', () => {
            const quoted = buildCloseWorkbookScript("C:\\work\\Bob's Book.xlsm", false);
            expect(quoted).toContain("$targetPath = 'C:\\work\\Bob''s Book.xlsm'");
            expect(quoted).toContain("$targetName = 'Bob''s Book.xlsm'");
        });
    });

    describe('XLIDE-opened workbook tracking', () => {
        const wb = 'C:\\track\\Book.xlsm';
        beforeEach(() => forgetWorkbookOpenedByXlide(wb));

        it('remembers and forgets, ignoring case and separator casing', () => {
            expect(wasWorkbookOpenedByXlide(wb)).toBe(false);
            markWorkbookOpenedByXlide(wb);
            expect(wasWorkbookOpenedByXlide('c:\\TRACK\\book.xlsm')).toBe(true);
            forgetWorkbookOpenedByXlide(wb);
            expect(wasWorkbookOpenedByXlide(wb)).toBe(false);
        });
    });

    describe('shouldAttemptClose', () => {
        const wb = 'C:\\policy\\Book.xlsm';
        beforeEach(() => forgetWorkbookOpenedByXlide(wb));

        it('never closes under block mode, even for an XLIDE-opened workbook', () => {
            markWorkbookOpenedByXlide(wb);
            expect(shouldAttemptClose(settings({ mode: 'block' }), wb)).toBe(false);
        });

        it('always closes under closeForce', () => {
            expect(shouldAttemptClose(settings({ mode: 'closeForce' }), wb)).toBe(true);
        });

        it('closeTracked closes only XLIDE-opened workbooks when tracking is on', () => {
            const tracked = settings({ mode: 'closeTracked', trackOpenedWorkbooks: true });
            expect(shouldAttemptClose(tracked, wb)).toBe(false);
            markWorkbookOpenedByXlide(wb);
            expect(shouldAttemptClose(tracked, wb)).toBe(true);
        });

        it('closeTracked closes any matching workbook when tracking is off', () => {
            const untracked = settings({ mode: 'closeTracked', trackOpenedWorkbooks: false });
            expect(shouldAttemptClose(untracked, wb)).toBe(true);
        });
    });

    describe('buildRefreshReadOnlyScript', () => {
        const script = buildRefreshReadOnlyScript('C:\\ro\\Book.xlsm');

        it('only refreshes a workbook that is actually open read-only', () => {
            expect(script).toContain('if ($wb.ReadOnly) {');
            // close + reopen read-only ($true) so a closed workbook is never opened
            expect(script).toContain('$wb.Close($false)');
            expect(script).toContain('$excel.Workbooks.Open($targetPath, 0, $true)');
            expect(script).toContain('XLIDE_REFRESH|refreshed=');
        });

        it('escapes single quotes in the path', () => {
            const quoted = buildRefreshReadOnlyScript("C:\\ro\\Bob's Book.xlsm");
            expect(quoted).toContain("$targetPath = 'C:\\ro\\Bob''s Book.xlsm'");
        });
    });

    describe('resolveReopenReadOnly', () => {
        it('forces read-only / read-write explicitly', () => {
            expect(resolveReopenReadOnly('readOnly', false)).toBe(true);
            expect(resolveReopenReadOnly('readWrite', true)).toBe(false);
        });

        it('restores the prior state under lastState', () => {
            expect(resolveReopenReadOnly('lastState', true)).toBe(true);
            expect(resolveReopenReadOnly('lastState', false)).toBe(false);
        });

        it('falls back to read-only when the prior state is unknown', () => {
            expect(resolveReopenReadOnly('lastState', undefined)).toBe(true);
        });
    });

    describe('resolveExcelCoordinationSettings', () => {
        it('resolves the safe defaults from an unset configuration', () => {
            expect(resolveExcelCoordinationSettings()).toEqual({
                mode: 'block',
                trackOpenedWorkbooks: true,
                reopenAfterClose: true,
                reopenMode: 'lastState',
                reopenReadOnlyAfterSave: false,
            });
        });
    });
});
