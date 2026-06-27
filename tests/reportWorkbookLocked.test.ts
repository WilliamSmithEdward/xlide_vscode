import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import * as vscode from 'vscode';
import { reportWorkbookLocked } from '../src/xlideFileSystem';

// reportWorkbookLocked must surface at most one "workbook is open in Excel"
// popup per workbook within a short window, so a burst of failed operations
// (or a writeFile failure followed by a re-read) never stacks notifications.
describe('reportWorkbookLocked throttling', () => {
    beforeEach(() => {
        vi.useFakeTimers();
        vi.setSystemTime(0);
        vi.mocked(vscode.window.showWarningMessage).mockReset();
        vi.mocked(vscode.window.showWarningMessage).mockResolvedValue(undefined);
    });

    afterEach(() => {
        vi.useRealTimers();
    });

    it('collapses rapid repeats for the same workbook into one popup', () => {
        reportWorkbookLocked('C:\\rapid\\Book.xlsm', 'write');
        reportWorkbookLocked('C:\\rapid\\Book.xlsm', 'read');
        reportWorkbookLocked('C:\\rapid\\Book.xlsm', 'write');
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(1);
    });

    it('shows the popup again once the throttle window elapses', () => {
        reportWorkbookLocked('C:\\elapsed\\Book.xlsm', 'write');
        vi.setSystemTime(2500);
        reportWorkbookLocked('C:\\elapsed\\Book.xlsm', 'write');
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(2);
    });

    it('does not throttle across different workbooks', () => {
        reportWorkbookLocked('C:\\distinct\\A.xlsm', 'write');
        reportWorkbookLocked('C:\\distinct\\B.xlsm', 'write');
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(2);
    });
});
