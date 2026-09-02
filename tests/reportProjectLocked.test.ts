import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import * as vscode from 'vscode';
import { reportProjectLocked } from '../src/xlideFileSystem';

// reportProjectLocked must surface at most one "workbook is open in Excel"
// popup per workbook within a short window, so a burst of failed operations
// (or a writeFile failure followed by a re-read) never stacks notifications.
describe('reportProjectLocked throttling', () => {
    beforeEach(() => {
        vi.useFakeTimers();
        vi.setSystemTime(0);
        vi.mocked(vscode.window.showWarningMessage).mockReset();
        vi.mocked(vscode.window.showWarningMessage).mockResolvedValue(undefined);
    });

    afterEach(() => {
        vi.useRealTimers();
    });

    it('collapses rapid repeats for the same project into one popup', () => {
        reportProjectLocked('C:\\rapid\\Book.xlsm', 'write');
        reportProjectLocked('C:\\rapid\\Book.xlsm', 'read');
        reportProjectLocked('C:\\rapid\\Book.xlsm', 'write');
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(1);
    });

    it('shows the popup again once the throttle window elapses', () => {
        reportProjectLocked('C:\\elapsed\\Book.xlsm', 'write');
        vi.setSystemTime(2500);
        reportProjectLocked('C:\\elapsed\\Book.xlsm', 'write');
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(2);
    });

    it('does not throttle across different projects', () => {
        reportProjectLocked('C:\\distinct\\A.xlsm', 'write');
        reportProjectLocked('C:\\distinct\\B.xlsm', 'write');
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(2);
    });
});
