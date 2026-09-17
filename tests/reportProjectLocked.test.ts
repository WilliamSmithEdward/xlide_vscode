import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import * as vscode from 'vscode';
import { isProjectLockedError, reportProjectLocked } from '../src/xlideFileSystem';

// reportProjectLocked must surface at most one "file is open in its
// application" popup per file within a short window, so a burst of failed
// operations (or a writeFile failure followed by a re-read) never stacks
// notifications.
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

    it('names the application that owns the file', () => {
        for (const [file, app] of [
            ['C:\\named\\Book.xlsm', 'Excel'],
            ['C:\\named\\Report.docm', 'Word'],
            ['C:\\named\\Deck.pptm', 'PowerPoint'],
            ['C:\\named\\Orders.accdb', 'Access'],
        ] as const) {
            vi.mocked(vscode.window.showWarningMessage).mockClear();
            reportProjectLocked(file, 'write');
            expect(vi.mocked(vscode.window.showWarningMessage).mock.calls[0][0])
                .toContain(`it appears to be open in ${app}.`);
        }
    });
});

describe('isProjectLockedError', () => {
    it('recognizes the failure a save meets while the file is open in its application', () => {
        // XLIDE renames a temp file over the container, and Windows refuses
        // the rename with EPERM while Excel, Word, PowerPoint or Access has
        // the file open (measured on Office 16.0). Without this the notice
        // never showed and write coordination never started.
        expect(isProjectLockedError(
            "EPERM: operation not permitted, rename 'C:\\work\\.xlide-1-2.tmp' -> 'C:\\work\\Report.docm'",
        )).toBe(true);
        expect(isProjectLockedError("EBUSY: resource busy or locked, open 'C:\\work\\Book.xlsm'")).toBe(true);
        expect(isProjectLockedError('Module "Main" does not exist.')).toBe(false);
    });
});
