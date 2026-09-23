import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

// The notice asks Restart Manager who holds the file; no test here starts a
// real PowerShell, and each one says what the lookup found.
const lookup = vi.hoisted(() => ({ stdout: [] as string[], calls: 0 }));
vi.mock('../src/util/powershell', async (importOriginal) => ({
    ...(await importOriginal<typeof import('../src/util/powershell')>()),
    runPowerShell: vi.fn(() => {
        lookup.calls += 1;
        return {
            kill: () => undefined,
            result: Promise.resolve({ code: 0, signal: null, timedOut: false, stdoutLines: lookup.stdout, stderrLines: [] }),
        };
    }),
}));

import * as path from 'path';
import * as vscode from 'vscode';
import { annotateLockError } from '../src/fileLockHolders';
import { isProjectLockedError, reportProjectLocked } from '../src/xlideFileSystem';

/** The notice waits for the lookup, which settles in a few microtasks. */
async function settled(): Promise<void> {
    for (let i = 0; i < 10; i++) {
        await Promise.resolve();
    }
}

const platform = Object.getOwnPropertyDescriptor(process, 'platform')!;
const onWindows = (): void => { Object.defineProperty(process, 'platform', { ...platform, value: 'win32' }); };

// reportProjectLocked must surface at most one "file is open in its
// application" popup per file within a short window, so a burst of failed
// operations (or a writeFile failure followed by a re-read) never stacks
// notifications.
describe('reportProjectLocked throttling', () => {
    beforeEach(() => {
        vi.useFakeTimers();
        vi.setSystemTime(0);
        lookup.stdout = [];
        lookup.calls = 0;
        vi.mocked(vscode.window.showWarningMessage).mockReset();
        vi.mocked(vscode.window.showWarningMessage).mockResolvedValue(undefined);
    });

    afterEach(() => {
        vi.useRealTimers();
        Object.defineProperty(process, 'platform', platform);
    });

    it('collapses rapid repeats for the same project into one popup', async () => {
        reportProjectLocked('C:\\rapid\\Book.xlsm', 'write');
        reportProjectLocked('C:\\rapid\\Book.xlsm', 'read');
        reportProjectLocked('C:\\rapid\\Book.xlsm', 'write');
        await settled();
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(1);
    });

    it('shows the popup again once the throttle window elapses', async () => {
        reportProjectLocked('C:\\elapsed\\Book.xlsm', 'write');
        vi.setSystemTime(2500);
        reportProjectLocked('C:\\elapsed\\Book.xlsm', 'write');
        await settled();
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(2);
    });

    it('does not throttle across different projects', async () => {
        reportProjectLocked('C:\\distinct\\A.xlsm', 'write');
        reportProjectLocked('C:\\distinct\\B.xlsm', 'write');
        await settled();
        expect(vscode.window.showWarningMessage).toHaveBeenCalledTimes(2);
    });

    it('names the application that owns the file when nothing more is known', async () => {
        for (const [file, app] of [
            ['C:\\named\\Book.xlsm', 'Excel'],
            ['C:\\named\\Report.docm', 'Word'],
            ['C:\\named\\Deck.pptm', 'PowerPoint'],
            ['C:\\named\\Orders.accdb', 'Access'],
        ] as const) {
            vi.mocked(vscode.window.showWarningMessage).mockClear();
            reportProjectLocked(file, 'write');
            await settled();
            expect(vi.mocked(vscode.window.showWarningMessage).mock.calls[0][0])
                .toContain(`it appears to be open in ${app}.`);
        }
    });

    it('names the process Restart Manager finds holding the file', async () => {
        onWindows();
        lookup.stdout = ['XLIDE_LOCK_HOLDERS|[{"pid":4242,"appName":"Microsoft Word","image":"WINWORD.EXE"}]'];
        // A path native to the machine running the tests: the notice names the
        // file by that machine's rules, which the platform stand-in above does
        // not change, and CI runs on Linux.
        reportProjectLocked(path.join(path.resolve('held'), 'Report.docm'), 'write');
        await settled();
        expect(vi.mocked(vscode.window.showWarningMessage).mock.calls[0][0])
            .toBe('XLIDE: Cannot save "Report.docm" - it is open in Microsoft Word (WINWORD.EXE, process 4242). Close the file and try again.');
    });

    it('does not look again when the write already found the holder', async () => {
        onWindows();
        lookup.stdout = ['XLIDE_LOCK_HOLDERS|[{"pid":77,"appName":"OneDrive","image":"OneDrive.exe"}]'];
        const err = await annotateLockError(new Error('EPERM: operation not permitted, rename'), 'C:\\sync\\Book.xlsm');
        expect(lookup.calls).toBe(1);
        reportProjectLocked('C:\\sync\\Book.xlsm', 'write', err);
        await settled();
        expect(lookup.calls).toBe(1);
        expect(vi.mocked(vscode.window.showWarningMessage).mock.calls[0][0])
            .toContain('it is open in OneDrive (OneDrive.exe, process 77).');
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
