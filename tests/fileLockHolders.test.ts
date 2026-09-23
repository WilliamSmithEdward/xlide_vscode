import { afterEach, describe, expect, it, vi } from 'vitest';

const lookup = vi.hoisted(() => ({ stdout: [] as string[], scripts: [] as string[] }));
vi.mock('../src/util/powershell', async (importOriginal) => ({
    ...(await importOriginal<typeof import('../src/util/powershell')>()),
    runPowerShell: vi.fn((options: { script?: string }) => {
        lookup.scripts.push(options.script ?? '');
        return {
            kill: () => undefined,
            result: Promise.resolve({ code: 0, signal: null, timedOut: false, stdoutLines: lookup.stdout, stderrLines: [] }),
        };
    }),
}));

import {
    annotateLockError,
    buildLockHoldersScript,
    describeLockHolders,
    findFileLockHolders,
    lockHoldersOf,
    parseLockHolders,
} from '../src/fileLockHolders';

const platform = Object.getOwnPropertyDescriptor(process, 'platform')!;
const onWindows = (): void => { Object.defineProperty(process, 'platform', { ...platform, value: 'win32' }); };

describe('asking Restart Manager who holds a file', () => {
    afterEach(() => {
        Object.defineProperty(process, 'platform', platform);
        lookup.stdout = [];
        lookup.scripts = [];
    });

    it('declares the calls the way RestartManager.h does', () => {
        const script = buildLockHoldersScript("C:\\work\\Bob's Book.xlsm");
        expect(script).toContain("$targetPath = 'C:\\work\\Bob''s Book.xlsm'");
        // CCH_RM_MAX_APP_NAME 255 and CCH_RM_MAX_SVC_NAME 63, each plus a
        // terminator; a session key of CCH_RM_SESSION_KEY 32 plus one.
        expect(script).toContain('SizeConst = 256)] public string strAppName;');
        expect(script).toContain('SizeConst = 64)] public string strServiceShortName;');
        expect(script).toContain('new StringBuilder(33)');
        // RmGetList asks again, bigger, while it answers ERROR_MORE_DATA.
        expect(script).toContain('if (rc != 234 || attempt == 4)');
        // The C# rides in a here-string, whose closing mark must start a line.
        expect(script.split('\n')).toContain("'@");
    });

    it('reads the holders it prints, and an empty list as nobody', () => {
        expect(parseLockHolders(['noise', 'XLIDE_LOCK_HOLDERS|[{"pid":12,"appName":"Microsoft Excel","image":"EXCEL.EXE"}]']))
            .toEqual([{ pid: 12, appName: 'Microsoft Excel', image: 'EXCEL.EXE' }]);
        expect(parseLockHolders(['XLIDE_LOCK_HOLDERS|[]'])).toEqual([]);
        expect(parseLockHolders(['XLIDE_LOCK_HOLDERS|[{"pid":3,"appName":"Svc","image":""}]'])).toEqual([{ pid: 3, appName: 'Svc' }]);
    });

    it('has no answer when the lookup printed none', () => {
        expect(parseLockHolders([])).toBeUndefined();
        expect(parseLockHolders(['XLIDE_LOCK_HOLDERS|not json'])).toBeUndefined();
        expect(parseLockHolders(['XLIDE_LOCK_HOLDERS|{"pid":1}'])).toBeUndefined();
    });

    it('names each holder by application, executable and process', () => {
        expect(describeLockHolders([{ pid: 4242, appName: 'Microsoft Excel', image: 'EXCEL.EXE' }]))
            .toBe('Microsoft Excel (EXCEL.EXE, process 4242)');
        expect(describeLockHolders([
            { pid: 1, appName: 'Microsoft Excel', image: 'EXCEL.EXE' },
            { pid: 2, appName: 'OneDrive', image: 'OneDrive.exe' },
            { pid: 3, appName: '', image: 'backup.exe' },
        ])).toBe('Microsoft Excel (EXCEL.EXE, process 1), OneDrive (OneDrive.exe, process 2) and backup.exe (process 3)');
    });

    it('does not run anywhere but Windows', async () => {
        Object.defineProperty(process, 'platform', { ...platform, value: 'linux' });
        expect(await findFileLockHolders('/work/Book.xlsm')).toBeUndefined();
        expect(lookup.scripts).toHaveLength(0);
    });

    it('sends its script whole, the way every other script goes', async () => {
        onWindows();
        lookup.stdout = ['XLIDE_LOCK_HOLDERS|[]'];
        expect(await findFileLockHolders('C:\\work\\Book.xlsm')).toEqual([]);
        expect(lookup.scripts[0]).toBe(buildLockHoldersScript('C:\\work\\Book.xlsm'));
    });
});

describe('a lock error that names its holder', () => {
    afterEach(() => {
        Object.defineProperty(process, 'platform', platform);
        lookup.stdout = [];
        lookup.scripts = [];
    });

    it('says who holds the file, for an agent that only sees the message', async () => {
        onWindows();
        lookup.stdout = ['XLIDE_LOCK_HOLDERS|[{"pid":4242,"appName":"Microsoft Word","image":"WINWORD.EXE"}]'];
        const err = new Error("EPERM: operation not permitted, rename 'C:\\w\\.xlide-1.tmp' -> 'C:\\w\\Report.docm'");
        expect(await annotateLockError(err, 'C:\\w\\Report.docm')).toBe(err);
        expect(err.message).toBe("EPERM: operation not permitted, rename 'C:\\w\\.xlide-1.tmp' -> 'C:\\w\\Report.docm'"
            + ' (held open by Microsoft Word (WINWORD.EXE, process 4242))');
        expect(lockHoldersOf(err)).toEqual([{ pid: 4242, appName: 'Microsoft Word', image: 'WINWORD.EXE' }]);
        // Annotated once, however many layers pass it on.
        await annotateLockError(err, 'C:\\w\\Report.docm');
        expect(lookup.scripts).toHaveLength(1);
    });

    it('leaves any other error alone, and a lock nothing holds any more', async () => {
        onWindows();
        const other = new Error('Module not found: Main');
        expect(await annotateLockError(other, 'C:\\w\\Book.xlsm')).toBe(other);
        expect(lookup.scripts).toHaveLength(0);

        lookup.stdout = ['XLIDE_LOCK_HOLDERS|[]'];
        const released = new Error('EPERM: operation not permitted, rename');
        await annotateLockError(released, 'C:\\w\\Book.xlsm');
        expect(released.message).toBe('EPERM: operation not permitted, rename');
        expect(lockHoldersOf(released)).toBeUndefined();
    });
});
