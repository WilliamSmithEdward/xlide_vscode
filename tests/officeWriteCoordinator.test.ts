import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as vscode from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
vi.mock('../src/util/powershell', async (original) => ({
    ...(await original<typeof import('../src/util/powershell')>()),
    runPowerShell: vi.fn(),
}));
vi.mock('../src/officeHostLauncher', () => ({ openFileInHost: vi.fn(async () => undefined) }));

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { fakeConfig } from './helpers/fakeConfig';
import { checkProjectFile } from '../src/projectFileChanges';
import { runPowerShell } from '../src/util/powershell';
import { openFileInHost } from '../src/officeHostLauncher';
import {
    buildCloseFileScript,
    buildRefreshReadOnlyScript,
    closeScopeForWrite,
    forgetFileOpenedByXlide,
    markFileOpenedByXlide,
    resolveHostCoordinationSettings,
    resolveReopenReadOnly,
    runWriteWithHostCoordination,
    shouldAttemptClose,
    wasFileOpenedByXlide,
    withFileReopenSuppressed,
    type HostCoordinationSettings,
} from '../src/officeWriteCoordinator';

const settings = (over: Partial<HostCoordinationSettings> = {}): HostCoordinationSettings => ({
    mode: 'block',
    trackOpenedFiles: true,
    reopenAfterClose: true,
    reopenMode: 'lastState',
    reopenReadOnlyAfterSave: false,
    ...over,
});

describe('officeWriteCoordinator', () => {
    describe('buildCloseFileScript', () => {
        const script = buildCloseFileScript('C:\\work\\Book.xlsm', { force: false });

        it('targets the file by full path and basename', () => {
            expect(script).toContain("$targetPath = 'C:\\work\\Book.xlsm'");
            expect(script).toContain("$targetName = 'Book.xlsm'");
        });

        it('attaches to the running Excel and closes without saving Excel\'s copy', () => {
            expect(script).toContain('GetActiveObject("Excel.Application")');
            expect(script).toContain('$files = @($app.Workbooks)');
            expect(script).toContain('$file.Close($false)');
        });

        it('reports the close, lock, and prior read-only state through the sentinel', () => {
            expect(script).toContain('XLIDE_CLOSE|closed=');
            expect(script).toContain('|locked=');
            expect(script).toContain('|wasReadOnly=');
            expect(script).toContain('|forced=');
            // captures the file's read-only state before closing it
            expect(script).toContain('$wasReadOnly = [bool]$file.ReadOnly');
        });

        it('passes the force flag and only kills Excel under force', () => {
            expect(script).toContain('$force = $false');
            const forceScript = buildCloseFileScript('C:\\work\\Book.xlsm', { force: true });
            expect(forceScript).toContain('$force = $true');
            // The kill is runtime-gated by `if ($locked -and $force)`.
            expect(forceScript).toContain('if ($locked -and $force)');
            expect(forceScript).toContain('Get-Process -Name EXCEL ');
            expect(forceScript).toContain('Stop-Process -Force');
        });

        it('escapes single quotes in the interpolated path', () => {
            const quoted = buildCloseFileScript("C:\\work\\Bob's Book.xlsm", { force: false });
            expect(quoted).toContain("$targetPath = 'C:\\work\\Bob''s Book.xlsm'");
            expect(quoted).toContain("$targetName = 'Bob''s Book.xlsm'");
        });

        // Every call below was measured live on Office 16.0 against scratch
        // copies of the fixtures: the attach, the collection, the read-only
        // test and a close of an EDITED file that returns without prompting.
        it.each([
            ['Word', 'C:\\work\\Report.docm', 'Word.Application', '$app.Documents', '[bool]$file.ReadOnly', '$file.Close(0)', 'WINWORD'],
            ['PowerPoint', 'C:\\work\\Deck.pptm', 'PowerPoint.Application', '$app.Presentations', '($file.ReadOnly -ne 0)', '$file.Close()', 'POWERPNT'],
        ])('closes a %s file through its own application', (_app, file, progId, collection, readOnly, close, processName) => {
            const hostScript = buildCloseFileScript(file, { force: true });
            expect(hostScript).toContain(`GetActiveObject("${progId}")`);
            expect(hostScript).toContain(`$files = @(${collection})`);
            expect(hostScript).toContain(`$wasReadOnly = ${readOnly}`);
            expect(hostScript).toContain(`try { ${close}; $closed = $true } catch { }`);
            expect(hostScript).toContain(`Get-Process -Name ${processName} `);
            expect(hostScript).not.toContain('Excel');
            expect(hostScript).not.toContain('EXCEL');
        });

        it('closes an Access database by the one database the instance holds', () => {
            const access = buildCloseFileScript('C:\\work\\Orders.accdb', { force: true });
            expect(access).toContain('GetActiveObject("Access.Application")');
            expect(access).toContain('$open = $app.CurrentProject.FullName');
            expect(access).toContain('if ($open -ieq $targetPath) { $file = $app.CurrentProject }');
            expect(access).toContain('try { $app.CloseCurrentDatabase(); $closed = $true } catch { }');
            // A database has no read-only open.
            expect(access).toContain('$wasReadOnly = $false');
            expect(access).toContain('Get-Process -Name MSACCESS ');
        });

        it('never starts an application just to close a file', () => {
            for (const file of ['C:\\w\\a.xlsm', 'C:\\w\\a.docm', 'C:\\w\\a.pptm', 'C:\\w\\a.accdb']) {
                expect(buildCloseFileScript(file, { force: true })).not.toContain('New-Object');
            }
        });

        it('prefers the full path, and never closes another local file that shares the name', () => {
            // Word and PowerPoint can hold Report.docm from two folders at once,
            // and a close discards unsaved edits, so the name alone is only
            // trusted when Office reports the path in a form that cannot be
            // compared (a OneDrive URL, a UNC path).
            const word = buildCloseFileScript('C:\\work\\Report.docm', { force: false });
            const byPath = word.indexOf('if ($f.FullName -ieq $targetPath)');
            const byName = word.indexOf('($f.Name -ieq $targetName) -and -not (Test-XlideOtherLocalFile $f.FullName)');
            expect(byPath).toBeGreaterThan(-1);
            expect(byName).toBeGreaterThan(byPath);
            expect(word).toContain('function Test-XlideOtherLocalFile($openPath)');
            expect(word).toContain('($openPath -match "^[A-Za-z]:\\\\") -and ($targetPath -match "^[A-Za-z]:\\\\")');
        });

        it('can be told to close only a read-only copy', () => {
            expect(script).toContain('$onlyReadOnlyCopy = $false');
            const guarded = buildCloseFileScript('C:\\work\\Report.docm', { force: false, onlyReadOnlyCopy: true });
            expect(guarded).toContain('$onlyReadOnlyCopy = $true');
            expect(guarded).toContain('if ($wasReadOnly -or -not $onlyReadOnlyCopy) {');
        });

        it('waits briefly for the application to let go of the file it closed', () => {
            expect(script).toContain('for ($__i = 0; $locked -and $closed -and $__i -lt 5; $__i++)');
        });
    });

    describe('XLIDE-opened file tracking', () => {
        const wb = 'C:\\track\\Book.xlsm';
        beforeEach(() => forgetFileOpenedByXlide(wb));

        it('remembers and forgets, ignoring case and separator casing', () => {
            expect(wasFileOpenedByXlide(wb)).toBe(false);
            markFileOpenedByXlide(wb);
            expect(wasFileOpenedByXlide('c:\\TRACK\\book.xlsm')).toBe(true);
            forgetFileOpenedByXlide(wb);
            expect(wasFileOpenedByXlide(wb)).toBe(false);
        });
    });

    describe('shouldAttemptClose', () => {
        const wb = 'C:\\policy\\Book.xlsm';
        beforeEach(() => forgetFileOpenedByXlide(wb));

        it('never closes under block mode, even for an XLIDE-opened file', () => {
            markFileOpenedByXlide(wb);
            expect(shouldAttemptClose(settings({ mode: 'block' }), wb)).toBe(false);
        });

        it('always closes under closeForce', () => {
            expect(shouldAttemptClose(settings({ mode: 'closeForce' }), wb)).toBe(true);
        });

        it('closeTracked closes only XLIDE-opened files when tracking is on', () => {
            const tracked = settings({ mode: 'closeTracked', trackOpenedFiles: true });
            expect(shouldAttemptClose(tracked, wb)).toBe(false);
            markFileOpenedByXlide(wb);
            expect(shouldAttemptClose(tracked, wb)).toBe(true);
        });

        it('closeTracked closes any matching file when tracking is off', () => {
            const untracked = settings({ mode: 'closeTracked', trackOpenedFiles: false });
            expect(shouldAttemptClose(untracked, wb)).toBe(true);
        });
    });

    describe('closeScopeForWrite', () => {
        const doc = 'C:\\scope\\Report.docm';
        beforeEach(() => forgetFileOpenedByXlide(doc));

        it('closes nothing for a file XLIDE did not open, under block', () => {
            expect(closeScopeForWrite(settings({ mode: 'block' }), doc)).toBeUndefined();
        });

        it('closes only a read-only copy of a file XLIDE opened, under block', () => {
            // Word and PowerPoint lock a file even when it is open read-only,
            // which is how F5 leaves it, so a save after F5 depends on this.
            markFileOpenedByXlide(doc);
            expect(closeScopeForWrite(settings({ mode: 'block' }), doc)).toBe('readOnlyCopy');
        });

        it('closes any open copy where the mode itself allows a close', () => {
            expect(closeScopeForWrite(settings({ mode: 'closeForce' }), doc)).toBe('any');
            markFileOpenedByXlide(doc);
            expect(closeScopeForWrite(settings({ mode: 'closeTracked' }), doc)).toBe('any');
        });

        it('falls back to the read-only rule when closeTracked may not close the file', () => {
            expect(closeScopeForWrite(settings({ mode: 'closeTracked' }), doc)).toBeUndefined();
        });
    });

    describe('buildRefreshReadOnlyScript', () => {
        const script = buildRefreshReadOnlyScript('C:\\ro\\Book.xlsm');

        it('only refreshes a file that is actually open read-only', () => {
            expect(script).toContain('if ($file -and [bool]$file.ReadOnly) {');
            // close + reopen read-only ($true) so a closed file is never opened
            expect(script).toContain('$file.Close($false)');
            expect(script).toContain('$app.Workbooks.Open($targetPath, 0, $true)');
            expect(script).toContain('XLIDE_REFRESH|refreshed=');
        });

        it('escapes single quotes in the path', () => {
            const quoted = buildRefreshReadOnlyScript("C:\\ro\\Bob's Book.xlsm");
            expect(quoted).toContain("$targetPath = 'C:\\ro\\Bob''s Book.xlsm'");
        });

        it('reopens read-only with each application\'s own open call', () => {
            expect(buildRefreshReadOnlyScript('C:\\ro\\Report.docm'))
                .toContain('$app.Documents.Open($targetPath, $false, $true, $false)');
            expect(buildRefreshReadOnlyScript('C:\\ro\\Deck.pptm'))
                .toContain('$app.Presentations.Open($targetPath, -1, 0, -1)');
        });

        it('has nothing to refresh in Access, which has no read-only open', () => {
            expect(buildRefreshReadOnlyScript('C:\\ro\\Orders.accdb')).toBeUndefined();
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

    describe('withFileReopenSuppressed', () => {
        it('runs the wrapped function and returns its result', async () => {
            await expect(
                withFileReopenSuppressed('C:\\s\\Book.xlsm', async () => 42),
            ).resolves.toBe(42);
        });

        it('restores suppression state even when the wrapped function throws', async () => {
            await expect(
                withFileReopenSuppressed('C:\\s\\Book.xlsm', async () => { throw new Error('boom'); }),
            ).rejects.toThrow('boom');
            // The refcount must not be left stuck, or the reopen would be
            // permanently suppressed for this file.
            await expect(
                withFileReopenSuppressed('C:\\s\\Book.xlsm', async () => 'ok'),
            ).resolves.toBe('ok');
        });

        it('scopes suppression per file (one file does not suppress another)', async () => {
            // While A is suppressed, B must be free to resolve independently; the
            // per-path keying is what prevents an F5 on A from suppressing B's refresh.
            let bRan = false;
            await withFileReopenSuppressed('C:\\s\\A.xlsm', async () => {
                await withFileReopenSuppressed('C:\\s\\B.xlsm', async () => { bRan = true; });
            });
            expect(bRan).toBe(true);
        });
    });

    describe('resolveHostCoordinationSettings', () => {
        it('resolves the safe defaults from an unset configuration', () => {
            expect(resolveHostCoordinationSettings()).toEqual({
                mode: 'block',
                trackOpenedFiles: true,
                reopenAfterClose: true,
                reopenMode: 'lastState',
                reopenReadOnlyAfterSave: false,
            });
        });

        it('still honors a value stored under the setting\'s old Excel name', () => {
            const config = fakeConfig(
                { 'excelIntegration.coordinationMode': 'closeTracked', 'excelIntegration.trackOpenedWorkbooks': false },
                new Set(['excelIntegration.coordinationMode', 'excelIntegration.trackOpenedWorkbooks']),
            );
            vi.mocked(vscode.workspace.getConfiguration).mockReturnValueOnce(config);
            expect(resolveHostCoordinationSettings()).toMatchObject({ mode: 'closeTracked', trackOpenedFiles: false });
        });
    });

    describe('runWriteWithHostCoordination', () => {
        // The exact error a save meets while the file is open in its
        // application, measured against Excel, Word, PowerPoint and Access.
        const lockError = (file: string): Error => new Error(
            `EPERM: operation not permitted, rename 'C:\\work\\.xlide-1-2.tmp' -> '${file}'`,
        );
        const platform = Object.getOwnPropertyDescriptor(process, 'platform')!;
        let scripts: string[];

        const useSettings = (values: Record<string, unknown>): void => {
            vi.mocked(vscode.workspace.getConfiguration).mockReturnValue(fakeConfig(values, new Set(Object.keys(values))));
        };
        const closeReports = (sentinel: string): void => {
            vi.mocked(runPowerShell).mockImplementation((options) => {
                scripts.push(options.args[1]);
                return {
                    kill: () => undefined,
                    result: Promise.resolve({
                        code: 0, signal: null, timedOut: false, stdoutLines: [sentinel], stderrLines: [],
                    }),
                };
            });
        };
        /** A write that fails with the lock error `failures` times, then succeeds. */
        const lockedWrite = (file: string, failures: number) => {
            let calls = 0;
            return vi.fn(async () => {
                calls += 1;
                if (calls <= failures) { throw lockError(file); }
                return 'written';
            });
        };

        beforeEach(() => {
            scripts = [];
            Object.defineProperty(process, 'platform', { ...platform, value: 'win32' });
            vi.mocked(openFileInHost).mockClear();
            vi.mocked(runPowerShell).mockReset();
        });

        afterEach(() => {
            Object.defineProperty(process, 'platform', platform);
            vi.mocked(vscode.workspace.getConfiguration).mockReset();
            vi.mocked(vscode.workspace.getConfiguration).mockImplementation(() => fakeConfig({}));
        });

        it('closes a locked Word document in Word, retries, and puts the view back', async () => {
            const doc = 'C:\\flow\\Report.docm';
            markFileOpenedByXlide(doc);
            useSettings({ 'officeIntegration.coordinationMode': 'closeTracked' });
            closeReports('XLIDE_CLOSE|closed=True|locked=False|found=True|wasReadOnly=False|forced=False');
            const write = lockedWrite(doc, 1);

            await expect(runWriteWithHostCoordination(doc, write)).resolves.toBe('written');

            expect(write).toHaveBeenCalledTimes(2);
            expect(scripts).toHaveLength(1);
            expect(scripts[0]).toContain('GetActiveObject("Word.Application")');
            expect(scripts[0]).toContain('$onlyReadOnlyCopy = $false');
            // lastState: it was open for editing, so it goes back that way.
            expect(openFileInHost).toHaveBeenCalledWith(doc, { attachToRunning: true, readOnly: false }, expect.any(Function));
            forgetFileOpenedByXlide(doc);
        });

        it('under block, closes only the read-only copy XLIDE opened and reopens it read-only', async () => {
            const deck = 'C:\\flow\\Deck.pptm';
            markFileOpenedByXlide(deck);
            useSettings({});
            closeReports('XLIDE_CLOSE|closed=True|locked=False|found=True|wasReadOnly=True|forced=False');
            const write = lockedWrite(deck, 1);

            await expect(runWriteWithHostCoordination(deck, write)).resolves.toBe('written');

            expect(scripts[0]).toContain('GetActiveObject("PowerPoint.Application")');
            expect(scripts[0]).toContain('$onlyReadOnlyCopy = $true');
            expect(scripts[0]).toContain('$force = $false');
            expect(openFileInHost).toHaveBeenCalledWith(deck, { attachToRunning: true, readOnly: true }, expect.any(Function));
            forgetFileOpenedByXlide(deck);
        });

        it('under block, leaves a file XLIDE did not open alone and surfaces the lock', async () => {
            const db = 'C:\\flow\\Orders.accdb';
            useSettings({});
            const write = lockedWrite(db, 1);

            await expect(runWriteWithHostCoordination(db, write)).rejects.toThrow(/EPERM/);

            expect(write).toHaveBeenCalledTimes(1);
            expect(scripts).toEqual([]);
            expect(openFileInHost).not.toHaveBeenCalled();
        });

        it('never opens a file that was not open: a lock that clears on its own reopens nothing', async () => {
            const book = 'C:\\flow\\Book.xlsm';
            useSettings({ 'officeIntegration.coordinationMode': 'closeForce' });
            closeReports('XLIDE_CLOSE|closed=False|locked=False|found=False|wasReadOnly=False|forced=False');
            const write = lockedWrite(book, 1);

            await expect(runWriteWithHostCoordination(book, write)).resolves.toBe('written');

            expect(openFileInHost).not.toHaveBeenCalled();
        });

        it('leaves the file closed when reopening is turned off', async () => {
            const book = 'C:\\flow\\Closed.xlsm';
            markFileOpenedByXlide(book);
            useSettings({
                'officeIntegration.coordinationMode': 'closeTracked',
                'officeIntegration.reopenAfterClose': false,
            });
            closeReports('XLIDE_CLOSE|closed=True|locked=False|found=True|wasReadOnly=False|forced=False');

            await expect(runWriteWithHostCoordination(book, lockedWrite(book, 1))).resolves.toBe('written');

            expect(openFileInHost).not.toHaveBeenCalled();
            // It is no longer open anywhere, so it is no longer XLIDE's to close.
            expect(wasFileOpenedByXlide(book)).toBe(false);
        });

        it('does not race a caller that is about to reopen the file itself', async () => {
            const doc = 'C:\\flow\\F5.docm';
            markFileOpenedByXlide(doc);
            useSettings({});
            closeReports('XLIDE_CLOSE|closed=True|locked=False|found=True|wasReadOnly=True|forced=False');

            await withFileReopenSuppressed(doc, () => runWriteWithHostCoordination(doc, lockedWrite(doc, 1)));

            expect(scripts).toHaveLength(1);
            expect(openFileInHost).not.toHaveBeenCalled();
            forgetFileOpenedByXlide(doc);
        });

        it('passes any other failure straight through', async () => {
            useSettings({ 'officeIntegration.coordinationMode': 'closeForce' });
            const write = vi.fn(async () => { throw new Error('Module "Main" does not exist.'); });

            await expect(runWriteWithHostCoordination('C:\\flow\\Book.xlsm', write)).rejects.toThrow('does not exist');

            expect(write).toHaveBeenCalledTimes(1);
            expect(scripts).toEqual([]);
        });

        it('counts its write as XLIDE own, so the new file is not taken for a change made elsewhere', async () => {
            useSettings({ 'officeIntegration.reopenReadOnlyAfterSave': false });
            const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-coordinated-'));
            const book = path.join(dir, 'Book.xlsm');
            try {
                fs.writeFileSync(book, 'before');
                checkProjectFile(book);

                await runWriteWithHostCoordination(book, async () => fs.writeFileSync(book, 'after the write'));

                expect(checkProjectFile(book)).toBe(false);
                fs.writeFileSync(book, 'saved by the VBE afterwards');
                expect(checkProjectFile(book)).toBe(true);
            } finally {
                fs.rmSync(dir, { recursive: true, force: true });
            }
        });

        it('writes a VB6 project straight through: no application holds its files', async () => {
            useSettings({ 'officeIntegration.coordinationMode': 'closeForce' });
            const write = lockedWrite('C:\\flow\\App.vbp', 1);

            await expect(runWriteWithHostCoordination('C:\\flow\\App.vbp', write)).rejects.toThrow(/EPERM/);

            expect(scripts).toEqual([]);
        });
    });
});
