import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as vscode from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());
vi.mock('../src/util/powershell', async (original) => ({
    ...(await original<typeof import('../src/util/powershell')>()),
    runPowerShell: vi.fn(),
}));
vi.mock('../src/officeHostLauncher', async (original) => ({
    ...(await original<typeof import('../src/officeHostLauncher')>()),
    openFileInHost: vi.fn(async () => ({})),
}));

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
            expect(guarded).toContain('if (($wasReadOnly -or -not $onlyReadOnlyCopy) -and ($force -or -not $unsaved)) {');
        });

        it('never closes a copy holding unsaved work, read-only or not, outside force', () => {
            // Measured on build 16.0.20326: Excel, Word and PowerPoint all take
            // an edit in a copy opened read-only, which only Save As can keep,
            // and this close discards. It used to close such a copy anyway on
            // the grounds that a read-only copy holds nothing to lose.
            for (const file of ['C:\\work\\Book.xlsm', 'C:\\work\\Report.docm', 'C:\\work\\Deck.pptm', 'C:\\work\\Orders.accdb']) {
                const close = buildCloseFileScript(file, { force: false, onlyReadOnlyCopy: true });
                expect(close).toContain('function Test-XlideUnsavedWork($copy)');
                expect(close).toContain('$unsaved = Test-XlideUnsavedWork $file');
                expect(close).toContain('($force -or -not $unsaved)');
                expect(close).toContain('|unsaved=');
            }
        });

        it('asks each application for unsaved work in its own terms', () => {
            expect(buildCloseFileScript('C:\\work\\Book.xlsm', { force: false })).toContain('return (-not [bool]$copy.Saved)');
            expect(buildCloseFileScript('C:\\work\\Report.docm', { force: false })).toContain('return (-not [bool]$copy.Saved)');
            // MsoTriState: only msoTrue (-1) means saved.
            expect(buildCloseFileScript('C:\\work\\Deck.pptm', { force: false })).toContain('return ($copy.Saved -ne -1)');
            // Access has no document flag; each loaded object carries a dirty
            // bit, which is 2 in Access's own type library.
            const access = buildCloseFileScript('C:\\work\\Orders.accdb', { force: false });
            expect(access).toContain('$app.SysCmd(10, $type, $object.Name) -band 2');
            expect(access).toContain('2 = $app.CurrentProject.AllForms');
        });

        it('keeps its statements on lines of their own, for every application and scope', () => {
            // See the launcher's test of the same: the scripts used to be
            // joined with "; ", which ran a clause after it as a command.
            const detached = /\}\s*;\s*(else|elseif|catch|finally)\b/i;
            for (const file of ['C:\\w\\Book.xlsm', 'C:\\w\\Report.docm', 'C:\\w\\Deck.pptm', 'C:\\w\\Orders.accdb']) {
                for (const force of [false, true]) {
                    for (const onlyReadOnlyCopy of [false, true]) {
                        const script = buildCloseFileScript(file, { force, onlyReadOnlyCopy });
                        expect(script.split('\n').length).toBeGreaterThan(10);
                        expect(script).not.toMatch(detached);
                    }
                }
                expect(buildRefreshReadOnlyScript(file) ?? '').not.toMatch(detached);
            }
        });

        it('treats anything it cannot determine as unsaved', () => {
            // The two ways to be wrong: a save that asks you to close the file
            // yourself, or somebody's work. Only one of them is acceptable.
            expect(buildCloseFileScript('C:\\work\\Book.xlsm', { force: false }))
                .toMatch(/function Test-XlideUnsavedWork\(\$copy\) \{ try \{ .* \} catch \{ return \$true \} \}/);
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

        it('closes a read-only copy under block, whoever opened it', () => {
            // Word and PowerPoint lock a file even when it is open read-only.
            // It used to take XLIDE having opened the copy, so a document you
            // opened read-only to look at blocked every save until you closed
            // it by hand. The script still checks it IS read-only and holds
            // no unsaved work before it closes anything.
            expect(closeScopeForWrite(settings({ mode: 'block' }), doc)).toBe('readOnlyCopy');
            expect(closeScopeForWrite(settings({ mode: 'block' }), 'C:\\scope\\Deck.pptm')).toBe('readOnlyCopy');
            markFileOpenedByXlide(doc);
            expect(closeScopeForWrite(settings({ mode: 'block' }), doc)).toBe('readOnlyCopy');
        });

        it('closes nothing under block where a read-only open never locks', () => {
            // A read-only workbook leaves the file writable, so a lock in Excel
            // is a copy open for editing: running the close would only delay
            // the refusal. Access has no read-only open at all.
            expect(closeScopeForWrite(settings({ mode: 'block' }), 'C:\\scope\\Book.xlsm')).toBeUndefined();
            expect(closeScopeForWrite(settings({ mode: 'block' }), 'C:\\scope\\Orders.accdb')).toBeUndefined();
        });

        it('closes any open copy where the mode itself allows a close', () => {
            expect(closeScopeForWrite(settings({ mode: 'closeForce' }), doc)).toBe('any');
            markFileOpenedByXlide(doc);
            expect(closeScopeForWrite(settings({ mode: 'closeTracked' }), doc)).toBe('any');
        });

        it('falls back to the read-only rule when closeTracked may not close the file', () => {
            expect(closeScopeForWrite(settings({ mode: 'closeTracked' }), doc)).toBe('readOnlyCopy');
            expect(closeScopeForWrite(settings({ mode: 'closeTracked' }), 'C:\\scope\\Book.xlsm')).toBeUndefined();
        });
    });

    describe('buildRefreshReadOnlyScript', () => {
        const script = buildRefreshReadOnlyScript('C:\\ro\\Book.xlsm');

        it('only refreshes a file that is actually open read-only, with nothing unsaved in it', () => {
            // A read-only copy takes edits; this closes it without saving.
            expect(script).toContain('if ($file -and [bool]$file.ReadOnly -and -not (Test-XlideUnsavedWork $file)) {');
            expect(script).toContain('function Test-XlideUnsavedWork($copy)');
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

        it('keeps the reader\'s place: sheet, selection, active cell and scroll', () => {
            const at = (text: string): number => script!.indexOf(text);
            for (const read of [
                '$place.sheet = [string]$window.ActiveSheet.Name',
                '$place.selection = [string]$window.RangeSelection.Address()',
                '$place.activeCell = [string]$window.ActiveCell.Address()',
                '$place.scrollRow = [int]$window.ScrollRow',
            ]) {
                expect(at(read), read).toBeGreaterThan(-1);
                expect(at(read), `${read} comes before the close`).toBeLessThan(at('try { $file.Close($false) } catch { }'));
            }
            // Put back into the copy just opened, sheet first, scroll last.
            const reopened = at('$file = $app.Workbooks.Open($targetPath, 0, $true)');
            expect(reopened).toBeGreaterThan(-1);
            expect(at('$file.Sheets.Item($place.sheet).Activate()')).toBeGreaterThan(reopened);
            expect(at('$file.ActiveSheet.Range($place.selection).Select()')).toBeGreaterThan(at('$file.Sheets.Item($place.sheet).Activate()'));
            expect(at('$window.ScrollRow = $place.scrollRow')).toBeGreaterThan(at('$file.ActiveSheet.Range($place.activeCell).Activate()'));
            // And the workbook you had in front stays in front.
            expect(script).toContain('$place.otherActive = [string]$active.FullName');
            expect(script).toContain('if ($other.FullName -ieq $place.otherActive) { $other.Activate(); break }');
        });

        it('does not run Workbook_Open again on every save', () => {
            // Events go off for the reopen and come back whatever happens.
            expect(script).toContain([
                '    $eventsWere = $app.EnableEvents',
                '    try {',
                '      $app.EnableEvents = $false',
            ].join('\n'));
            expect(script).toContain([
                '    finally {',
                '      $app.EnableEvents = $eventsWere',
                '    }',
            ].join('\n'));
            // Word and PowerPoint have no such switch.
            expect(buildRefreshReadOnlyScript('C:\\ro\\Report.docm')).not.toContain('EnableEvents');
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
                reopenReadOnlyAfterSave: true,
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
                // Every coordination script goes over whole, as the script
                // itself, never as a -Command argument.
                expect(options.args).toBeUndefined();
                scripts.push(options.script ?? '');
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
            expect(openFileInHost).toHaveBeenCalledWith(doc, { attachToRunning: true, readOnly: false, background: true }, expect.any(Function));
            forgetFileOpenedByXlide(doc);
        });

        it('under block, closes a read-only copy and reopens it read-only', async () => {
            const deck = 'C:\\flow\\Deck.pptm';
            markFileOpenedByXlide(deck);
            useSettings({});
            closeReports('XLIDE_CLOSE|closed=True|locked=False|found=True|wasReadOnly=True|forced=False|unsaved=False');
            const write = lockedWrite(deck, 1);

            await expect(runWriteWithHostCoordination(deck, write)).resolves.toBe('written');

            expect(scripts[0]).toContain('GetActiveObject("PowerPoint.Application")');
            expect(scripts[0]).toContain('$onlyReadOnlyCopy = $true');
            expect(scripts[0]).toContain('$force = $false');
            expect(openFileInHost).toHaveBeenCalledWith(deck, { attachToRunning: true, readOnly: true, background: true }, expect.any(Function));
            forgetFileOpenedByXlide(deck);
        });

        it('under block, frees a document you opened read-only yourself', async () => {
            // The reported case: a Word document open read-only to look at,
            // which XLIDE did not open, used to block every save until it was
            // closed by hand.
            const doc = 'C:\\flow\\Mine.docm';
            forgetFileOpenedByXlide(doc);
            useSettings({});
            closeReports('XLIDE_CLOSE|closed=True|locked=False|found=True|wasReadOnly=True|forced=False|unsaved=False');
            const write = lockedWrite(doc, 1);

            await expect(runWriteWithHostCoordination(doc, write)).resolves.toBe('written');

            expect(write).toHaveBeenCalledTimes(2);
            expect(scripts[0]).toContain('$onlyReadOnlyCopy = $true');
            expect(openFileInHost).toHaveBeenCalledWith(doc, { attachToRunning: true, readOnly: true, background: true }, expect.any(Function));
            forgetFileOpenedByXlide(doc);
        });

        it('puts the reader back where they were in the copy it closed', async () => {
            const doc = 'C:\\flow\\Place.docm';
            useSettings({});
            vi.mocked(runPowerShell).mockImplementation((options) => {
                scripts.push(options.script ?? '');
                return {
                    kill: () => undefined,
                    result: Promise.resolve({
                        code: 0, signal: null, timedOut: false, stderrLines: [], stdoutLines: [
                            'XLIDE_PLACE|{"start":120,"end":131,"scrolled":42}',
                            'XLIDE_CLOSE|closed=True|locked=False|found=True|wasReadOnly=True|forced=False|unsaved=False',
                        ],
                    }),
                };
            });

            await expect(runWriteWithHostCoordination(doc, lockedWrite(doc, 1))).resolves.toBe('written');

            // Read before the close, in the close script.
            expect(scripts[0].indexOf('$place.start = [int]$window.Selection.Start'))
                .toBeLessThan(scripts[0].indexOf('try { $file.Close(0); $closed = $true } catch { }'));
            expect(openFileInHost).toHaveBeenCalledWith(doc, {
                attachToRunning: true, readOnly: true, background: true, place: { start: 120, end: 131, scrolled: 42 },
            }, expect.any(Function));
            forgetFileOpenedByXlide(doc);
        });

        it('refreshes once more for a save that arrives while a refresh runs, never piling them up', async () => {
            const book = 'C:\\flow\\Live.xlsm';
            useSettings({ 'officeIntegration.reopenReadOnlyAfterSave': true });
            const finish: Array<() => void> = [];
            vi.mocked(runPowerShell).mockImplementation((options) => {
                scripts.push(options.script ?? '');
                return {
                    kill: () => undefined,
                    result: new Promise((resolve) => finish.push(() => resolve({
                        code: 0, signal: null, timedOut: false, stderrLines: [], stdoutLines: ['XLIDE_REFRESH|refreshed=True'],
                    }))),
                };
            });
            const save = vi.fn(async () => 'saved');

            await runWriteWithHostCoordination(book, save);
            await runWriteWithHostCoordination(book, save);
            await runWriteWithHostCoordination(book, save);
            await vi.waitFor(() => expect(finish).toHaveLength(1));
            finish[0]();
            // The two saves during the first refresh make one more, not two.
            await vi.waitFor(() => expect(finish).toHaveLength(2));
            finish[1]();
            await new Promise((resolve) => setTimeout(resolve, 20));
            expect(scripts.filter((script) => script.includes('XLIDE_REFRESH|'))).toHaveLength(2);
            forgetFileOpenedByXlide(book);
        });

        it('leaves a read-only copy with unsaved work open, and the save fails', async () => {
            const doc = 'C:\\flow\\Typed.docm';
            useSettings({});
            closeReports('XLIDE_CLOSE|closed=False|locked=True|found=True|wasReadOnly=True|forced=False|unsaved=True');
            const write = lockedWrite(doc, 2);
            const log = vi.fn();

            await expect(runWriteWithHostCoordination(doc, write, log)).rejects.toThrow(/EPERM/);

            // Nothing was closed, so nothing is reopened over the top of it.
            expect(openFileInHost).not.toHaveBeenCalled();
            expect(log).toHaveBeenCalledWith(expect.stringContaining('holds unsaved work'));
        });

        it('under block, runs no close at all where a read-only open never locks', async () => {
            // A lock on an Access database is always an open for editing.
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
