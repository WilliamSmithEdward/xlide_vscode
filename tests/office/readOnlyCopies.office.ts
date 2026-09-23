// Read-only copies of a file in its application, and XLIDE's saves while
// one is open: the lock each application holds, never discarding work typed
// into a read-only copy, and a copy XLIDE puts back where the reader was and
// behind the editor.
//
// Word and PowerPoint run through XLIDE's own attach paths, so only when
// neither is running when the check starts. Access runs in an instance the
// check starts itself.

import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import {
    buildHostOpenScript,
    buildPowerPointMacroLaunchScript,
    buildWordMacroLaunchScript,
    hostOpenOutcome,
} from '../../src/officeHostLauncher';
import { runWriteWithHostCoordination } from '../../src/officeWriteCoordinator';
import { unsavedWorkFunction } from '../../src/officeFileState';
import { readModule, writeModule } from '../../src/vba/projectService';
import { foreground, foregroundAwayFrom, hostIsFree, ps, psFile, q, releaseHost, scratchCopy, sleep } from './officeHarness';
import { settings } from './vscodeStub';

const log = (): void => undefined;
const probe = (file: string, body: string) => async (): Promise<unknown> =>
    writeModule(file, 'XlideProbe', `Public Sub Probe()\r\n    ' ${body}\r\nEnd Sub\r\n`, 'standard');

/** `name=value` lines a probe script prints, as a record. */
function fields(lines: readonly string[]): Record<string, string> {
    return Object.fromEntries(lines.filter((line) => line.includes('=')).map((line) => {
        const at = line.indexOf('=');
        return [line.slice(0, at), line.slice(at + 1)];
    }));
}

const wordFree = await hostIsFree('word');

describe.runIf(wordFree)('Word', () => {
    let file: string;
    const open = (readOnly: boolean): string => buildHostOpenScript({ host: 'word', filePath: file, attachToRunning: true, readOnly });
    const inDocument = (lines: string[]): string => [
        `$targetPath = ${q(file)}`,
        '$app = $null',
        'try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application") } catch { }',
        '$d = $null',
        'if ($app) { foreach ($x in @($app.Documents)) { if ($x.FullName -ieq $targetPath) { $d = $x } } }',
        'if (-not $d) { "state=closed" }',
        'else {',
        ...lines.map((line) => `  ${line}`),
        '}',
    ].join('\n');
    const state = async (): Promise<Record<string, string>> => fields((await ps(inDocument([
        '"readOnly=" + $d.ReadOnly',
        '"saved=" + $d.Saved',
        '"typed=" + $d.Content.Text.Contains("XLIDE-TYPED")',
        '"start=" + $d.ActiveWindow.Selection.Start',
        '"end=" + $d.ActiveWindow.Selection.End',
    ]))).out);

    beforeAll(() => {
        settings.clear();
        file = scratchCopy('WordShapesFixture.docm', 'copies-word');
    });

    afterAll(async () => {
        await releaseHost('word');
    });

    it('turns a copy open for editing read-only, when asked to open read-only', async () => {
        await ps(open(false));
        expect(await state()).toMatchObject({ readOnly: 'False' });
        const result = await ps(open(true));
        expect(hostOpenOutcome('word', true, result.out)).toEqual({});
        expect(await state()).toMatchObject({ readOnly: 'True' });
    });

    it('saves through a read-only copy it did not open, and puts it back where the reader was, behind the editor', async () => {
        await ps(inDocument(['$d.ActiveWindow.Selection.SetRange(10, 25)']));
        expect(await state()).toMatchObject({ start: '10', end: '25' });
        const front = await foregroundAwayFrom('WINWORD', () => runWriteWithHostCoordination(file, probe(file, 'first'), log));

        await runWriteWithHostCoordination(file, probe(file, 'first'), log);

        expect(readModule(file, 'XlideProbe', false).source).toContain('first');
        expect(await state()).toMatchObject({ readOnly: 'True', start: '10', end: '25' });
        const after = await foreground();
        expect(after.hwnd, `in front before: ${front.process}, after: ${after.process}`).toBe(front.hwnd);
    });

    it('refuses a save that would discard typing in the read-only copy, and says Word holds the file', async () => {
        await ps(inDocument(['$d.Content.InsertAfter("XLIDE-TYPED")']));
        expect(await state()).toMatchObject({ saved: 'False', typed: 'True' });
        const failure = await runWriteWithHostCoordination(file, probe(file, 'second'), log).then(() => undefined, (err: Error) => err);
        expect(failure?.message).toMatch(/EPERM/);
        expect(failure?.message).toMatch(/\(held open by Microsoft Word \(WINWORD\.EXE, process \d+\)\)$/);
        expect(await state()).toMatchObject({ readOnly: 'True', saved: 'False', typed: 'True' });
    });

    it('will not reopen over that typing to run a macro either', async () => {
        const result = await ps(buildWordMacroLaunchScript(file, 'XlideProbe.Probe'));
        expect(result.err.join(' ')).toMatch(/REOPEN_BLOCKED\|.*never saved/);
        expect(await state()).toMatchObject({ typed: 'True' });
    });

    it('leaves a copy open for editing with unsaved work exactly as it is', async () => {
        await releaseHost('word');
        await ps(open(false));
        await ps(inDocument(['$d.Content.InsertAfter("XLIDE-TYPED")']));
        const result = await ps(open(true));
        expect(hostOpenOutcome('word', true, result.out)).toEqual({ keptEditing: 'unsaved' });
        expect(await state()).toMatchObject({ readOnly: 'False', saved: 'False', typed: 'True' });
    });
});

const powerpointFree = await hostIsFree('powerpoint');

describe.runIf(powerpointFree)('PowerPoint', () => {
    let file: string;
    const open = (readOnly: boolean): string => buildHostOpenScript({ host: 'powerpoint', filePath: file, attachToRunning: true, readOnly });
    const inPresentation = (lines: string[]): string => [
        `$targetPath = ${q(file)}`,
        '$app = $null',
        'try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch { }',
        '$p = $null',
        'if ($app) { foreach ($x in @($app.Presentations)) { if ($x.FullName -ieq $targetPath) { $p = $x } } }',
        'if (-not $p) { "state=closed" }',
        'else {',
        ...lines.map((line) => `  ${line}`),
        '}',
    ].join('\n');
    const state = async (): Promise<Record<string, string>> => fields((await ps(inPresentation([
        '"readOnly=" + $p.ReadOnly',
        '"saved=" + $p.Saved',
        '"slides=" + $p.Slides.Count',
        '"slide=" + $p.Windows.Item(1).View.Slide.SlideIndex',
    ]))).out);

    beforeAll(() => {
        settings.clear();
        file = scratchCopy('PowerPointShapesFixture.pptm', 'copies-powerpoint');
    });

    afterAll(async () => {
        await releaseHost('powerpoint');
    });

    it('turns a copy open for editing read-only, when asked to open read-only', async () => {
        await ps(open(false));
        expect(await state()).toMatchObject({ readOnly: '0' });
        const result = await ps(open(true));
        expect(hostOpenOutcome('powerpoint', true, result.out)).toEqual({});
        // MsoTriState: msoTrue is -1.
        expect(await state()).toMatchObject({ readOnly: '-1' });
    });

    it('saves through the read-only copy and puts it back on the slide the reader was on, behind the editor', async () => {
        await ps(inPresentation(['$p.Windows.Item(1).View.GotoSlide(2)']));
        expect(await state()).toMatchObject({ slide: '2' });
        const front = await foregroundAwayFrom('POWERPNT', () => runWriteWithHostCoordination(file, probe(file, 'first'), log));

        await runWriteWithHostCoordination(file, probe(file, 'first'), log);

        expect(readModule(file, 'XlideProbe', false).source).toContain('first');
        expect(await state()).toMatchObject({ readOnly: '-1', slide: '2' });
        const after = await foreground();
        expect(after.hwnd, `in front before: ${front.process}, after: ${after.process}`).toBe(front.hwnd);
    });

    it('refuses a save that would discard a slide added to the read-only copy', async () => {
        // ppLayoutBlank is 12, from src/analyzer/host/powerpointObjectModelData.ts.
        await ps(inPresentation(['[void]$p.Slides.Add($p.Slides.Count + 1, 12)']));
        const before = await state();
        expect(before).toMatchObject({ saved: '0', slides: '3' });
        const failure = await runWriteWithHostCoordination(file, probe(file, 'second'), log).then(() => undefined, (err: Error) => err);
        expect(failure?.message).toMatch(/\(held open by .*POWERPNT\.EXE, process \d+\)\)$/);
        expect(await state()).toMatchObject({ saved: '0', slides: '3' });

        const f5 = await ps(buildPowerPointMacroLaunchScript(file, 'XlideProbe.Probe'));
        expect(f5.err.join(' ')).toMatch(/REOPEN_BLOCKED\|.*never saved/);
        expect(await state()).toMatchObject({ slides: '3' });
    });
});

const accessFree = await hostIsFree('access');

describe.runIf(accessFree)('Access', () => {
    it('knows a design edit is unsaved work, and that opening one is not', async () => {
        const file = scratchCopy('AccessFormFixture.accdb', 'copies-access');
        const result = await psFile([
            '$ErrorActionPreference = "Continue"',
            unsavedWorkFunction('access'),
            '$app = New-Object -ComObject Access.Application',
            `$app.OpenCurrentDatabase(${q(file)})`,
            '$name = $app.CurrentProject.AllForms.Item(0).Name',
            '"untouched=" + (Test-XlideUnsavedWork $app)',
            // acDesign is 1, from src/analyzer/host/accessObjectModelData.ts.
            '$app.DoCmd.OpenForm($name, 1)',
            '"opened=" + (Test-XlideUnsavedWork $app)',
            '$app.Forms.Item($name).Caption = "edited by the check"',
            '"edited=" + (Test-XlideUnsavedWork $app)',
            // acForm 2, acSaveNo 2.
            '$app.DoCmd.Close(2, $name, 2)',
            '"discarded=" + (Test-XlideUnsavedWork $app)',
            '$app.CloseCurrentDatabase()',
            '$app.Quit()',
            '$app = $null',
            '[GC]::Collect()',
        ].join('\r\n'), 'access-unsaved.ps1');
        expect(fields(result.out)).toEqual({ untouched: 'False', opened: 'False', edited: 'True', discarded: 'False' });
    });
});

describe('Excel, in instances the check starts itself', () => {
    it('reports a read-only open that another copy still locks, and names that copy', async () => {
        const file = scratchCopy('FormFixture.xlsm', 'copies-excel-locked');
        const readOnlyOpen = buildHostOpenScript({ host: 'excel', filePath: file, attachToRunning: false, readOnly: true });
        const { buildLockHoldersScript } = await import('../../src/fileLockHolders');
        // The launcher's lines run in this file's scope, so its $excel and
        // $workbook are this check's own and are closed without attaching.
        const result = await psFile([
            '$ErrorActionPreference = "Continue"',
            'Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);\' -Name Owner -Namespace XlideLive',
            'function Get-OwnerPid($hwnd) { $owner = 0; [void][XlideLive.Owner]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$owner); return $owner }',
            'Write-Output "--- alone"',
            readOnlyOpen,
            'try { $workbook.Close($false); $excel.Quit() } catch { }',
            '$workbook = $null; $excel = $null; [GC]::Collect(); Start-Sleep -Seconds 2',
            'Write-Output "--- behind an editing copy"',
            '$a = New-Object -ComObject Excel.Application',
            '$a.DisplayAlerts = $false',
            `$wbA = $a.Workbooks.Open(${q(file)}, 0, $false)`,
            '"editorPid=" + (Get-OwnerPid $a.Hwnd)',
            readOnlyOpen,
            '& {',
            buildLockHoldersScript(file),
            '}',
            '$ErrorActionPreference = "Continue"',
            'try { $workbook.Close($false); $excel.Quit() } catch { }',
            'try { $wbA.Close($false); $a.Quit() } catch { }',
            '$workbook = $null; $excel = $null; $wbA = $null; $a = $null; [GC]::Collect()',
        ].join('\r\n'), 'excel-locked-elsewhere.ps1');
        const alone = result.out.slice(result.out.indexOf('--- alone'), result.out.indexOf('--- behind an editing copy'));
        const behind = result.out.slice(result.out.indexOf('--- behind an editing copy'));
        expect(hostOpenOutcome('excel', true, alone)).toEqual({});
        expect(hostOpenOutcome('excel', true, behind)).toEqual({ lockedElsewhere: true });
        const { parseLockHolders } = await import('../../src/fileLockHolders');
        const holders = parseLockHolders(behind);
        // The read-only copy holds nothing; the editing copy is the one named.
        expect(holders?.map((holder) => [holder.pid, holder.image])).toEqual([[Number(fields(behind).editorPid), 'EXCEL.EXE']]);
    });
});

// Gives the applications a moment to let go of their files between files.
afterAll(async () => {
    await sleep(500);
});
