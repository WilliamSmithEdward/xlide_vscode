// Excel as a live preview: a workbook open read-only follows XLIDE's saves.
//
// A read-only workbook does not lock the file, so a module save goes through
// and the copy in Excel goes stale; the refresh closes and reopens it after
// each save (buildRefreshReadOnlyScript). The check drives the real refresh
// script against an Excel it starts itself - the script's one attach line
// becomes that instance - because the Excel someone is working in must never
// be attached to.

import * as fs from 'fs';
import * as path from 'path';
import { describe, expect, it } from 'vitest';
import { buildRefreshReadOnlyScript } from '../../src/officeWriteCoordinator';
import { readModule, writeModule } from '../../src/vba/projectService';
import { hostInstalled, psFile, q, scratchCopy, sleep } from './officeHarness';

const ATTACH_LINE = 'try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }';

/** The refresh script, run against the check's own Excel in `$own`. */
function refreshAgainstOwnExcel(file: string): string {
    const script = buildRefreshReadOnlyScript(file)!;
    expect(script.split('\n').filter((line) => line === ATTACH_LINE)).toHaveLength(1);
    return script.replace(ATTACH_LINE, '$app = $own');
}

/** `name=value` lines, as a record. */
function fields(lines: readonly string[]): Record<string, string> {
    return Object.fromEntries(lines.filter((line) => line.includes('=')).map((line) => {
        const at = line.indexOf('=');
        return [line.slice(0, at), line.slice(at + 1)];
    }));
}

async function waitForFile(file: string, timeoutMs: number): Promise<void> {
    const until = Date.now() + timeoutMs;
    while (!fs.existsSync(file)) {
        if (Date.now() > until) {
            throw new Error(`timed out waiting for ${file}`);
        }
        await sleep(100);
    }
}

/** What one save and refresh looked like: the script's report, and each Workbook_Open. */
interface FollowedSave {
    seen: Record<string, string>;
    err: string[];
    opened: string[];
}

/**
 * Opens a scratch copy read-only in an Excel the check starts, puts a place
 * in it, saves a module into the file, runs the refresh, and reports.
 */
async function followOneSave(): Promise<FollowedSave> {
    const file = scratchCopy('ShapesFixture.xlsm', 'live-preview');
    const dir = path.dirname(file);
    const marker = path.join(dir, 'opened.txt');
    const ready = path.join(dir, 'ready.txt');
    const go = path.join(dir, 'go.txt');
    // Workbook_Open leaves a line each time it runs.
    writeModule(file, 'ThisWorkbook', [
        'Private Sub Workbook_Open()',
        '    Dim f As Integer',
        '    f = FreeFile',
        `    Open "${marker}" For Append As #f`,
        '    Print #f, "opened"',
        '    Close #f',
        'End Sub',
        '',
    ].join('\r\n'), 'standard');

    const run = psFile([
        '$ErrorActionPreference = "Continue"',
        'Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);\' -Name Front -Namespace XlideLive',
        'function Get-Front { $hwnd = [XlideLive.Front]::GetForegroundWindow(); $owner = 0; [void][XlideLive.Front]::GetWindowThreadProcessId($hwnd, [ref]$owner); $name = ""; try { $name = (Get-Process -Id $owner -ErrorAction Stop).ProcessName } catch { }; return ([string]$hwnd + ":" + $name) }',
        '$own = New-Object -ComObject Excel.Application',
        '$own.Visible = $true',
        `$wb = $own.Workbooks.Open(${q(file)}, 0, $true)`,
        '$sheet = $wb.Worksheets.Item(2)',
        '$sheet.Activate()',
        '$sheet.Range("B5:C7").Select()',
        '$sheet.Range("C6").Activate()',
        '$own.ActiveWindow.ScrollRow = 20',
        '$own.ActiveWindow.ScrollColumn = 3',
        'function Write-Place($label, $book) {',
        '  $window = $book.Windows.Item(1)',
        '  "$label.sheet=" + $window.ActiveSheet.Name',
        '  "$label.selection=" + $window.RangeSelection.Address()',
        '  "$label.activeCell=" + $window.ActiveCell.Address()',
        '  "$label.scroll=" + $window.ScrollRow + "," + $window.ScrollColumn',
        '  "$label.readOnly=" + $book.ReadOnly',
        '}',
        'Write-Place "before" $wb',
        `Set-Content -Path ${q(ready)} -Value "ready"`,
        `while (-not (Test-Path ${q(go)})) { Start-Sleep -Milliseconds 100 }`,
        '"front.before=" + (Get-Front)',
        refreshAgainstOwnExcel(file),
        '"front.after=" + (Get-Front)',
        'Write-Place "after" $file',
        '"events=" + $own.EnableEvents',
        '$text = ""',
        'try { $text = $file.VBProject.VBComponents.Item("XlideProbe").CodeModule.Lines(1, 5) } catch { $text = "unreadable: " + $_ }',
        '"probe=" + ($text -replace "[\\r\\n]+", " ")',
        'try { $file.Close($false) } catch { }',
        'try { $wb.Close($false) } catch { }',
        '$own.Quit()',
        '$sheet = $null; $wb = $null; $file = $null; $app = $null; $own = $null',
        '[GC]::Collect()',
    ].join('\r\n'), 'live-preview.ps1');

    await waitForFile(ready, 120_000);
    // The save the preview follows, written while Excel shows the old copy.
    writeModule(file, 'XlideProbe', 'Public Sub Probe()\r\n    \' live\r\nEnd Sub\r\n', 'standard');
    expect(readModule(file, 'XlideProbe', false).source).toContain('live');
    fs.writeFileSync(go, 'go');

    const result = await run;
    return {
        seen: fields(result.out),
        err: result.err,
        opened: fs.readFileSync(marker, 'utf8').trim().split(/\r?\n/),
    };
}

describe.runIf(await hostInstalled('excel'))('Excel as a live preview', () => {
    it('follows a save, keeping the sheet, selection and scroll, without running Workbook_Open again', async () => {
        // Whether the check's own Excel takes the foreground when it is shown
        // is Windows' call, and "stays behind whatever was in front" says
        // nothing when that was Excel. It happened in two runs of fifteen, so
        // the scenario runs again, a bounded number of times, until it was not.
        let followed = await followOneSave();
        for (let attempt = 1; attempt < 3 && /:EXCEL$/i.test(followed.seen['front.before'] ?? ''); attempt++) {
            followed = await followOneSave();
        }
        const { seen, err, opened } = followed;
        expect(seen['XLIDE_REFRESH|refreshed'], err.join('\n')).toBe('True');
        expect(seen['before.sheet']).toBe('Sheet2');
        for (const part of ['sheet', 'selection', 'activeCell', 'scroll', 'readOnly']) {
            expect(seen[`after.${part}`], part).toBe(seen[`before.${part}`]);
        }
        expect(seen['after.selection']).toBe('$B$5:$C$7');
        expect(seen['after.activeCell']).toBe('$C$6');
        expect(seen['after.readOnly']).toBe('True');
        // The reopened copy shows the save.
        expect(seen.probe).toContain('live');
        // Workbook_Open ran when the workbook was first opened, and not for
        // the refresh; and events are back on afterwards.
        expect(opened).toEqual(['opened']);
        expect(seen.events).toBe('True');
        // Whatever was in front stays there: the window, as `hwnd:process`.
        expect(seen['front.before']).not.toMatch(/:EXCEL$/i);
        expect(seen['front.after']).toBe(seen['front.before']);
    });
});
