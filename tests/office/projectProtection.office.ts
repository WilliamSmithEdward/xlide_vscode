// Excel itself must recognize the permanent fixture's viewing lock.
import { describe, expect, it } from 'vitest';
import * as fs from 'fs';
import { hostInstalled, ps, q, scratchCopy } from './officeHarness';

describe('Protected VBA password fixture in Excel', () => {
    it('Excel reports the permanent fixture as locked for viewing', async context => {
        if (!await hostInstalled('excel')) { context.skip(); }
        const file = scratchCopy('PasswordProtectedFixture.xlsm', 'password-protection');
        const before = fs.readFileSync(file);
        const result = await ps(`
$excel = $null; $book = $null; $owned = $false
$previousPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class XlideFixtureWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); }'
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelPid = 0
    [void][XlideFixtureWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    if ($excelPid -eq 0 -or $previousPids -contains $excelPid) { "owned-instance=no"; return }
    $owned = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open(${q(file)}, 0, $true)
    $project = $book.VBProject
    "protection=$([int]$project.Protection)"
    # Accessing VBComponents on a locked project can show a modal password
    # dialog. Protection=1 is Office's own authoritative viewing-lock state.
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($excel) { if ($owned) { $excel.Quit() }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
`);
        expect(result.code, JSON.stringify(result)).toBe(0);
        if (result.out.includes('owned-instance=no')) { context.skip(); }
        expect(result.out).toContain('protection=1');
        expect(fs.readFileSync(file)).toEqual(before);
    });
});
