# Run by .github/workflows/office.yml once Office is set up. Starts each of
# the four oracle hosts visibly with a new file and watches it for 20
# seconds: with the workflow's registry values in place, none should show a
# dialog. Any dialog fails the step, named with when it appeared, and is
# asked to close so the next app starts cleanly. An app that cannot start or
# open its file fails it too. Every app has its turn before the step fails.
$ErrorActionPreference = 'Stop'
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class StartDialogs {
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr w, IntPtr l);
    // The visible dialogs of one process, each asked to close (WM_CLOSE).
    public static List<string> CloseDialogs(uint pid) {
        var found = new List<string>();
        EnumWindows((hWnd, l) => {
            uint owner; GetWindowThreadProcessId(hWnd, out owner);
            if (owner != pid || !IsWindowVisible(hWnd)) { return true; }
            var c = new StringBuilder(256); GetClassName(hWnd, c, 256);
            if (c.ToString() != "#32770" && c.ToString() != "NUIDialog") { return true; }
            var t = new StringBuilder(1024); GetWindowText(hWnd, t, 1024);
            found.Add(c + " | " + t);
            PostMessage(hWnd, 0x0010, IntPtr.Zero, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@
$apps = @(
    @{ ProgId = 'Word.Application'; Image = 'WINWORD'; Open = { param($a) [void]$a.Documents.Add() } },
    @{ ProgId = 'PowerPoint.Application'; Image = 'POWERPNT'; Open = { param($a) [void]$a.Presentations.Add() } },
    @{ ProgId = 'Excel.Application'; Image = 'EXCEL'; Open = { param($a) [void]$a.Workbooks.Add() } },
    @{ ProgId = 'Access.Application'; Image = 'MSACCESS'; Open = { param($a) $a.NewCurrentDatabase((Join-Path $env:RUNNER_TEMP 'start-check.accdb')) } }
)
$failures = @()
foreach ($app in $apps) {
    $o = $null
    try {
        $o = New-Object -ComObject $app.ProgId
        $o.Visible = $true
        & $app.Open $o
        $proc = Get-Process -Name $app.Image | Select-Object -First 1
        $seen = [ordered]@{}
        $started = Get-Date
        while (((Get-Date) - $started).TotalSeconds -lt 20) {
            foreach ($dialog in [StartDialogs]::CloseDialogs([uint32]$proc.Id)) {
                if (-not $seen.Contains($dialog)) { $seen[$dialog] = ((Get-Date) - $started).TotalSeconds }
            }
            Start-Sleep -Milliseconds 250
        }
        $found = $seen.Keys | ForEach-Object { "$_ (at $('{0:N1}' -f $seen[$_]) s)" }
        Write-Host "$($app.ProgId): $(if ($found) { $found -join '; ' } else { 'no dialog' })"
        if ($found) { $failures += "$($app.ProgId) showed: $($found -join '; ')" }
    } catch {
        $failures += "$($app.ProgId): $($_.Exception.Message)"
    } finally {
        if ($o) {
            try {
                if ($app.Image -eq 'EXCEL') { $o.DisplayAlerts = $false }
                if ($app.Image -eq 'MSACCESS') { $o.CloseCurrentDatabase() }
                $o.Quit()
            } catch { Write-Host "$($app.ProgId): quit failed: $($_.Exception.Message)" }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)
        }
        Start-Sleep -Seconds 2
        Get-Process -Name $app.Image -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}
Remove-Item (Join-Path $env:RUNNER_TEMP 'start-check.accdb') -Force -ErrorAction SilentlyContinue
if ($failures) { throw "Office showed a dialog on start, or did not start:`n$($failures -join "`n")" }
