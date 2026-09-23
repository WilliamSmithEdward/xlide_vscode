// What every live Office check shares: scratch files, the PowerShell the
// product itself uses, and the rules that keep the suite away from anything
// that is not its own.
//
// The rules, which no check relaxes:
//
//   - Only scratch copies of tests/fixtures/binaries are opened, and only
//     files under this run's scratch folder are ever closed.
//   - A check that drives an application through XLIDE's own attach paths -
//     which find the RUNNING application - runs only when no instance of that
//     application was running when the check began, so whatever it attaches
//     to is an instance the suite started. Otherwise it is skipped.
//   - Excel checks that must run while someone's Excel is open start their
//     own instances inside a PowerShell file and never attach.
//   - An application is quit only when nothing but the suite's files was open
//     in it, and no process is ever killed.

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { runPowerShell } from '../../src/util/powershell';
import { OFFICE_HOST_APPS, type OfficeHostApp } from '../../src/officeHostApps';

export const FIXTURES = fileURLToPath(new URL('../fixtures/binaries/', import.meta.url));

/**
 * The suite's scratch folder; everything it opens lives under it. Emptied at
 * the start of each run (officeGlobalSetup.ts) and left afterwards, so the
 * files a failed check worked on can still be looked at.
 */
export const SCRATCH = path.join(os.tmpdir(), 'xlide-office-live');

/** A fresh copy of a fixture in its own scratch folder. */
export function scratchCopy(fixture: string, label: string, name = fixture): string {
    const dir = path.join(SCRATCH, label);
    fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    const target = path.join(dir, name);
    fs.copyFileSync(path.join(FIXTURES, fixture), target);
    return target;
}

export interface PsResult {
    out: string[];
    err: string[];
    code: number | null;
}

/** Runs a script the way the product runs its own: whole, as -EncodedCommand. */
export async function ps(script: string, timeoutMs = 120_000): Promise<PsResult> {
    const result = await runPowerShell({ script, timeoutMs }).result;
    return { out: result.stdoutLines, err: result.stderrLines, code: result.code };
}

/**
 * Runs a script from a file in the scratch folder: for a check that starts
 * its own instance and has to keep hold of it across what it runs.
 */
export async function psFile(body: string, name: string, timeoutMs = 300_000): Promise<PsResult> {
    fs.mkdirSync(SCRATCH, { recursive: true });
    const file = path.join(SCRATCH, name);
    fs.writeFileSync(file, body);
    const result = await runPowerShell({ args: ['-File', file], timeoutMs }).result;
    return { out: result.stdoutLines, err: result.stderrLines, code: result.code };
}

/** A PowerShell single-quoted literal. */
export const q = (value: string): string => `'${value.replace(/'/g, "''")}'`;

/** How many processes of the host's application are running. */
export async function runningCount(host: OfficeHostApp): Promise<number> {
    const result = await ps(`@(Get-Process -Name ${OFFICE_HOST_APPS[host].processName} -ErrorAction SilentlyContinue).Count`);
    return Number(result.out[0] ?? '0');
}

/** Whether the host is installed for automation on this machine. */
export async function hostInstalled(host: OfficeHostApp): Promise<boolean> {
    const result = await ps(`if ([type]::GetTypeFromProgID("${OFFICE_HOST_APPS[host].progId}")) { "yes" } else { "no" }`);
    return result.out[0] === 'yes';
}

/**
 * Whether the suite may drive this host through XLIDE's attach paths: it is
 * installed and not running at all, so any instance attached to is ours.
 */
export async function hostIsFree(host: OfficeHostApp): Promise<boolean> {
    return await hostInstalled(host) && await runningCount(host) === 0;
}

/**
 * Closes whatever the suite left open in a host it started - only files
 * under the scratch folder - and quits the application when nothing else is
 * open in it. Then waits for the process to go, so the next check finds the
 * host free again. Never kills.
 */
export async function releaseHost(host: Exclude<OfficeHostApp, 'access'>): Promise<void> {
    const collection = { excel: 'Workbooks', word: 'Documents', powerpoint: 'Presentations' }[host];
    const close = { excel: '$f.Close($false)', word: '$f.Close(0)', powerpoint: '$f.Close()' }[host];
    const quit = { excel: '$app.Quit()', word: '$app.Quit(0)', powerpoint: '$app.Quit()' }[host];
    await ps([
        `$root = ${q(SCRATCH)}`,
        '$app = $null',
        `try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("${OFFICE_HOST_APPS[host].progId}") } catch { }`,
        'if ($app) {',
        `  foreach ($f in @($app.${collection})) { if ($f.FullName -like ($root + "*")) { try { ${close} } catch { } } }`,
        `  if ($app.${collection}.Count -eq 0) { ${quit} }`,
        '}',
        '$app = $null',
        '[GC]::Collect()',
    ].join('\n'));
    await waitForExit(host);
}

/** Waits for every process of the host to end on its own. */
export async function waitForExit(host: OfficeHostApp, timeoutMs = 30_000): Promise<boolean> {
    const until = Date.now() + timeoutMs;
    while (Date.now() < until) {
        if (await runningCount(host) === 0) {
            return true;
        }
        await sleep(500);
    }
    return false;
}

export const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

/** The window in front, and the process that owns it. */
/**
 * The window in front, once it is not the application under test. Whether a
 * window the check opened takes the foreground is Windows' call; it did in
 * three of eight full runs of this suite. A copy that is itself in front is
 * the one a save replaces, so "the reopened copy stays behind" cannot be
 * judged from there.
 * `settle` runs the operation under test once more, which leaves that window
 * behind if the product does, and the foreground is read again. A product
 * that brings it forward keeps it in front, and fails the check that follows.
 */
export async function foregroundAwayFrom(
    processName: string,
    settle: () => Promise<unknown>,
): Promise<{ hwnd: string; process: string }> {
    let front = await foreground();
    for (let attempt = 1; attempt < 3 && front.process.toLowerCase() === processName.toLowerCase(); attempt++) {
        await settle();
        front = await foreground();
    }
    return front;
}

export async function foreground(): Promise<{ hwnd: string; process: string }> {
    const result = await ps([
        'Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow(); [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);\' -Name Foreground -Namespace XlideLive',
        '$hwnd = [XlideLive.Foreground]::GetForegroundWindow()',
        '$owner = 0',
        '[void][XlideLive.Foreground]::GetWindowThreadProcessId($hwnd, [ref]$owner)',
        '$name = ""',
        'try { $name = (Get-Process -Id $owner -ErrorAction Stop).ProcessName } catch { }',
        '[Console]::Out.WriteLine([string]$hwnd + "|" + $name)',
    ].join('\n'));
    const [hwnd = '', process = ''] = (result.out[0] ?? '').split('|');
    return { hwnd, process };
}

/**
 * Records every dialog a process raises, and closes it, until the process
 * ends: a repair prompt is one of these. Run in its own PowerShell beside the
 * application, since an open that raises one does not return until it closes.
 */
export const WATCHER = [
    'param([int]$TargetPid, [string]$LogPath, [int]$Seconds = 180)',
    "Add-Type -TypeDefinition @'",
    'using System;',
    'using System.Collections.Generic;',
    'using System.Runtime.InteropServices;',
    'using System.Text;',
    'public static class XlideDialogs {',
    '    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);',
    '    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr lParam);',
    '    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr lParam);',
    '    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);',
    '    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);',
    '    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int max);',
    '    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);',
    '    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr w, IntPtr l);',
    '    static string Text(IntPtr hWnd) { var b = new StringBuilder(1024); GetWindowText(hWnd, b, 1024); return b.ToString(); }',
    '    static string Class(IntPtr hWnd) { var b = new StringBuilder(256); GetClassName(hWnd, b, 256); return b.ToString(); }',
    '    public static List<string> Dialogs(uint pid) {',
    '        var found = new List<string>();',
    '        EnumWindows((hWnd, l) => {',
    '            uint owner; GetWindowThreadProcessId(hWnd, out owner);',
    '            if (owner != pid || !IsWindowVisible(hWnd)) { return true; }',
    '            string cls = Class(hWnd);',
    '            if (cls != "#32770" && cls != "NUIDialog") { return true; }',
    '            var parts = new List<string>();',
    '            EnumChildWindows(hWnd, (child, l2) => { string t = Text(child); if (t.Length > 0) { parts.Add(t); } return true; }, IntPtr.Zero);',
    '            found.Add(cls + " | " + Text(hWnd) + " | " + string.Join(" / ", parts));',
    '            PostMessage(hWnd, 0x0010, IntPtr.Zero, IntPtr.Zero);',
    '            return true;',
    '        }, IntPtr.Zero);',
    '        return found;',
    '    }',
    '}',
    "'@",
    '$until = (Get-Date).AddSeconds($Seconds)',
    'while ((Get-Date) -lt $until -and (Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)) {',
    '    foreach ($line in [XlideDialogs]::Dialogs([uint32]$TargetPid)) { Add-Content -Path $LogPath -Value $line }',
    '    Start-Sleep -Milliseconds 250',
    '}',
].join('\r\n');

/** PowerShell starting the watcher on the process that owns window `hwnd`. */
export function watchLines(hwndExpression: string, log: string): string[] {
    const watcher = path.join(SCRATCH, 'dialog-watcher.ps1');
    fs.mkdirSync(SCRATCH, { recursive: true });
    fs.writeFileSync(watcher, WATCHER);
    return [
        'Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);\' -Name Owner -Namespace XlideLive',
        '$owner = 0',
        `[void][XlideLive.Owner]::GetWindowThreadProcessId([IntPtr]${hwndExpression}, [ref]$owner)`,
        `Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ${q(watcher)}, '-TargetPid', $owner, '-LogPath', ${q(log)}) -WindowStyle Hidden | Out-Null`,
        'Start-Sleep -Milliseconds 1500',
    ];
}

/** `name=value` lines, as a record. */
export function fields(lines: readonly string[]): Record<string, string> {
    return Object.fromEntries(lines.filter((line) => line.includes('=')).map((line) => {
        const at = line.indexOf('=');
        return [line.slice(0, at), line.slice(at + 1)];
    }));
}

/** The dialogs the watcher saw, one per line. */
export function dialogsSeen(log: string): string[] {
    return fs.existsSync(log) ? fs.readFileSync(log, 'utf8').split(/\r?\n/).filter(Boolean) : [];
}
