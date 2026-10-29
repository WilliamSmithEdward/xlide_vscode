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

import * as cp from 'child_process';
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

/** The oracle's Security Notice accepter, which the suite shares (officeGlobalSetup.ts). */
export const SECURITY_NOTICE_SCRIPT = fileURLToPath(new URL('../../syntax_corpus/oracle/office_security_notice.ps1', import.meta.url));

/** Each Security Notice the run answered, or left alone, one per line. */
export const NOTICE_LOG = path.join(SCRATCH, 'security-notices.log');

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
    // Enumerate all rows: filtered tasklist emits localized prose for no match.
    // Never interpret a failed or malformed query as permission to attach.
    const output = await new Promise<string>((resolve, reject) => {
        cp.execFile('tasklist.exe', ['/FO', 'CSV', '/NH'], { windowsHide: true, maxBuffer: 4 * 1024 * 1024 }, (error, stdout) => {
            if (error) { reject(error); } else { resolve(stdout); }
        });
    });
    const rows = output.trim().split(/\r?\n/).map((line) => /^"([^"\r\n]+)","(\d+)",/.exec(line));
    if (rows.length === 0 || rows.some((row) => !row)) {
        throw new Error('Cannot establish Office process count from tasklist');
    }
    const image = `${OFFICE_HOST_APPS[host].processName}.exe`.toLowerCase();
    return rows.filter((row) => row![1].toLowerCase() === image).length;
}

/** Whether the host is installed for automation on this machine. */
export async function hostInstalled(host: OfficeHostApp): Promise<boolean> {
    // Vitest isolates test files. A scratch marker shares this immutable probe
    // across them; global setup removes it before each run. Never cache liveness.
    const marker = path.join(SCRATCH, `installed-${host}.txt`);
    if (fs.existsSync(marker)) {
        const cached = fs.readFileSync(marker, 'utf8');
        if (cached === 'yes' || cached === 'no') { return cached === 'yes'; }
    }
    const result = await ps(`if ([type]::GetTypeFromProgID("${OFFICE_HOST_APPS[host].progId}")) { "yes" } else { "no" }`);
    const answer = result.out[0];
    if (result.code !== 0 || (answer !== 'yes' && answer !== 'no')) {
        throw new Error(`Cannot establish whether ${host} is installed: ${result.err.join(' ')}`);
    }
    fs.mkdirSync(SCRATCH, { recursive: true });
    fs.writeFileSync(marker, answer);
    return answer === 'yes';
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
        await sleep(100);
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
 * A macro Security Notice is not one: closing it would disable the file's
 * code, so it is left to the run's notice accepter (officeGlobalSetup.ts).
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
    '            string title = Text(hWnd);',
    '            if (title.StartsWith("Microsoft ") && title.EndsWith(" Security Notice")) { return true; }',
    '            var parts = new List<string>();',
    '            EnumChildWindows(hWnd, (child, l2) => { string t = Text(child); if (t.Length > 0) { parts.Add(t); } return true; }, IntPtr.Zero);',
    '            found.Add(cls + " | " + title + " | " + string.Join(" / ", parts));',
    '            PostMessage(hWnd, 0x0010, IntPtr.Zero, IntPtr.Zero);',
    '            return true;',
    '        }, IntPtr.Zero);',
    '        return found;',
    '    }',
    '}',
    "'@",
    // Ready only once the type has loaded; a failed compile ends the watcher.
    'if (-not (\'XlideDialogs\' -as [type])) { exit 1 }',
    'Set-Content -LiteralPath ($LogPath + ".ready") -Value "ready"',
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
        `Remove-Item -LiteralPath ${q(log + '.ready')}, ${q(log + '.unready')} -Force -ErrorAction SilentlyContinue`,
        `$watch = Start-Process powershell.exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ${q(watcher)}, '-TargetPid', $owner, '-LogPath', ${q(log)}) -WindowStyle Hidden -PassThru`,
        // Carries on if the watcher is not ready in time, so the script still
        // reaches its own Quit, but records that; dialogsSeen then fails the
        // check, even if the watcher gets ready later.
        '$watchUntil = [DateTime]::UtcNow.AddSeconds(30)',
        `while (-not (Test-Path -LiteralPath ${q(log + '.ready')}) -and -not $watch.HasExited -and [DateTime]::UtcNow -lt $watchUntil) {`,
        '  Start-Sleep -Milliseconds 25',
        '}',
        `if (-not (Test-Path -LiteralPath ${q(log + '.ready')})) { Set-Content -LiteralPath ${q(log + '.unready')} -Value 'unready' }`,
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
    if (fs.existsSync(`${log}.unready`) || !fs.existsSync(`${log}.ready`)) {
        throw new Error(`The dialog watcher for ${log} was not ready before the open, so dialogs were not watched.`);
    }
    return fs.existsSync(log) ? fs.readFileSync(log, 'utf8').split(/\r?\n/).filter(Boolean) : [];
}
