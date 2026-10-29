import * as cp from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { NOTICE_LOG, SCRATCH, SECURITY_NOTICE_SCRIPT, sleep } from './officeHarness';

/**
 * Starts each run with an empty scratch folder, and answers any macro
 * Security Notice about a scratch file until the run ends.
 */
export default async function setup(): Promise<() => void> {
    try {
        fs.rmSync(SCRATCH, { recursive: true, force: true });
    } catch {
        // Something still holds a file from the last run; the checks each
        // start from a fresh subfolder anyway.
    }
    fs.mkdirSync(SCRATCH, { recursive: true });
    // rmSync above can fail on a held fixture. Installation probes still
    // must never survive into a new run.
    for (const host of ['excel', 'word', 'powerpoint', 'access']) {
        fs.rmSync(path.join(SCRATCH, `installed-${host}.txt`), { force: true });
    }
    fs.rmSync(NOTICE_LOG, { force: true });
    fs.rmSync(`${NOTICE_LOG}.ready`, { force: true });
    if (process.platform !== 'win32') {
        return () => undefined;
    }
    return startNoticeAccepter();
}

/**
 * Office can stop an open with a modal "Microsoft <Application> Security
 * Notice", whether a check's own script or XLIDE's opened the file, and the
 * call that opened it waits until it is answered. One PowerShell for the
 * whole run answers each one that names a file under the scratch folder, in
 * any Office application, and logs it to NOTICE_LOG; a notice about anything
 * else is left alone. When no notice comes, nothing waits for one.
 */
async function startNoticeAccepter(): Promise<() => void> {
    const script = path.join(SCRATCH, 'security-notice-accepter.ps1');
    fs.writeFileSync(script, [
        'param([int]$ParentPid, [string]$Scratch, [string]$LogPath, [string]$Helper)',
        '$ErrorActionPreference = "Stop"',
        '. $Helper',
        // Ready only once the type has loaded: a failed load ends the script first.
        'if (-not (\'XlideSecurityNotice\' -as [type])) { exit 1 }',
        'Set-Content -LiteralPath ($LogPath + ".ready") -Value "ready"',
        'while (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue) {',
        '    foreach ($line in [XlideSecurityNotice]::Poll(0, $Scratch)) { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 }',
        '    Start-Sleep -Milliseconds 100',
        '}',
    ].join('\r\n'));
    // Its errors go beside the log, for a run that has to say why it stopped.
    const errors = fs.openSync(`${NOTICE_LOG}.errors`, 'w');
    const child = cp.spawn('powershell.exe', [
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script,
        '-ParentPid', String(process.pid), '-Scratch', SCRATCH, '-LogPath', NOTICE_LOG, '-Helper', SECURITY_NOTICE_SCRIPT,
    ], { stdio: ['ignore', errors, errors], windowsHide: true });
    fs.closeSync(errors);
    let exited = false;
    child.on('exit', () => { exited = true; });
    const until = Date.now() + 60_000;
    while (!fs.existsSync(`${NOTICE_LOG}.ready`) && !exited && Date.now() < until) {
        await sleep(50);
    }
    if (!fs.existsSync(`${NOTICE_LOG}.ready`)) {
        child.kill();
        const said = fs.readFileSync(`${NOTICE_LOG}.errors`, 'utf8').trim();
        throw new Error(`The Security Notice accepter did not start, so a notice would leave a check waiting. ${said}`.trim());
    }
    return () => { child.kill(); };
}
