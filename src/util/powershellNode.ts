// The PowerShell spawner. The browser build never reaches this module -
// webBuild.js aliases it to powershellWeb.ts - which is what keeps
// child_process out of the web bundle.

import * as cp from 'child_process';
import type {
    PowerShellRun,
    RunPowerShellOptions,
    RunPowerShellResult,
} from './powershell';

/**
 * Kill the spawned powershell.exe and, on Windows, its whole process tree.
 * Node's ChildProcess.kill() TerminateProcess-es only powershell.exe itself, not
 * processes it spawned, so use `taskkill /PID <pid> /T /F` to tear down the tree
 * (mirrors killOwnedHost in vbaTestHostSession). Falls back to a bare kill().
 */
function killProcessTree(child: cp.ChildProcess): void {
    if (process.platform === 'win32' && child.pid !== undefined) {
        try {
            const killer = cp.spawn(
                'taskkill.exe',
                ['/PID', String(child.pid), '/T', '/F'],
                { windowsHide: true },
            );
            killer.on('error', () => child.kill());
            return;
        } catch {
            /* fall through to bare kill */
        }
    }
    child.kill();
}

/**
 * Single PowerShell child-process launcher shared by the Excel launcher, the
 * COM availability probe, and the VBA test host: spawn with a hidden window,
 * buffer stdout/stderr into trimmed non-empty lines split on \r?\n (flushed
 * on close), and optionally kill on timeout.
 */
/**
 * The arguments that run `script`: -EncodedCommand with the script as
 * base64 UTF-16LE, which is what PowerShell decodes it from. The script
 * arrives with its own line breaks and is parsed whole.
 *
 * Windows PowerShell 5.1 writes redirected stdout/stderr in the OEM/ANSI code
 * page, so non-ASCII text (a localized COM error message) would mojibake when
 * Node decodes the bytes as UTF-8. The script's first line switches output to
 * UTF-8 so the two ends agree; the -File host (run-vba-tests.ps1) sets it
 * itself.
 */
export function encodedCommandArgs(script: string): string[] {
    const withUtf8Output = `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n${script}`;
    return ['-EncodedCommand', Buffer.from(withUtf8Output, 'utf16le').toString('base64')];
}

export function runPowerShell(options: RunPowerShellOptions): PowerShellRun {
    const child = cp.spawn('powershell.exe', [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        ...(options.script !== undefined ? encodedCommandArgs(options.script) : options.args ?? []),
    ], { windowsHide: options.windowsHide ?? true });

    const stdoutLines: string[] = [];
    const stderrLines: string[] = [];
    let settled = false;
    let timedOut = false;
    let timer: ReturnType<typeof setTimeout> | undefined;

    const result = new Promise<RunPowerShellResult>((resolve) => {
        const finish = (partial: Pick<RunPowerShellResult, 'code' | 'signal' | 'spawnError'>) => {
            if (settled) { return; }
            settled = true;
            if (timer) { clearTimeout(timer); }
            resolve({ ...partial, timedOut, stdoutLines, stderrLines });
        };

        const lineBuffer = (sink: string[], onLine?: (line: string) => void) => {
            let buffered = '';
            const push = (line: string) => {
                const trimmed = line.trimEnd();
                if (!trimmed) { return; }
                sink.push(trimmed);
                onLine?.(trimmed);
            };
            return {
                append: (chunk: Buffer) => {
                    buffered += chunk.toString();
                    const lines = buffered.split(/\r?\n/);
                    buffered = lines.pop() ?? '';
                    for (const line of lines) { push(line); }
                },
                flush: () => {
                    if (buffered) {
                        push(buffered);
                        buffered = '';
                    }
                },
            };
        };

        const stdout = lineBuffer(stdoutLines, options.onStdoutLine);
        const stderr = lineBuffer(stderrLines, options.onStderrLine);

        if (options.timeoutMs !== undefined && options.timeoutMs > 0) {
            timer = setTimeout(() => {
                timedOut = true;
                killProcessTree(child);
                // Flush buffered partial lines before settling, the same way the
                // error/close handlers do, so a sentinel/diagnostic emitted just
                // before the timeout is not lost from the resolved result.
                stdout.flush();
                stderr.flush();
                finish({ code: null, signal: null });
            }, options.timeoutMs);
        }

        child.on('spawn', () => options.onSpawn?.(child.pid ?? undefined));
        child.stdout?.on('data', (chunk: Buffer) => stdout.append(chunk));
        child.stderr?.on('data', (chunk: Buffer) => stderr.append(chunk));
        child.on('error', (err) => {
            stdout.flush();
            stderr.flush();
            finish({ code: null, signal: null, spawnError: err });
        });
        // 'close' rather than 'exit' so buffered stdio is fully drained first.
        child.on('close', (code, signal) => {
            stdout.flush();
            stderr.flush();
            finish({ code, signal });
        });
    });

    return {
        result,
        kill: () => { killProcessTree(child); },
    };
}
