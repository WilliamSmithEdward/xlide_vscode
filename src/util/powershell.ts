import * as cp from 'child_process';

/** Quotes a value for safe interpolation into a single-quoted PowerShell string. */
export function psSingleQuoted(value: string): string {
    return `'${value.replace(/'/g, "''")}'`;
}

export interface RunPowerShellOptions {
    /** Arguments appended after the standard -NoProfile -ExecutionPolicy Bypass prefix. */
    args: string[];
    /** When set (> 0), the process is killed and the result resolves with timedOut. */
    timeoutMs?: number;
    /** Hide the console window; defaults to true. */
    windowsHide?: boolean;
    onSpawn?: (pid: number | undefined) => void;
    onStdoutLine?: (line: string) => void;
    onStderrLine?: (line: string) => void;
}

export interface RunPowerShellResult {
    code: number | null;
    signal: NodeJS.Signals | null;
    timedOut: boolean;
    /** Set when powershell.exe could not be started. */
    spawnError?: Error;
    stdoutLines: string[];
    stderrLines: string[];
}

export interface PowerShellRun {
    result: Promise<RunPowerShellResult>;
    /** Kills the process early; the result still resolves via the close handler. */
    kill(): void;
}

/**
 * Single PowerShell child-process launcher shared by the Excel launcher, the
 * COM availability probe, and the VBA test host: spawn with a hidden window,
 * buffer stdout/stderr into trimmed non-empty lines split on \r?\n (flushed
 * on close), and optionally kill on timeout.
 */
export function runPowerShell(options: RunPowerShellOptions): PowerShellRun {
    const child = cp.spawn('powershell.exe', [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        ...options.args,
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
                child.kill();
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
        kill: () => { child.kill(); },
    };
}
