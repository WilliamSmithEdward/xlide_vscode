// PowerShell command construction and result shapes.
//
// Quoting and the option/result types are plain data and live here. The
// spawner itself is a leaf the browser build swaps (see webBuild.js): a
// browser has no child_process, and nothing it could launch anyway.

/** Quotes a value for safe interpolation into a single-quoted PowerShell string. */
export function psSingleQuoted(value: string): string {
    return `'${value.replace(/'/g, "''")}'`;
}

export interface RunPowerShellOptions {
    /**
     * A script to run, one statement per line. It travels as -EncodedCommand,
     * so PowerShell parses it whole, the way it parses a script file: an
     * `else` or `catch` may start its own line, which it could not while
     * scripts were joined onto one line with "; ".
     */
    script?: string;
    /**
     * Arguments after the standard -NoProfile -ExecutionPolicy Bypass prefix,
     * for running a script file (`-File`). Ignored when `script` is given.
     */
    args?: string[];
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

export { runPowerShell } from './powershellNode';
