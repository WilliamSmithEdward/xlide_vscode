import * as path from 'path';
import { psSingleQuoted, runPowerShell } from './util/powershell';

export type ExcelMacroFailureCode = 'REOPEN_BLOCKED' | 'REOPEN_FAILED' | 'RUN_FAILED' | 'UNKNOWN';

/** Typed macro-run failure decoded from the script's CODE|message sentinels. */
export class ExcelMacroError extends Error {
    constructor(message: string, readonly code: ExcelMacroFailureCode) {
        super(message);
        this.name = 'ExcelMacroError';
    }
}

const MACRO_ERROR_SENTINEL = 'XLIDE_MACRO_ERROR|';

export type ExcelLaunchMode =
    | { kind: 'open'; readOnly: boolean }
    | { kind: 'macroReadOnly'; macroName: string };

export interface ExcelLaunchScriptOptions {
    filePath: string;
    attachToRunning: boolean;
    mode: ExcelLaunchMode;
}

// Attach to a running Excel instance (optional), or start one, then look for
// an already-open copy of the target workbook.
function attachLines(filePath: string, attachToRunning: boolean): string[] {
    return [
        `$targetPath = ${psSingleQuoted(filePath)}`,
        // win32 semantics regardless of host platform: this script always
        // drives Windows Excel against a Windows path.
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        '$excel = $null',
        '$workbook = $null',
        `$attachToRunning = ${attachToRunning ? '$true' : '$false'}`,
        'if ($attachToRunning) {',
        '  try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
        '}',
        'if (-not $excel) {',
        '  $excel = New-Object -ComObject Excel.Application',
        '}',
        '$excel.Visible = $true',
        'foreach ($wb in @($excel.Workbooks)) {',
        '  if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) { $workbook = $wb; break }',
        '}',
    ];
}

// PowerShell helper that retries a COM call while Excel is busy. Excel reports a
// busy state (a modal dialog such as a MsgBox left open from a previous run, or a
// transient busy moment) as RPC_E_CALL_REJECTED (HResult -2147418111 / 0x80010001)
// or RPC_E_SERVERCALL_RETRYLATER (-2147417846 / 0x8001010A). ~3s of retries rides
// out a transient busy and gives the user a moment to dismiss a dialog.
const COM_RETRY_HELPER =
    'function Invoke-XlideCom($Action) { for ($__i = 0; $__i -le 12; $__i++) { try { return (& $Action) } catch { if (($_.Exception.HResult -eq -2147418111 -or $_.Exception.HResult -eq -2147417846 -or $_.Exception.InnerException.HResult -eq -2147418111 -or $_.Exception.InnerException.HResult -eq -2147417846 -or $_.Exception.Message -match "rejected by callee|RETRYLATER|0x80010001|0x8001010A") -and $__i -lt 12) { Start-Sleep -Milliseconds 250; continue } else { throw } } } }';

// Activate the workbook and bring the Excel window to the foreground. Activate is
// a COM call, so it is best-effort: a busy Excel must not fail the launch/run.
const FOREGROUND_LINES: readonly string[] = [
    'try { $workbook.Activate() } catch { }',
    'try { Add-Type -MemberDefinition \'[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);\' -Name XlideWin32 -Namespace XlideHelper } catch { }',
    '[XlideHelper.XlideWin32]::ShowWindow([IntPtr]$excel.Hwnd, 9)',
    '[XlideHelper.XlideWin32]::SetForegroundWindow([IntPtr]$excel.Hwnd)',
];

/**
 * Single parameterized attach/open/foreground script for the Windows Excel
 * launch paths: plain open (reuses an already-open workbook) and read-only
 * macro runs (reopen read-only, foreground, run the macro, report failures
 * through the XLIDE_MACRO_ERROR stderr sentinel).
 */
export function buildExcelLaunchScript(options: ExcelLaunchScriptOptions): string {
    const { filePath, attachToRunning, mode } = options;
    if (mode.kind === 'open') {
        return [
            '$ErrorActionPreference = "Stop"',
            COM_RETRY_HELPER,
            ...attachLines(filePath, attachToRunning),
            'if (-not $workbook) {',
            `  $workbook = Invoke-XlideCom { $excel.Workbooks.Open($targetPath, 0, ${mode.readOnly ? '$true' : '$false'}) }`,
            '}',
            ...FOREGROUND_LINES,
        ].join('; ');
    }
    return [
        '$ErrorActionPreference = "Stop"',
        COM_RETRY_HELPER,
        'try {',
        `  $macroName = ${psSingleQuoted(mode.macroName)}`,
        ...attachLines(filePath, attachToRunning),
        'if ($workbook) {',
        '  if (-not $workbook.ReadOnly) {',
        '    throw "REOPEN_BLOCKED|Workbook is already open for editing in Excel. Close it in Excel, then press F5 again so XLIDE can reopen the saved workbook before running the macro."',
        '  }',
        '  try {',
        '    Invoke-XlideCom { $workbook.Close($false) }',
        '    $workbook = $null',
        '  } catch {',
        '    throw ("REOPEN_FAILED|XLIDE could not close the existing read-only workbook before running the macro: " + $_.Exception.Message)',
        '  }',
        '}',
        'try {',
        '  $workbook = Invoke-XlideCom { $excel.Workbooks.Open($targetPath, 0, $true) }',
        '} catch {',
        '  throw ("REOPEN_FAILED|XLIDE could not reopen the workbook. If it is open outside XLIDE, close it in Excel and try again: " + $_.Exception.Message)',
        '}',
        ...FOREGROUND_LINES,
        '$macroRef = "\'" + ($workbook.Name -replace "\'", "\'\'") + "\'!" + $macroName',
        'try {',
        '  Invoke-XlideCom { $excel.Run($macroRef) }',
        '} catch {',
        '  throw ("RUN_FAILED|XLIDE could not run the macro: " + $_.Exception.Message)',
        '}',
        '[Console]::Out.WriteLine("XLIDE_MACRO_OK")',
        '} catch {',
        '  [Console]::Error.WriteLine("XLIDE_MACRO_ERROR|" + $_.Exception.Message)',
        '  exit 1',
        '}',
    ].join('; ');
}

// Backstop so an open/macro-run COM call cannot hang (and leak its powershell +
// Excel COM reference) forever - the common cause is a modal dialog (e.g. a
// MsgBox in the macro) waiting for the user. Generous so a legitimately
// long-running macro is not cut short.
const EXCEL_SCRIPT_TIMEOUT_MS = 300_000;

async function runExcelScript(
    script: string,
    tag: string,
    log: (message: string) => void,
): Promise<{ code: number | null; stderrLines: string[]; spawnError?: Error; timedOut: boolean }> {
    const run = runPowerShell({
        args: ['-Command', script],
        timeoutMs: EXCEL_SCRIPT_TIMEOUT_MS,
        onSpawn: (pid) => log(`[${tag}] Spawned powershell.exe (pid=${pid ?? 'unknown'})`),
        onStdoutLine: (line) => log(`[${tag} stdout] ${line}`),
        onStderrLine: (line) => log(`[${tag} stderr] ${line}`),
    });
    const result = await run.result;
    if (!result.spawnError) {
        log(`[${tag}] powershell exited with code=${result.code} signal=${result.signal ?? 'none'}${result.timedOut ? ' (timed out)' : ''}`);
    }
    return result;
}

/** Opens (or re-foregrounds) the workbook in Excel. Rejects only on spawn failure. */
export async function openWorkbookInExcel(
    filePath: string,
    options: { attachToRunning: boolean; readOnly: boolean },
    log: (message: string) => void,
): Promise<void> {
    const script = buildExcelLaunchScript({
        filePath,
        attachToRunning: options.attachToRunning,
        mode: { kind: 'open', readOnly: options.readOnly },
    });
    log(`[openWorkbook] Running: powershell -Command "${script}"`);
    const result = await runExcelScript(script, 'openWorkbook', log);
    if (result.spawnError) {
        log(`[openWorkbook] Error: ${result.spawnError.message}`);
        throw result.spawnError;
    }
    if (result.timedOut) {
        log('[openWorkbook] Timed out and the launch process was killed; Excel may be showing a dialog.');
    }
}

/** Reopens the workbook read-only and runs the macro; rejects with ExcelMacroError. */
export async function runWorkbookMacroReadOnly(
    filePath: string,
    macroName: string,
    options: { attachToRunning: boolean },
    log: (message: string) => void,
): Promise<void> {
    const script = buildExcelLaunchScript({
        filePath,
        attachToRunning: options.attachToRunning,
        mode: { kind: 'macroReadOnly', macroName },
    });
    log(`[runMacro] Running: ${macroName}`);
    log(`[runMacro] Script: ${script}`);
    const result = await runExcelScript(script, 'runMacro', log);
    if (result.spawnError) {
        log(`[runMacro] Error: ${result.spawnError.message}`);
        throw new ExcelMacroError(result.spawnError.message, 'RUN_FAILED');
    }
    if (result.timedOut) {
        throw new ExcelMacroError(
            'Excel did not respond within the time limit. A dialog may be open in Excel '
            + '(for example a MsgBox from the macro); close it and try again, or the macro may be running too long.',
            'RUN_FAILED',
        );
    }
    if (result.code === 0) {
        return;
    }
    const sentinel = result.stderrLines.find((line) => line.includes(MACRO_ERROR_SENTINEL));
    const raw = sentinel
        ? sentinel.slice(sentinel.indexOf(MACRO_ERROR_SENTINEL) + MACRO_ERROR_SENTINEL.length)
        : result.stderrLines.join('\n') || `PowerShell exited with code ${result.code}`;
    throw excelMacroErrorFromRaw(raw);
}

function excelMacroErrorFromRaw(raw: string): ExcelMacroError {
    const pipe = raw.indexOf('|');
    const code = pipe >= 0 ? raw.slice(0, pipe) : '';
    if (code === 'REOPEN_BLOCKED' || code === 'REOPEN_FAILED' || code === 'RUN_FAILED') {
        return new ExcelMacroError(raw.slice(pipe + 1), code);
    }
    return new ExcelMacroError(raw, 'UNKNOWN');
}
