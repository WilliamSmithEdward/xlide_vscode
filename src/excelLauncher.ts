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
        `$targetName = ${psSingleQuoted(path.basename(filePath))}`,
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

// Activate the workbook and bring the Excel window to the foreground.
const FOREGROUND_LINES: readonly string[] = [
    '$workbook.Activate()',
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
            ...attachLines(filePath, attachToRunning),
            'if (-not $workbook) {',
            `  $workbook = $excel.Workbooks.Open($targetPath, 0, ${mode.readOnly ? '$true' : '$false'})`,
            '}',
            ...FOREGROUND_LINES,
        ].join('; ');
    }
    return [
        '$ErrorActionPreference = "Stop"',
        'try {',
        `  $macroName = ${psSingleQuoted(mode.macroName)}`,
        ...attachLines(filePath, attachToRunning),
        'if ($workbook) {',
        '  if (-not $workbook.ReadOnly) {',
        '    throw "REOPEN_BLOCKED|Workbook is already open for editing in Excel. Close it in Excel, then press F5 again so XLIDE can reopen the saved workbook before running the macro."',
        '  }',
        '  try {',
        '    $workbook.Close($false)',
        '    $workbook = $null',
        '  } catch {',
        '    throw ("REOPEN_FAILED|XLIDE could not close the existing read-only workbook before running the macro: " + $_.Exception.Message)',
        '  }',
        '}',
        'try {',
        '  $workbook = $excel.Workbooks.Open($targetPath, 0, $true)',
        '} catch {',
        '  throw ("REOPEN_FAILED|XLIDE could not reopen the workbook. If it is open outside XLIDE, close it in Excel and try again: " + $_.Exception.Message)',
        '}',
        ...FOREGROUND_LINES,
        '$macroRef = "\'" + $workbook.Name + "\'!" + $macroName',
        'try {',
        '  $excel.Run($macroRef)',
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

async function runExcelScript(
    script: string,
    tag: string,
    log: (message: string) => void,
): Promise<{ code: number | null; stderrLines: string[]; spawnError?: Error }> {
    const run = runPowerShell({
        args: ['-Command', script],
        onSpawn: (pid) => log(`[${tag}] Spawned powershell.exe (pid=${pid ?? 'unknown'})`),
        onStdoutLine: (line) => log(`[${tag} stdout] ${line}`),
        onStderrLine: (line) => log(`[${tag} stderr] ${line}`),
    });
    const result = await run.result;
    if (!result.spawnError) {
        log(`[${tag}] powershell exited with code=${result.code} signal=${result.signal ?? 'none'}`);
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
