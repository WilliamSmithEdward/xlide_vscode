import * as path from 'path';
import * as vscode from 'vscode';
import { psSingleQuoted, runPowerShell } from './util/powershell';
import { openWorkbookInExcel } from './excelLauncher';
import { WORKBOOK_LOCKED_ERROR_RE } from './xlideCommandLog';
import { errorMessage } from './util/errors';
import {
    xlideExcelCoordinationModeFromConfig,
    xlideExcelReopenAfterCloseFromConfig,
    xlideExcelReopenModeFromConfig,
    xlideExcelReopenReadOnlyAfterSaveFromConfig,
    xlideExcelTrackOpenedWorkbooksFromConfig,
    type ExcelCoordinationMode,
    type ExcelReopenMode,
} from './globalSettings';

/**
 * Coordinates XLIDE writes/launches with Excel's file lock. When a workbook is
 * open in Excel the OS locks the .xlsm, so a save/add/rename/delete or an F5
 * macro reopen fails. The user-chosen xlide.excelIntegration.coordinationMode
 * decides what happens: block (default), gracefully close the workbook XLIDE
 * opened, or force-close any Excel (killing it as a last resort).
 *
 * Windows + Excel-COM only; every entry point no-ops on other platforms so the
 * caller falls back to its existing block-and-warn path.
 */

// Shared logger so callers without their own output channel (the file-system
// provider, the shared module operations) still trace coordination steps.
// Set once at activation; defaults to a no-op for tests.
let sharedLog: (message: string) => void = () => { /* no-op until wired */ };

export function setExcelCoordinationLog(log: (message: string) => void): void {
    sharedLog = log;
}

// Session-scoped set of workbook paths XLIDE itself opened in Excel, so
// closeTracked only ever closes workbooks the user opened through XLIDE.
const xlideOpenedWorkbooks = new Set<string>();

function workbookKey(filePath: string): string {
    return path.win32.normalize(filePath).toLowerCase();
}

export function markWorkbookOpenedByXlide(filePath: string): void {
    xlideOpenedWorkbooks.add(workbookKey(filePath));
}

export function wasWorkbookOpenedByXlide(filePath: string): boolean {
    return xlideOpenedWorkbooks.has(workbookKey(filePath));
}

export function forgetWorkbookOpenedByXlide(filePath: string): void {
    xlideOpenedWorkbooks.delete(workbookKey(filePath));
}

export interface ExcelCoordinationSettings {
    mode: ExcelCoordinationMode;
    trackOpenedWorkbooks: boolean;
    reopenAfterClose: boolean;
    reopenMode: ExcelReopenMode;
    reopenReadOnlyAfterSave: boolean;
}

export function resolveExcelCoordinationSettings(): ExcelCoordinationSettings {
    const config = vscode.workspace.getConfiguration('xlide');
    return {
        mode: xlideExcelCoordinationModeFromConfig(config).value,
        trackOpenedWorkbooks: xlideExcelTrackOpenedWorkbooksFromConfig(config).value,
        reopenAfterClose: xlideExcelReopenAfterCloseFromConfig(config).value,
        reopenMode: xlideExcelReopenModeFromConfig(config).value,
        reopenReadOnlyAfterSave: xlideExcelReopenReadOnlyAfterSaveFromConfig(config).value,
    };
}

/**
 * Whether the active policy permits closing this workbook to free the lock.
 * - closeForce: always (force is the user's explicit "close no matter what").
 * - closeTracked: only workbooks XLIDE opened, unless the user opted out of
 *   tracking (then any matching workbook).
 * - block: never.
 */
export function shouldAttemptClose(
    settings: ExcelCoordinationSettings,
    filePath: string,
): boolean {
    switch (settings.mode) {
        case 'closeForce':
            return true;
        case 'closeTracked':
            return settings.trackOpenedWorkbooks ? wasWorkbookOpenedByXlide(filePath) : true;
        default:
            return false;
    }
}

const CLOSE_SENTINEL = 'XLIDE_CLOSE|';
const CLOSE_ERROR_SENTINEL = 'XLIDE_CLOSE_ERROR|';

/**
 * PowerShell that attaches to the running Excel, closes the target workbook
 * WITHOUT saving (so XLIDE's file write wins, never Excel's stale copy), then
 * checks whether the file lock is gone. For force mode, if the file is still
 * locked it kills every EXCEL.EXE and re-checks. Exported for unit testing.
 */
export function buildCloseWorkbookScript(filePath: string, force: boolean): string {
    return [
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        `$force = ${force ? '$true' : '$false'}`,
        '$closed = $false',
        '$found = $false',
        '$wasReadOnly = $false',
        '$excel = $null',
        'try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
        'if ($excel) {',
        '  foreach ($wb in @($excel.Workbooks)) {',
        '    if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) {',
        '      $found = $true',
        '      $wasReadOnly = [bool]$wb.ReadOnly',
        '      try { $wb.Close($false); $closed = $true } catch { }',
        '      break',
        '    }',
        '  }',
        '}',
        'function Test-XlideLocked { try { $fs = [System.IO.File]::Open($targetPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None); $fs.Close(); return $false } catch { return $true } }',
        '$locked = Test-XlideLocked',
        'if ($locked -and $force) {',
        '  Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue',
        '  Start-Sleep -Milliseconds 600',
        '  $locked = Test-XlideLocked',
        '}',
        '[Console]::Out.WriteLine("XLIDE_CLOSE|closed=" + $closed + "|locked=" + $locked + "|found=" + $found + "|wasReadOnly=" + $wasReadOnly)',
    ].join('; ');
}

export interface CloseWorkbookResult {
    closed: boolean;
    stillLocked: boolean;
    /** Whether the workbook was open read-only before closing; undefined if it
     *  was not found in the reachable Excel instance. */
    wasReadOnly?: boolean;
    error?: string;
}

/** Runs the close script. Rejects only on a PowerShell spawn failure. */
export async function closeWorkbookInExcel(
    filePath: string,
    options: { force: boolean },
    log: (message: string) => void = sharedLog,
): Promise<CloseWorkbookResult> {
    if (process.platform !== 'win32') {
        return { closed: false, stillLocked: true, error: 'not-win32' };
    }
    const script = buildCloseWorkbookScript(filePath, options.force);
    log(`[excelCoord] close (force=${options.force}): ${filePath}`);
    const run = runPowerShell({
        args: ['-Command', script],
        timeoutMs: 20000,
        onStdoutLine: (line) => log(`[excelCoord stdout] ${line}`),
        onStderrLine: (line) => log(`[excelCoord stderr] ${line}`),
    });
    const result = await run.result;
    if (result.spawnError) {
        return { closed: false, stillLocked: true, error: result.spawnError.message };
    }
    const sentinel = result.stdoutLines.find((line) => line.startsWith(CLOSE_SENTINEL));
    if (sentinel) {
        const found = /found=True/i.test(sentinel);
        return {
            closed: /closed=True/i.test(sentinel),
            stillLocked: /locked=True/i.test(sentinel),
            wasReadOnly: found ? /wasReadOnly=True/i.test(sentinel) : undefined,
        };
    }
    const errLine = result.stderrLines.find((line) => line.includes(CLOSE_ERROR_SENTINEL));
    return {
        closed: false,
        stillLocked: true,
        error: errLine ?? `powershell exited with code ${result.code}`,
    };
}

/**
 * Closes the workbook to free the lock when the policy allows it. No-op (and
 * `attempted: false`) under block mode, off-Windows, or when closeTracked is
 * limited to XLIDE-opened workbooks and this one was opened elsewhere.
 */
export async function tryCoordinatedClose(
    filePath: string,
    log: (message: string) => void = sharedLog,
    settings: ExcelCoordinationSettings = resolveExcelCoordinationSettings(),
): Promise<{ attempted: boolean; freed: boolean; wasReadOnly?: boolean }> {
    if (process.platform !== 'win32' || !shouldAttemptClose(settings, filePath)) {
        return { attempted: false, freed: false };
    }
    const result = await closeWorkbookInExcel(
        filePath,
        { force: settings.mode === 'closeForce' },
        log,
    );
    if (result.error) {
        log(`[excelCoord] close reported: ${result.error}`);
    }
    return { attempted: true, freed: !result.stillLocked, wasReadOnly: result.wasReadOnly };
}

/**
 * Resolves the reopen mode to a concrete read-only flag. 'lastState' restores
 * how the workbook was before XLIDE closed it; when that is unknown (the close
 * could not see it), it falls back to the safe read-only choice.
 */
export function resolveReopenReadOnly(mode: ExcelReopenMode, wasReadOnly: boolean | undefined): boolean {
    switch (mode) {
        case 'readOnly':
            return true;
        case 'readWrite':
            return false;
        case 'lastState':
            return wasReadOnly ?? true;
    }
}

async function reopenWorkbookAfterClose(
    filePath: string,
    readOnly: boolean,
    log: (message: string) => void = sharedLog,
): Promise<void> {
    try {
        await openWorkbookInExcel(
            filePath,
            { attachToRunning: true, readOnly },
            log,
        );
        markWorkbookOpenedByXlide(filePath);
    } catch (err) {
        log(`[excelCoord] reopen after close failed: ${errorMessage(err)}`);
    }
}

const REFRESH_SENTINEL = 'XLIDE_REFRESH|';

/**
 * PowerShell that refreshes Excel's stale view of a workbook open READ-ONLY:
 * a read-only open does not lock the file, so XLIDE's save succeeds but Excel
 * keeps its older in-memory copy. This closes and reopens the workbook READ-ONLY
 * only when it is actually open read-only in the running Excel. It never opens
 * a workbook that is closed, and never touches one open for editing. Exported
 * for unit testing.
 */
export function buildRefreshReadOnlyScript(filePath: string): string {
    return [
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        '$refreshed = $false',
        '$excel = $null',
        'try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } catch { }',
        'if ($excel) {',
        '  foreach ($wb in @($excel.Workbooks)) {',
        '    if (($wb.FullName -ieq $targetPath) -or ($wb.Name -ieq $targetName)) {',
        '      if ($wb.ReadOnly) {',
        '        try { $wb.Close($false); $excel.Workbooks.Open($targetPath, 0, $true) | Out-Null; $refreshed = $true } catch { }',
        '      }',
        '      break',
        '    }',
        '  }',
        '}',
        '[Console]::Out.WriteLine("XLIDE_REFRESH|refreshed=" + $refreshed)',
    ].join('; ');
}

/**
 * After a save that succeeded while the workbook was open read-only in Excel,
 * reopen it read-only so Excel's view matches the saved file. Best-effort and
 * silent (it does not steal focus or block the save); rejects nothing.
 */
export async function refreshReadOnlyViewAfterSave(
    filePath: string,
    log: (message: string) => void = sharedLog,
): Promise<void> {
    if (process.platform !== 'win32') {
        return;
    }
    const script = buildRefreshReadOnlyScript(filePath);
    log(`[excelCoord] refresh read-only view: ${filePath}`);
    try {
        const run = runPowerShell({
            args: ['-Command', script],
            timeoutMs: 20000,
            onStdoutLine: (line) => log(`[excelCoord stdout] ${line}`),
            onStderrLine: (line) => log(`[excelCoord stderr] ${line}`),
        });
        const result = await run.result;
        const sentinel = result.stdoutLines.find((line) => line.startsWith(REFRESH_SENTINEL));
        if (sentinel && /refreshed=True/i.test(sentinel)) {
            markWorkbookOpenedByXlide(filePath);
        }
    } catch (err) {
        log(`[excelCoord] refresh read-only view failed: ${errorMessage(err)}`);
    }
}

/**
 * Runs a workbook file write (save / add / rename / delete) and, if it fails
 * because Excel holds the lock, applies the coordination policy: under a close
 * mode it closes the workbook in Excel, retries the write once, and (when
 * configured) reopens the workbook so the user's Excel view is restored. Under
 * block mode (or when closing is not permitted) it rethrows the lock error so
 * the caller surfaces the existing "close it in Excel" guidance.
 *
 * On a write that SUCCEEDS while the workbook is open read-only in Excel (no
 * lock error, so coordination never fires), the reopenReadOnlyAfterSave setting
 * refreshes Excel's stale view in the background.
 */
export async function runWriteWithExcelCoordination<T>(
    filePath: string,
    write: () => Promise<T>,
    log: (message: string) => void = sharedLog,
): Promise<T> {
    try {
        const result = await write();
        if (process.platform === 'win32'
            && resolveExcelCoordinationSettings().reopenReadOnlyAfterSave) {
            // Fire-and-forget: refresh Excel's stale read-only view without
            // delaying the save. The script no-ops unless the workbook is
            // actually open read-only in Excel.
            void refreshReadOnlyViewAfterSave(filePath, log);
        }
        return result;
    } catch (err) {
        if (process.platform !== 'win32' || !WORKBOOK_LOCKED_ERROR_RE.test(errorMessage(err))) {
            throw err;
        }
        const settings = resolveExcelCoordinationSettings();
        if (!shouldAttemptClose(settings, filePath)) {
            throw err;
        }
        log(`[excelCoord] write locked; coordinationMode=${settings.mode}, closing in Excel`);
        const { freed, wasReadOnly } = await tryCoordinatedClose(filePath, log, settings);
        if (!freed) {
            log('[excelCoord] close did not confirm the lock was freed; retrying the write anyway');
        }
        // The retry is the source of truth: if it still fails the lock survived,
        // and the caller surfaces the locked-workbook guidance.
        const result = await write();
        if (settings.reopenAfterClose) {
            // reopenWorkbookAfterClose re-marks the workbook as XLIDE-opened on a
            // successful reopen; on failure the prior tracking is left intact so a
            // later closeTracked save can still free the lock. We must NOT forget
            // here; doing so before/around a failed reopen would strand tracking.
            const readOnly = resolveReopenReadOnly(settings.reopenMode, wasReadOnly);
            await reopenWorkbookAfterClose(filePath, readOnly, log);
        } else {
            // Intentionally left closed: the workbook is no longer open in Excel.
            forgetWorkbookOpenedByXlide(filePath);
        }
        return result;
    }
}
