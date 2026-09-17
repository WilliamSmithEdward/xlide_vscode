import * as path from 'path';
import * as vscode from 'vscode';
import { psSingleQuoted, runPowerShell } from './util/powershell';
import { OFFICE_HOST_APPS, officeHostForPath, type OfficeHostApp } from './officeHostApps';
import { openFileInHost } from './officeHostLauncher';
import { PROJECT_LOCKED_ERROR_RE } from './xlideCommandLog';
import { errorMessage } from './util/errors';
import {
    xlideOfficeCoordinationModeFromConfig,
    xlideOfficeReopenAfterCloseFromConfig,
    xlideOfficeReopenModeFromConfig,
    xlideOfficeReopenReadOnlyAfterSaveFromConfig,
    xlideOfficeTrackOpenedFilesFromConfig,
    type OfficeCoordinationMode,
    type OfficeReopenMode,
} from './globalSettings';

/**
 * Coordinates XLIDE writes and launches with the file lock an Office
 * application holds. XLIDE saves by renaming a temp file over the container,
 * and Windows refuses that rename while Excel, Word, PowerPoint or Access has
 * the file open, so a save/add/rename/delete or an F5 reopen fails. The
 * user-chosen xlide.officeIntegration.coordinationMode decides what happens:
 * block (default), gracefully close the file XLIDE opened, or force-close it
 * in any instance (killing the application as a last resort).
 *
 * The applications differ in one way that matters here, measured on Office
 * 16.0: a workbook open READ-ONLY in Excel does not hold the lock, while a
 * read-only document in Word or presentation in PowerPoint does, and Access
 * has no read-only open at all.
 *
 * Windows + COM only; every entry point no-ops on other platforms so the
 * caller falls back to its existing block-and-warn path.
 */

// Shared logger so callers without their own output channel (the file-system
// provider, the shared module operations) still trace coordination steps.
// Set once at activation; defaults to a no-op for tests.
let sharedLog: (message: string) => void = () => { /* no-op until wired */ };

export function setHostCoordinationLog(log: (message: string) => void): void {
    sharedLog = log;
}

// File paths (refcounted) whose XLIDE-driven post-save reopen is suppressed
// because a caller is about to reopen the file itself. F5 (Run Macro at
// Cursor) saves the dirty module and then reopens the file read-only to run the
// macro; without this the save's background read-only refresh (and a
// close-mode reopen) would race that reopen on the SAME file (transient "open
// for editing" / run failures). Keyed per file so an F5 on one file never
// suppresses a concurrent save's reopen on a different file.
const reopenSuppressedPaths = new Map<string, number>();

export async function withFileReopenSuppressed<T>(
    filePath: string,
    fn: () => Promise<T>,
): Promise<T> {
    const key = projectKey(filePath);
    reopenSuppressedPaths.set(key, (reopenSuppressedPaths.get(key) ?? 0) + 1);
    try {
        return await fn();
    } finally {
        const next = (reopenSuppressedPaths.get(key) ?? 1) - 1;
        if (next > 0) {
            reopenSuppressedPaths.set(key, next);
        } else {
            reopenSuppressedPaths.delete(key);
        }
    }
}

function isFileReopenSuppressed(filePath: string): boolean {
    return reopenSuppressedPaths.has(projectKey(filePath));
}

// Files whose read-only refresh PowerShell is currently in flight, so rapid
// saves cannot stack concurrent close/reopen passes on the same file.
const readOnlyRefreshInFlight = new Set<string>();

// Session-scoped set of file paths XLIDE itself opened in their application,
// so closeTracked only ever closes files the user opened through XLIDE.
const xlideOpenedFiles = new Set<string>();

function projectKey(filePath: string): string {
    return path.win32.normalize(filePath).toLowerCase();
}

export function markFileOpenedByXlide(filePath: string): void {
    xlideOpenedFiles.add(projectKey(filePath));
}

export function wasFileOpenedByXlide(filePath: string): boolean {
    return xlideOpenedFiles.has(projectKey(filePath));
}

export function forgetFileOpenedByXlide(filePath: string): void {
    xlideOpenedFiles.delete(projectKey(filePath));
}

export interface HostCoordinationSettings {
    mode: OfficeCoordinationMode;
    trackOpenedFiles: boolean;
    reopenAfterClose: boolean;
    reopenMode: OfficeReopenMode;
    reopenReadOnlyAfterSave: boolean;
}

export function resolveHostCoordinationSettings(): HostCoordinationSettings {
    const config = vscode.workspace.getConfiguration('xlide');
    return {
        mode: xlideOfficeCoordinationModeFromConfig(config).value,
        trackOpenedFiles: xlideOfficeTrackOpenedFilesFromConfig(config).value,
        reopenAfterClose: xlideOfficeReopenAfterCloseFromConfig(config).value,
        reopenMode: xlideOfficeReopenModeFromConfig(config).value,
        reopenReadOnlyAfterSave: xlideOfficeReopenReadOnlyAfterSaveFromConfig(config).value,
    };
}

/**
 * Whether the active policy permits closing this file, open for editing or
 * not, to free the lock.
 * - closeForce: always (force is the user's explicit "close no matter what").
 * - closeTracked: only files XLIDE opened, unless the user opted out of
 *   tracking (then any matching file).
 * - block: never.
 */
export function shouldAttemptClose(
    settings: HostCoordinationSettings,
    filePath: string,
): boolean {
    switch (settings.mode) {
        case 'closeForce':
            return true;
        case 'closeTracked':
            return settings.trackOpenedFiles ? wasFileOpenedByXlide(filePath) : true;
        default:
            return false;
    }
}

/** What a coordinated close may close: any open copy, or only a read-only one. */
export type CoordinatedCloseScope = 'any' | 'readOnlyCopy';

/**
 * What a blocked write may close, or undefined when it may close nothing.
 *
 * Beyond the mode's own answer, a READ-ONLY copy XLIDE itself opened is
 * XLIDE's to close under every mode: it holds no edits that can be saved, and
 * F5 already closes and reopens one on every run. Word and PowerPoint keep the
 * file locked even for a read-only open, so without this a save after F5 would
 * fail until the user closed the file by hand. The script confirms the
 * read-only state itself, so a copy switched to editing is left alone.
 */
export function closeScopeForWrite(
    settings: HostCoordinationSettings,
    filePath: string,
): CoordinatedCloseScope | undefined {
    if (shouldAttemptClose(settings, filePath)) {
        return 'any';
    }
    return wasFileOpenedByXlide(filePath) ? 'readOnlyCopy' : undefined;
}

/**
 * How each application exposes the file it has open. `$app` is the running
 * application, `$targetPath` / `$targetName` the file; every `find` leaves the
 * open copy in `$file` (or `$null`). Each call was measured on Office 16.0.
 */
interface HostOpenFileDialect {
    find: readonly string[];
    /** Expression: is `$file` open read-only? */
    readOnly: string;
    /** Statement closing `$file` WITHOUT saving, so XLIDE's write wins. */
    close: string;
    /** Statement reopening the file read-only; absent where no such open exists. */
    reopenReadOnly?: string;
}

/**
 * Matches by full path first. Office reports some files under another
 * spelling (a OneDrive URL, a UNC path for a mapped drive), so the file name
 * is the fallback - but never onto a different local file that merely shares
 * the name, which Word and PowerPoint can have open at the same time.
 */
function collectionFind(collection: string): string[] {
    return [
        `$files = @($app.${collection})`,
        'foreach ($f in $files) { if ($f.FullName -ieq $targetPath) { $file = $f; break } }',
        'if (-not $file) { foreach ($f in $files) { if (($f.Name -ieq $targetName) -and -not (Test-XlideOtherLocalFile $f.FullName)) { $file = $f; break } } }',
    ];
}

const OTHER_LOCAL_FILE_HELPER =
    'function Test-XlideOtherLocalFile($openPath) { return (($openPath -match "^[A-Za-z]:\\\\") -and ($targetPath -match "^[A-Za-z]:\\\\") -and ($openPath -ine $targetPath)) }';

const HOST_OPEN_FILE_DIALECTS: Record<OfficeHostApp, HostOpenFileDialect> = {
    excel: {
        find: collectionFind('Workbooks'),
        readOnly: '[bool]$file.ReadOnly',
        close: '$file.Close($false)',
        reopenReadOnly: '$app.Workbooks.Open($targetPath, 0, $true) | Out-Null',
    },
    word: {
        find: collectionFind('Documents'),
        readOnly: '[bool]$file.ReadOnly',
        // wdDoNotSaveChanges
        close: '$file.Close(0)',
        // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        reopenReadOnly: '$app.Documents.Open($targetPath, $false, $true, $false) | Out-Null',
    },
    powerpoint: {
        find: collectionFind('Presentations'),
        // MsoTriState: msoTrue is -1.
        readOnly: '($file.ReadOnly -ne 0)',
        // Presentation.Close never prompts under automation, edited or not.
        close: '$file.Close()',
        // Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
        reopenReadOnly: '$app.Presentations.Open($targetPath, -1, 0, -1) | Out-Null',
    },
    access: {
        // One instance holds one database, and there is no read-only open.
        find: [
            '$open = ""',
            'try { $open = $app.CurrentProject.FullName } catch { }',
            'if ($open -ieq $targetPath) { $file = $app.CurrentProject }',
        ],
        readOnly: '$false',
        close: '$app.CloseCurrentDatabase()',
    },
};

function targetLines(filePath: string, host: OfficeHostApp): string[] {
    return [
        `$targetPath = ${psSingleQuoted(filePath)}`,
        `$targetName = ${psSingleQuoted(path.win32.basename(filePath))}`,
        OTHER_LOCAL_FILE_HELPER,
        '$app = $null',
        '$file = $null',
        // Only ever a RUNNING application: a close must not start one.
        `try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("${OFFICE_HOST_APPS[host].progId}") } catch { }`,
    ];
}

const CLOSE_SENTINEL = 'XLIDE_CLOSE|';
const CLOSE_ERROR_SENTINEL = 'XLIDE_CLOSE_ERROR|';

export interface CloseFileScriptOptions {
    /** Kill the application's processes when the file stays locked. */
    force: boolean;
    /** Close the open copy only when it is open read-only. */
    onlyReadOnlyCopy?: boolean;
}

/**
 * PowerShell that attaches to the running application that owns the file,
 * closes the file WITHOUT saving (so XLIDE's file write wins, never the
 * application's stale copy), then checks whether the file lock is gone. For
 * force mode, if the file is still locked it kills every process of that
 * application and re-checks. Exported for unit testing.
 */
export function buildCloseFileScript(filePath: string, options: CloseFileScriptOptions): string {
    const host = officeHostForPath(filePath) ?? 'excel';
    const dialect = HOST_OPEN_FILE_DIALECTS[host];
    return [
        ...targetLines(filePath, host),
        `$force = ${options.force ? '$true' : '$false'}`,
        `$onlyReadOnlyCopy = ${options.onlyReadOnlyCopy ? '$true' : '$false'}`,
        '$closed = $false',
        '$found = $false',
        '$forced = $false',
        '$wasReadOnly = $false',
        'if ($app) {',
        ...dialect.find.map((line) => `  ${line}`),
        '  if ($file) {',
        '    $found = $true',
        `    $wasReadOnly = ${dialect.readOnly}`,
        `    if ($wasReadOnly -or -not $onlyReadOnlyCopy) { try { ${dialect.close}; $closed = $true } catch { } }`,
        '  }',
        '}',
        'function Test-XlideLocked { try { $fs = [System.IO.File]::Open($targetPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None); $fs.Close(); return $false } catch { return $true } }',
        '$locked = Test-XlideLocked',
        // The application can hold its handle a moment after Close returns.
        'for ($__i = 0; $locked -and $closed -and $__i -lt 5; $__i++) { Start-Sleep -Milliseconds 200; $locked = Test-XlideLocked }',
        'if ($locked -and $force) {',
        `  Get-Process -Name ${OFFICE_HOST_APPS[host].processName} -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue`,
        '  $forced = $true',
        '  Start-Sleep -Milliseconds 600',
        '  $locked = Test-XlideLocked',
        '}',
        '[Console]::Out.WriteLine("XLIDE_CLOSE|closed=" + $closed + "|locked=" + $locked + "|found=" + $found + "|wasReadOnly=" + $wasReadOnly + "|forced=" + $forced)',
    ].join('; ');
}

export interface CloseFileResult {
    /** The open copy was closed through the application. */
    closed: boolean;
    /** The application's processes were killed to free the lock. */
    forced: boolean;
    stillLocked: boolean;
    /** Whether the file was open read-only before closing; undefined if it
     *  was not found in the reachable instance of its application. */
    wasReadOnly?: boolean;
    error?: string;
}

/** Runs one coordination script under the caller's logging and a 20 second budget. */
function runCoordinationScript(script: string, log: (message: string) => void) {
    return runPowerShell({
        args: ['-Command', script],
        timeoutMs: 20000,
        onStdoutLine: (line) => log(`[hostCoord stdout] ${line}`),
        onStderrLine: (line) => log(`[hostCoord stderr] ${line}`),
    }).result;
}

/** Runs the close script. Rejects only on a PowerShell spawn failure. */
export async function closeFileInHost(
    filePath: string,
    options: CloseFileScriptOptions,
    log: (message: string) => void = sharedLog,
): Promise<CloseFileResult> {
    if (process.platform !== 'win32' || !officeHostForPath(filePath)) {
        return { closed: false, forced: false, stillLocked: true, error: 'not-supported' };
    }
    const script = buildCloseFileScript(filePath, options);
    log(`[hostCoord] close (force=${options.force}, onlyReadOnlyCopy=${Boolean(options.onlyReadOnlyCopy)}): ${filePath}`);
    const result = await runCoordinationScript(script, log);
    if (result.spawnError) {
        return { closed: false, forced: false, stillLocked: true, error: result.spawnError.message };
    }
    const sentinel = result.stdoutLines.find((line) => line.startsWith(CLOSE_SENTINEL));
    if (sentinel) {
        const found = /found=True/i.test(sentinel);
        return {
            closed: /closed=True/i.test(sentinel),
            forced: /forced=True/i.test(sentinel),
            stillLocked: /locked=True/i.test(sentinel),
            wasReadOnly: found ? /wasReadOnly=True/i.test(sentinel) : undefined,
        };
    }
    const errLine = result.stderrLines.find((line) => line.includes(CLOSE_ERROR_SENTINEL));
    return {
        closed: false,
        forced: false,
        stillLocked: true,
        error: errLine ?? `powershell exited with code ${result.code}`,
    };
}

export interface CoordinatedCloseOutcome {
    attempted: boolean;
    freed: boolean;
    /** Something the user was looking at went away: a close, or a forced kill. */
    viewLost: boolean;
    wasReadOnly?: boolean;
}

/**
 * Closes the file to free the lock when the policy allows it. No-op (and
 * `attempted: false`) off-Windows, or when neither the mode nor the
 * read-only-copy rule lets this file be closed.
 */
export async function tryCoordinatedClose(
    filePath: string,
    log: (message: string) => void = sharedLog,
    settings: HostCoordinationSettings = resolveHostCoordinationSettings(),
): Promise<CoordinatedCloseOutcome> {
    const scope = process.platform === 'win32' ? closeScopeForWrite(settings, filePath) : undefined;
    if (!scope) {
        return { attempted: false, freed: false, viewLost: false };
    }
    const result = await closeFileInHost(
        filePath,
        { force: scope === 'any' && settings.mode === 'closeForce', onlyReadOnlyCopy: scope === 'readOnlyCopy' },
        log,
    );
    if (result.error) {
        log(`[hostCoord] close reported: ${result.error}`);
    }
    return {
        attempted: true,
        freed: !result.stillLocked,
        viewLost: result.closed || result.forced,
        wasReadOnly: result.wasReadOnly,
    };
}

/**
 * Resolves the reopen mode to a concrete read-only flag. 'lastState' restores
 * how the file was before XLIDE closed it; when that is unknown (the close
 * could not see it), it falls back to the safe read-only choice.
 */
export function resolveReopenReadOnly(mode: OfficeReopenMode, wasReadOnly: boolean | undefined): boolean {
    switch (mode) {
        case 'readOnly':
            return true;
        case 'readWrite':
            return false;
        case 'lastState':
            return wasReadOnly ?? true;
    }
}

async function reopenFileAfterClose(
    filePath: string,
    readOnly: boolean,
    log: (message: string) => void = sharedLog,
): Promise<void> {
    try {
        await openFileInHost(
            filePath,
            { attachToRunning: true, readOnly },
            log,
        );
        markFileOpenedByXlide(filePath);
    } catch (err) {
        log(`[hostCoord] reopen after close failed: ${errorMessage(err)}`);
    }
}

const REFRESH_SENTINEL = 'XLIDE_REFRESH|';

/**
 * PowerShell that refreshes an application's stale view of a file open
 * READ-ONLY: where a read-only open does not lock the file (Excel), XLIDE's
 * save succeeds but the application keeps its older in-memory copy. This
 * closes and reopens the file READ-ONLY only when it is actually open
 * read-only in the running application. It never opens a file that is closed,
 * and never touches one open for editing. Undefined for an application with no
 * read-only open (Access). Exported for unit testing.
 */
export function buildRefreshReadOnlyScript(filePath: string): string | undefined {
    const host = officeHostForPath(filePath) ?? 'excel';
    const dialect = HOST_OPEN_FILE_DIALECTS[host];
    if (!dialect.reopenReadOnly) {
        return undefined;
    }
    return [
        ...targetLines(filePath, host),
        '$refreshed = $false',
        'if ($app) {',
        ...dialect.find.map((line) => `  ${line}`),
        `  if ($file -and ${dialect.readOnly}) {`,
        `    try { ${dialect.close} } catch { }`,
        '    for ($__i = 0; $__i -lt 5; $__i++) {',
        `      try { ${dialect.reopenReadOnly}; $refreshed = $true; break } catch { Start-Sleep -Milliseconds 200 }`,
        '    }',
        '  }',
        '}',
        '[Console]::Out.WriteLine("XLIDE_REFRESH|refreshed=" + $refreshed)',
    ].join('; ');
}

/**
 * After a save that succeeded while the file was open read-only in its
 * application, reopen it read-only so the application's view matches the saved
 * file. Best-effort and silent (it does not steal focus or block the save);
 * rejects nothing.
 */
export async function refreshReadOnlyViewAfterSave(
    filePath: string,
    log: (message: string) => void = sharedLog,
): Promise<void> {
    if (process.platform !== 'win32') {
        return;
    }
    const script = buildRefreshReadOnlyScript(filePath);
    const key = projectKey(filePath);
    if (!script || readOnlyRefreshInFlight.has(key)) {
        return;
    }
    readOnlyRefreshInFlight.add(key);
    log(`[hostCoord] refresh read-only view: ${filePath}`);
    try {
        const result = await runCoordinationScript(script, log);
        const sentinel = result.stdoutLines.find((line) => line.startsWith(REFRESH_SENTINEL));
        if (sentinel && /refreshed=True/i.test(sentinel)) {
            markFileOpenedByXlide(filePath);
        }
    } catch (err) {
        log(`[hostCoord] refresh read-only view failed: ${errorMessage(err)}`);
    } finally {
        readOnlyRefreshInFlight.delete(key);
    }
}

/**
 * Runs a container file write (save / add / rename / delete) and, if it fails
 * because the file's application holds the lock, applies the coordination
 * policy: it closes the file in that application where the policy allows,
 * retries the write once, and (when configured) reopens the file so the user's
 * view is restored. When closing is not permitted it rethrows the lock error
 * so the caller surfaces the "close it and try again" guidance.
 *
 * On a write that SUCCEEDS while the file is open read-only (no lock error, so
 * coordination never fires), the reopenReadOnlyAfterSave setting refreshes the
 * application's stale view in the background.
 */
export async function runWriteWithHostCoordination<T>(
    filePath: string,
    write: () => Promise<T>,
    log: (message: string) => void = sharedLog,
): Promise<T> {
    if (!officeHostForPath(filePath)) {
        // A VB6 project's modules are plain files no application holds open.
        return write();
    }
    try {
        const result = await write();
        if (process.platform === 'win32'
            && !isFileReopenSuppressed(filePath)
            && resolveHostCoordinationSettings().reopenReadOnlyAfterSave) {
            // Fire-and-forget: refresh the stale read-only view without
            // delaying the save. The script no-ops unless the file is actually
            // open read-only. Skipped while a caller (F5) is about to reopen
            // the file itself.
            void refreshReadOnlyViewAfterSave(filePath, log);
        }
        return result;
    } catch (err) {
        if (process.platform !== 'win32' || !PROJECT_LOCKED_ERROR_RE.test(errorMessage(err))) {
            throw err;
        }
        const settings = resolveHostCoordinationSettings();
        if (!closeScopeForWrite(settings, filePath)) {
            throw err;
        }
        const appName = OFFICE_HOST_APPS[officeHostForPath(filePath) ?? 'excel'].noun;
        log(`[hostCoord] write locked; coordinationMode=${settings.mode}, closing in ${appName}`);
        const { freed, viewLost, wasReadOnly } = await tryCoordinatedClose(filePath, log, settings);
        if (!freed) {
            log('[hostCoord] close did not confirm the lock was freed; retrying the write anyway');
        }
        // The retry is the source of truth: if it still fails the lock survived,
        // and the caller surfaces the locked-file guidance.
        let result: T;
        try {
            result = await write();
        } catch (retryErr) {
            // The retried write failed too. If the close actually freed the lock
            // the file is no longer open in its application, so drop the
            // now-stale "opened by XLIDE" tracking before rethrowing - otherwise
            // a later closeTracked save would act on a file that is not there.
            if (freed) {
                forgetFileOpenedByXlide(filePath);
            }
            throw retryErr;
        }
        if (!viewLost) {
            // Nothing was closed (the lock was someone else's and cleared on its
            // own), so there is no view to put back: never open a file the user
            // did not have open.
        } else if (isFileReopenSuppressed(filePath)) {
            // A caller (F5) is about to reopen this file itself; do not race it.
        } else if (settings.reopenAfterClose) {
            // reopenFileAfterClose re-marks the file as XLIDE-opened on a
            // successful reopen; on failure the prior tracking is left intact so a
            // later closeTracked save can still free the lock. We must NOT forget
            // here; doing so before/around a failed reopen would strand tracking.
            const readOnly = resolveReopenReadOnly(settings.reopenMode, wasReadOnly);
            await reopenFileAfterClose(filePath, readOnly, log);
        } else {
            // Intentionally left closed: the file is no longer open in its application.
            forgetFileOpenedByXlide(filePath);
        }
        return result;
    }
}
