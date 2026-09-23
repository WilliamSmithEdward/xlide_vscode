// Who holds a file open, from Windows' Restart Manager: the API behind
// Explorer's "the action can't be completed because the file is open in ...".
//
// XLIDE saves by renaming a temp file over the container, and Windows refuses
// that while anything holds the file. XLIDE used to name the holder from the
// file type - "open in Excel" - which was a guess, and wrong whenever it was
// something else: Word keeping a read-only copy locked, a second Excel,
// OneDrive syncing, a backup agent. Restart Manager names the processes that
// actually have the file open, with no COM involved. It only lists them;
// nothing is asked of the processes it finds.
//
// The lookup is Windows only and best-effort: anywhere it cannot answer, the
// callers fall back to naming the file's application, as they always did.

import { psSingleQuoted, runPowerShell } from './util/powershell';
import { osPlatform } from './util/osPlatform';
import { containerAppNameForPath } from './macroContainerUi';
import { PROJECT_LOCKED_ERROR_RE } from './xlideCommandLog';

/** One process Restart Manager reports holding the file. */
export interface FileLockHolder {
    pid: number;
    /**
     * The name Restart Manager gives the application: its product name, such
     * as "Microsoft Excel", or a service's display name.
     */
    appName: string;
    /** The executable's file name, such as EXCEL.EXE, when it could be read. */
    image?: string;
}

const HOLDERS_SENTINEL = 'XLIDE_LOCK_HOLDERS|';

/**
 * The structures and calls are as RestartManager.h declares them (Windows SDK
 * 10.0.26100): CCH_RM_MAX_APP_NAME 255 and CCH_RM_MAX_SVC_NAME 63, each with
 * room for its terminator; a session key of CCH_RM_SESSION_KEY (32) plus one
 * characters; and RmGetList answering ERROR_MORE_DATA (234) until the array
 * is big enough for every process it found.
 */
const RESTART_MANAGER_CSHARP = [
    'using System;',
    'using System.Collections.Generic;',
    'using System.ComponentModel;',
    'using System.Runtime.InteropServices;',
    'using System.Text;',
    '',
    'public static class XlideLockHolders {',
    '    [StructLayout(LayoutKind.Sequential)]',
    '    struct RM_UNIQUE_PROCESS {',
    '        public int dwProcessId;',
    '        public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;',
    '    }',
    '',
    '    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]',
    '    struct RM_PROCESS_INFO {',
    '        public RM_UNIQUE_PROCESS Process;',
    '        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string strAppName;',
    '        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string strServiceShortName;',
    '        public int ApplicationType;',
    '        public uint AppStatus;',
    '        public uint TSSessionId;',
    '        [MarshalAs(UnmanagedType.Bool)] public bool bRestartable;',
    '    }',
    '',
    '    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]',
    '    static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, StringBuilder strSessionKey);',
    '',
    '    [DllImport("rstrtmgr.dll")]',
    '    static extern int RmEndSession(uint pSessionHandle);',
    '',
    '    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]',
    '    static extern int RmRegisterResources(uint pSessionHandle, uint nFiles, string[] rgsFilenames, uint nApplications, RM_UNIQUE_PROCESS[] rgApplications, uint nServices, string[] rgsServiceNames);',
    '',
    '    [DllImport("rstrtmgr.dll")]',
    '    static extern int RmGetList(uint dwSessionHandle, out uint pnProcInfoNeeded, ref uint pnProcInfo, [In, Out] RM_PROCESS_INFO[] rgAffectedApps, ref uint lpdwRebootReasons);',
    '',
    '    public static string[] Find(string path) {',
    '        uint session;',
    '        int rc = RmStartSession(out session, 0, new StringBuilder(33));',
    '        if (rc != 0) { throw new Win32Exception(rc); }',
    '        try {',
    '            rc = RmRegisterResources(session, 1, new string[] { path }, 0, null, 0, null);',
    '            if (rc != 0) { throw new Win32Exception(rc); }',
    '            RM_PROCESS_INFO[] found = new RM_PROCESS_INFO[0];',
    '            uint needed = 0;',
    '            uint count = 0;',
    '            uint reasons = 0;',
    '            for (int attempt = 0; ; attempt++) {',
    '                rc = RmGetList(session, out needed, ref count, found, ref reasons);',
    '                if (rc == 0) { break; }',
    '                if (rc != 234 || attempt == 4) { throw new Win32Exception(rc); }',
    '                found = new RM_PROCESS_INFO[needed];',
    '                count = needed;',
    '            }',
    '            List<string> lines = new List<string>();',
    '            for (int i = 0; i < count; i++) {',
    '                lines.Add(found[i].Process.dwProcessId + "\\t" + found[i].strAppName);',
    '            }',
    '            return lines.ToArray();',
    '        }',
    '        finally {',
    '            RmEndSession(session);',
    '        }',
    '    }',
    '}',
];

/**
 * PowerShell listing the processes that hold `filePath` open, as one
 * `XLIDE_LOCK_HOLDERS|<json array>` line: each entry's pid, the application
 * name Restart Manager gives, and the executable's file name. Exported for
 * the tests.
 */
export function buildLockHoldersScript(filePath: string): string {
    return [
        '$ErrorActionPreference = "Stop"',
        `$targetPath = ${psSingleQuoted(filePath)}`,
        'if (-not ([System.Management.Automation.PSTypeName]"XlideLockHolders").Type) {',
        "Add-Type -TypeDefinition @'",
        ...RESTART_MANAGER_CSHARP,
        "'@",
        '}',
        '$holders = @()',
        'foreach ($entry in [XlideLockHolders]::Find($targetPath)) {',
        '  $parts = $entry.Split("`t", 2)',
        '  $holderId = [int]$parts[0]',
        '  $image = ""',
        '  try { $image = [System.IO.Path]::GetFileName((Get-Process -Id $holderId -ErrorAction Stop).Path) } catch { }',
        '  $holders += [pscustomobject]@{ pid = $holderId; appName = $parts[1]; image = $image }',
        '}',
        `[Console]::Out.WriteLine("${HOLDERS_SENTINEL}" + (ConvertTo-Json -InputObject @($holders) -Compress))`,
    ].join('\n');
}

/**
 * The holders a lookup printed, or undefined when it printed no answer. An
 * empty list is an answer: nothing holds the file now.
 */
export function parseLockHolders(stdoutLines: readonly string[]): FileLockHolder[] | undefined {
    const line = stdoutLines.find((candidate) => candidate.startsWith(HOLDERS_SENTINEL));
    if (!line) {
        return undefined;
    }
    let parsed: unknown;
    try {
        parsed = JSON.parse(line.slice(HOLDERS_SENTINEL.length));
    } catch {
        return undefined;
    }
    if (!Array.isArray(parsed)) {
        return undefined;
    }
    return parsed.flatMap((entry): FileLockHolder[] => {
        if (!entry || typeof entry !== 'object' || typeof (entry as { pid?: unknown }).pid !== 'number') {
            return [];
        }
        const { pid, appName, image } = entry as { pid: number; appName?: unknown; image?: unknown };
        return [{
            pid,
            appName: typeof appName === 'string' ? appName.trim() : '',
            ...(typeof image === 'string' && image ? { image } : {}),
        }];
    });
}

/** Long enough for PowerShell to start and compile the lookup on a busy machine. */
const LOOKUP_TIMEOUT_MS = 10_000;

/**
 * The processes holding the file open, from Restart Manager. Undefined when
 * that could not be found out: not Windows, or the lookup failed.
 */
export async function findFileLockHolders(
    filePath: string,
    log: (message: string) => void = () => undefined,
): Promise<FileLockHolder[] | undefined> {
    if (osPlatform() !== 'win32') {
        return undefined;
    }
    try {
        const result = await runPowerShell({ script: buildLockHoldersScript(filePath), timeoutMs: LOOKUP_TIMEOUT_MS }).result;
        const holders = parseLockHolders(result.stdoutLines);
        if (!holders) {
            log(`[lockHolders] no answer for ${filePath}: ${result.stderrLines.join(' ') || `exit ${result.code}`}`);
        }
        return holders;
    } catch (err) {
        log(`[lockHolders] lookup failed for ${filePath}: ${err instanceof Error ? err.message : String(err)}`);
        return undefined;
    }
}

/** "Microsoft Excel (EXCEL.EXE, process 4242)", one per holder. */
export function describeLockHolders(holders: readonly FileLockHolder[]): string {
    const one = (holder: FileLockHolder): string => {
        const name = holder.appName || holder.image || 'a program';
        const detail = [holder.image && holder.image !== name ? holder.image : undefined, `process ${holder.pid}`]
            .filter((part): part is string => Boolean(part))
            .join(', ');
        return `${name} (${detail})`;
    };
    const named = holders.map(one);
    return named.length <= 1
        ? named.join('')
        : `${named.slice(0, -1).join(', ')} and ${named[named.length - 1]}`;
}

/** Holders found for a lock error, so the surface that reports it need not look again. */
const holdersByError = new WeakMap<Error, FileLockHolder[]>();

/**
 * Names the processes holding the file in a write's lock error, in its
 * message and for {@link lockHoldersOf}. An agent tool passes the message on
 * as it is, so the name reaches the agent too. Anything that is not a lock
 * error, or a lock nothing is found holding any more, comes back unchanged.
 */
export async function annotateLockError(
    err: unknown,
    filePath: string,
    log?: (message: string) => void,
): Promise<unknown> {
    if (!(err instanceof Error) || !PROJECT_LOCKED_ERROR_RE.test(err.message) || holdersByError.has(err)) {
        return err;
    }
    const holders = await findFileLockHolders(filePath, log);
    if (holders && holders.length > 0) {
        holdersByError.set(err, holders);
        err.message = `${err.message} (held open by ${describeLockHolders(holders)})`;
        log?.(`[lockHolders] ${filePath} is held open by ${describeLockHolders(holders)}`);
    }
    return err;
}

/** The holders {@link annotateLockError} found for this error, if it looked. */
export function lockHoldersOf(err: unknown): FileLockHolder[] | undefined {
    return err instanceof Error ? holdersByError.get(err) : undefined;
}

/**
 * What holds the file, for a sentence of the form `"Book.xlsm" is open in
 * ...`: the processes Restart Manager names when it can, and otherwise the
 * application the file belongs to, which is what XLIDE used to say. The
 * holders already found for `err` are used rather than looked up again.
 */
export async function lockedFileHolderText(
    filePath: string,
    err?: unknown,
): Promise<{ text: string; known: boolean }> {
    const holders = lockHoldersOf(err) ?? await findFileLockHolders(filePath);
    return holders && holders.length > 0
        ? { text: describeLockHolders(holders), known: true }
        : { text: containerAppNameForPath(filePath), known: false };
}
