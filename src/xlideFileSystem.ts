import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import type { LiveShareIntegration } from './liveShare';
import { decodeRemoteModuleUri, encodeRemoteModuleUri } from './liveShare';
import { errorCategoryForSupportLog } from './xlideCommandLog';
import { formatChangeSummary, recordXlideWriteAudit } from './xlideWriteAudit';
import { startPerformanceTrace } from './performanceTrace';

export const XLIDE_SCHEME = 'xlide-vba';
export const XLIDE_VBA_LANGUAGE_ID = 'xlide-vba';
export const XLIDE_LIVESHARE_AUTHORITY = 'liveshare';

export function workbookIdentityKey(
    workbookPath: string,
    platform: NodeJS.Platform = process.platform,
): string {
    return platform === 'win32' ? workbookPath.toLowerCase() : workbookPath;
}

export function sameWorkbookPath(
    a: string,
    b: string,
    platform: NodeJS.Platform = process.platform,
): boolean {
    return workbookIdentityKey(a, platform) === workbookIdentityKey(b, platform);
}

export function moduleIdentityKey(moduleName: string): string {
    return moduleName.toLowerCase();
}

/**
 * Tracks workbook paths for which the signature-dropped notice has already
 * been shown this session, so the user sees it at most once per file.
 */
const _sigWarnedPaths = new Set<string>();

/**
 * Show a one-time warning when a VBA digital signature was invalidated by a
 * save.  Safe to call on every write — suppressed after the first occurrence
 * per workbook path per session.
 */
export function notifySignatureDropped(filePath: string, signatureDropped: boolean): void {
    const key = workbookIdentityKey(filePath);
    if (!signatureDropped || _sigWarnedPaths.has(key)) { return; }
    _sigWarnedPaths.add(key);
    void vscode.window.showWarningMessage(
        `XLIDE: "${path.basename(filePath)}" had a VBA digital signature that was invalidated by this edit. ` +
        `Re-sign the workbook externally to restore trust.`,
    );
}

/**
 * Heuristic: does this error string look like a Windows file-sharing violation
 * caused by Excel having the workbook open?
 */
function isWorkbookLockedError(message: string): boolean {
    return /WinError\s*32|being used by another process|sharing violation|Permission denied|PermissionError/i
        .test(message);
}

function reportWorkbookLocked(xlsmPath: string, op: 'read' | 'write'): void {
    const name = path.basename(xlsmPath);
    const verb = op === 'read' ? 'open' : 'save';
    void vscode.window.showWarningMessage(
        `XLIDE: Cannot ${verb} "${name}" - it appears to be open in Excel. Close the workbook and try again.`,
        'Retry',
        'Reveal File',
    ).then((choice) => {
        if (choice === 'Retry') {
            void vscode.commands.executeCommand('workbench.action.files.revert');
        } else if (choice === 'Reveal File') {
            void vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(xlsmPath));
        }
    });
}

/**
 * Encodes a (xlsmPath, moduleName) pair into a virtual URI.
 * URI form: xlide-vba:///C:/path/to/workbook.xlsm/ModuleName.bas
 */
export function encodeModuleUri(xlsmPath: string, moduleName: string): vscode.Uri {
    const forward = xlsmPath.replace(/\\/g, '/');
    const base = forward.startsWith('/') ? forward : `/${forward}`;
    return vscode.Uri.parse(
        `${XLIDE_SCHEME}:${base}/${encodeURIComponent(moduleName)}.bas`,
        true,
    );
}

/**
 * Decodes a virtual URI back to (xlsmPath, moduleName).
 */
export function decodeModuleUri(uri: vscode.Uri): { xlsmPath: string; moduleName: string } {
    const p = uri.path;
    // Match the .xlsm (or .xlsb/.xlam) boundary in the path
    const match = p.match(/^(.*\.xl(?:sm|sb|am))\/([^/]+)\.bas$/i);
    if (!match) {
        throw new Error(`Cannot decode xlide-vba URI: ${uri.toString()}`);
    }
    let rawPath = match[1]; // e.g. /C:/Users/.../workbook.xlsm
    const moduleName = decodeURIComponent(match[2]);

    // On Windows, the leading slash before the drive letter is artificial
    if (process.platform === 'win32' && /^\/[A-Za-z]:/.test(rawPath)) {
        rawPath = rawPath.slice(1);
    }
    const xlsmPath = rawPath.replace(/\//g, path.sep);
    return { xlsmPath, moduleName };
}

/**
 * Virtual FileSystemProvider for the xlide-vba:// scheme.
 *
 * - readFile  -> calls Python bridge readModule
 * - writeFile -> calls Python bridge writeModule (saves the .xlsm in place)
 * - All other mutation operations are rejected.
 */
export class XlideFileSystemProvider
    implements vscode.FileSystemProvider, vscode.Disposable
{
    private _emitter = new vscode.EventEmitter<vscode.FileChangeEvent[]>();
    readonly onDidChangeFile = this._emitter.event;

    private _liveShare: LiveShareIntegration | undefined;
    private _clock = Date.now();
    private readonly _stats = new Map<string, { ctime: number; mtime: number; size: number; workbookKey?: string }>();
    /** Last known real workbook file mtime per workbook identity key. */
    private readonly _workbookMtimes = new Map<string, number>();

    constructor(private readonly _bridge: PythonBridge) {}

    /** Attach the Live Share integration so remote xlide-vba://liveshare/... URIs are routed via RPC. */
    setLiveShare(liveShare: LiveShareIntegration): void {
        this._liveShare = liveShare;
        liveShare.onRemoteFileChanged = (workbookId, moduleName) => {
            const uri = encodeRemoteModuleUri(workbookId, moduleName);
            this.bumpStat(uri);
            this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
        };
        liveShare.onHostModuleWritten = (workbookPath, moduleName, signatureDropped) => {
            notifySignatureDropped(workbookPath, signatureDropped);
            this.notifyFileChanged(encodeModuleUri(workbookPath, moduleName));
        };
    }

    // ------------------------------------------------------------------
    // Required by FileSystemProvider but not meaningful for our use case
    // ------------------------------------------------------------------

    watch(): vscode.Disposable {
        return new vscode.Disposable(() => { /* no-op */ });
    }

    stat(uri: vscode.Uri): vscode.FileStat {
        const state = this.ensureStat(uri);
        this.syncWithWorkbookFile(state, uri);
        return {
            type: vscode.FileType.File,
            ctime: state.ctime,
            mtime: state.mtime,
            size: state.size,
        };
    }

    readDirectory(_uri: vscode.Uri): [string, vscode.FileType][] {
        return [];
    }

    createDirectory(_uri: vscode.Uri): void {
        throw vscode.FileSystemError.NoPermissions('XLIDE: createDirectory not supported');
    }

    delete(_uri: vscode.Uri): void {
        throw vscode.FileSystemError.NoPermissions('XLIDE: delete not supported via file system');
    }

    rename(_oldUri: vscode.Uri, _newUri: vscode.Uri): void {
        throw vscode.FileSystemError.NoPermissions('XLIDE: rename not supported via file system');
    }

    // ------------------------------------------------------------------
    // Core read/write
    // ------------------------------------------------------------------

    async readFile(uri: vscode.Uri): Promise<Uint8Array> {
        if (uri.authority === XLIDE_LIVESHARE_AUTHORITY) {
            const trace = startPerformanceTrace('filesystem.readFile', 'liveshare');
            if (!this._liveShare) {
                trace.end('failed', 'liveshare');
                throw vscode.FileSystemError.Unavailable('XLIDE: Live Share integration not initialized.');
            }
            const { workbookId, moduleName } = decodeRemoteModuleUri(uri);
            try {
                const source = await this._liveShare.guestReadModule(workbookId, moduleName);
                const bytes = Buffer.from(source, 'utf-8');
                this.updateSize(uri, bytes.byteLength);
                trace.end('ok', moduleName);
                return bytes;
            } catch (err) {
                trace.end('failed', moduleName);
                throw err;
            }
        }
        const { xlsmPath, moduleName } = decodeModuleUri(uri);
        const trace = startPerformanceTrace('filesystem.readFile', moduleName);
        try {
            const result = await this._bridge.call<{ source: string }>(
                'readModule',
                { path: xlsmPath, module: moduleName },
            );
            const bytes = Buffer.from(result.source, 'utf-8');
            this.updateSize(uri, bytes.byteLength);
            trace.end('ok', moduleName);
            return bytes;
        } catch (err) {
            trace.end('failed', moduleName);
            const message = err instanceof Error ? err.message : String(err);
            if (isWorkbookLockedError(message)) {
                reportWorkbookLocked(xlsmPath, 'read');
                throw vscode.FileSystemError.Unavailable(
                    `XLIDE: "${path.basename(xlsmPath)}" is open in Excel. Close it and click Retry.`,
                );
            }
            throw err;
        }
    }

    async writeFile(
        uri: vscode.Uri,
        content: Uint8Array,
        _options: { create: boolean; overwrite: boolean },
    ): Promise<void> {
        const source = Buffer.from(content).toString('utf-8');
        if (uri.authority === XLIDE_LIVESHARE_AUTHORITY) {
            const trace = startPerformanceTrace('filesystem.writeFile', 'liveshare');
            if (!this._liveShare) {
                trace.end('failed', 'liveshare');
                throw vscode.FileSystemError.Unavailable('XLIDE: Live Share integration not initialized.');
            }
            const { workbookId, moduleName } = decodeRemoteModuleUri(uri);
            try {
                await this._liveShare.guestWriteModule(workbookId, moduleName, source);
                const summary = formatChangeSummary({
                    operation: 'Save module',
                    changed: [moduleName],
                });
                recordXlideWriteAudit({
                    timestamp: new Date().toISOString(),
                    command: 'xlide.editorSave',
                    operation: 'write-module',
                    outcome: 'succeeded',
                    moduleName,
                    targetPath: workbookId,
                    summary,
                });
                this.markChanged(uri, Buffer.byteLength(source, 'utf-8'));
                this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
                trace.end('ok', moduleName);
                return;
            } catch (err) {
                trace.end('failed', moduleName);
                throw err;
            }
        }
        const { xlsmPath, moduleName } = decodeModuleUri(uri);
        const trace = startPerformanceTrace('filesystem.writeFile', moduleName);
        try {
            const result = await this._bridge.call<{ ok: boolean; signatureDropped: boolean }>(
                'writeModule',
                {
                    path: xlsmPath,
                    module: moduleName,
                    source,
                },
            );
            notifySignatureDropped(xlsmPath, result.signatureDropped);
            const summary = formatChangeSummary({
                operation: 'Save module',
                changed: [moduleName],
            });
            recordXlideWriteAudit({
                timestamp: new Date().toISOString(),
                command: 'xlide.editorSave',
                operation: 'write-module',
                outcome: 'succeeded',
                workbookPath: xlsmPath,
                moduleName,
                summary,
            });
        } catch (err) {
            trace.end('failed', moduleName);
            const message = err instanceof Error ? err.message : String(err);
            recordXlideWriteAudit({
                timestamp: new Date().toISOString(),
                command: 'xlide.editorSave',
                operation: 'write-module',
                outcome: 'failed',
                workbookPath: xlsmPath,
                moduleName,
                summary: 'Save module: 0 changed, 1 failed',
                errorCategory: errorCategoryForSupportLog(err),
            });
            if (isWorkbookLockedError(message)) {
                reportWorkbookLocked(xlsmPath, 'write');
                throw vscode.FileSystemError.Unavailable(
                    `XLIDE: "${path.basename(xlsmPath)}" is open in Excel. Close it and save again.`,
                );
            }
            throw err;
        }
        this.markChanged(uri, Buffer.byteLength(source, 'utf-8'));
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
        trace.end('ok', moduleName);
    }

    // Public method for out-of-band mutators (agent tools, commands, Live Share)
    // to notify that a module changed so open editors reload and stats refresh
    notifyFileChanged(uri: vscode.Uri): void {
        this.markChanged(uri);
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    }

    dispose(): void {
        this._emitter.dispose();
    }

    private ensureStat(uri: vscode.Uri): { ctime: number; mtime: number; size: number; workbookKey?: string } {
        const key = this.statKey(uri);
        const existing = this._stats.get(key);
        if (existing) {
            return existing;
        }
        const real = this.workbookFileMtime(uri);
        const now = real?.mtime ?? this.nextTimestamp();
        const created = { ctime: now, mtime: now, size: 0, workbookKey: real?.workbookKey };
        this._stats.set(key, created);
        if (real && !this._workbookMtimes.has(real.workbookKey)) {
            this._workbookMtimes.set(real.workbookKey, real.mtime);
        }
        return created;
    }

    /**
     * Real mtime of the backing workbook file, keyed by workbook identity.
     * Module mtimes are derived from it so VS Code's save-conflict detection
     * sees out-of-band changes (Excel VBE edits, module sync, agent writes).
     * Undefined for Live Share URIs and paths that cannot be statted.
     */
    private workbookFileMtime(uri: vscode.Uri): { workbookKey: string; mtime: number } | undefined {
        if (uri.authority === XLIDE_LIVESHARE_AUTHORITY) {
            return undefined;
        }
        try {
            const { xlsmPath } = decodeModuleUri(uri);
            return {
                workbookKey: workbookIdentityKey(xlsmPath),
                mtime: Math.floor(fs.statSync(xlsmPath).mtimeMs),
            };
        } catch {
            return undefined;
        }
    }

    /**
     * Detect workbook file changes made outside this provider. When the real
     * file mtime moved past the last mtime this provider produced or observed,
     * any module may differ, so every cached module stat of that workbook is
     * bumped to the new file mtime.
     */
    private syncWithWorkbookFile(state: { mtime: number; workbookKey?: string }, uri: vscode.Uri): void {
        const real = this.workbookFileMtime(uri);
        if (!real) {
            return;
        }
        state.workbookKey = real.workbookKey;
        const baseline = this._workbookMtimes.get(real.workbookKey);
        if (baseline !== undefined && real.mtime !== baseline) {
            for (const entry of this._stats.values()) {
                if (entry.workbookKey === real.workbookKey) {
                    entry.mtime = real.mtime;
                }
            }
        }
        this._workbookMtimes.set(real.workbookKey, real.mtime);
    }

    private updateSize(uri: vscode.Uri, size: number): void {
        this.ensureStat(uri).size = size;
    }

    private bumpStat(uri: vscode.Uri, size?: number): void {
        const state = this.ensureStat(uri);
        state.mtime = this.nextTimestamp(state.mtime);
        if (size !== undefined) {
            state.size = size;
        }
    }

    private markChanged(uri: vscode.Uri, size?: number): void {
        const state = this.ensureStat(uri);
        const real = this.workbookFileMtime(uri);
        if (real) {
            state.workbookKey = real.workbookKey;
            state.mtime = real.mtime;
            this._workbookMtimes.set(real.workbookKey, real.mtime);
        } else {
            state.mtime = this.nextTimestamp(state.mtime);
        }
        if (size !== undefined) {
            state.size = size;
        }
    }

    private nextTimestamp(after = 0): number {
        const now = Date.now();
        this._clock = Math.max(this._clock + 1, now, after + 1);
        return this._clock;
    }

    private statKey(uri: vscode.Uri): string {
        return uri.toString();
    }
}
