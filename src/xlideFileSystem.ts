import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import type { LiveShareIntegration } from './liveShare';
import { decodeRemoteModuleUri, encodeRemoteModuleUri } from './liveShare';

export const XLIDE_SCHEME = 'xlide-vba';
export const XLIDE_LIVESHARE_AUTHORITY = 'liveshare';

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

    constructor(private readonly _bridge: PythonBridge) {}

    /** Attach the Live Share integration so remote xlide-vba://liveshare/... URIs are routed via RPC. */
    setLiveShare(liveShare: LiveShareIntegration): void {
        this._liveShare = liveShare;
        liveShare.onRemoteFileChanged = (workbookId, moduleName) => {
            const uri = encodeRemoteModuleUri(workbookId, moduleName);
            this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
        };
    }

    // ------------------------------------------------------------------
    // Required by FileSystemProvider but not meaningful for our use case
    // ------------------------------------------------------------------

    watch(): vscode.Disposable {
        return new vscode.Disposable(() => { /* no-op */ });
    }

    stat(_uri: vscode.Uri): vscode.FileStat {
        return {
            type: vscode.FileType.File,
            ctime: 0,
            mtime: Date.now(),
            size: 0,
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
            if (!this._liveShare) {
                throw vscode.FileSystemError.Unavailable('XLIDE: Live Share integration not initialized.');
            }
            const { workbookId, moduleName } = decodeRemoteModuleUri(uri);
            const source = await this._liveShare.guestReadModule(workbookId, moduleName);
            return Buffer.from(source, 'utf-8');
        }
        const { xlsmPath, moduleName } = decodeModuleUri(uri);
        try {
            const result = await this._bridge.call<{ source: string }>(
                'readModule',
                { path: xlsmPath, module: moduleName },
            );
            return Buffer.from(result.source, 'utf-8');
        } catch (err) {
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
            if (!this._liveShare) {
                throw vscode.FileSystemError.Unavailable('XLIDE: Live Share integration not initialized.');
            }
            const { workbookId, moduleName } = decodeRemoteModuleUri(uri);
            await this._liveShare.guestWriteModule(workbookId, moduleName, source);
            this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
            return;
        }
        const { xlsmPath, moduleName } = decodeModuleUri(uri);
        try {
            await this._bridge.call('writeModule', {
                path: xlsmPath,
                module: moduleName,
                source,
            });
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            if (isWorkbookLockedError(message)) {
                reportWorkbookLocked(xlsmPath, 'write');
                throw vscode.FileSystemError.Unavailable(
                    `XLIDE: "${path.basename(xlsmPath)}" is open in Excel. Close it and save again.`,
                );
            }
            throw err;
        }
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    }

    // Public method for agent tools to notify that a file has changed
    notifyFileChanged(uri: vscode.Uri): void {
        this._emitter.fire([{ type: vscode.FileChangeType.Changed, uri }]);
    }

    dispose(): void {
        this._emitter.dispose();
    }
}
