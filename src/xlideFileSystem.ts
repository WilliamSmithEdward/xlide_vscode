import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';

export const XLIDE_SCHEME = 'xlide-vba';

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

    constructor(private readonly _bridge: PythonBridge) {}

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
        const { xlsmPath, moduleName } = decodeModuleUri(uri);
        const result = await this._bridge.call<{ source: string }>(
            'readModule',
            { path: xlsmPath, module: moduleName },
        );
        return Buffer.from(result.source, 'utf-8');
    }

    async writeFile(
        uri: vscode.Uri,
        content: Uint8Array,
        _options: { create: boolean; overwrite: boolean },
    ): Promise<void> {
        const { xlsmPath, moduleName } = decodeModuleUri(uri);
        const source = Buffer.from(content).toString('utf-8');
        await this._bridge.call('writeModule', {
            path: xlsmPath,
            module: moduleName,
            source,
        });
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
