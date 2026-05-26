import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';

export type VbaSymbolKind = 'Sub' | 'Function' | 'PropertyGet' | 'PropertyLet' | 'PropertySet';

export interface VbaSymbol {
    name: string;
    kind: VbaSymbolKind;
    /** 0-based start line within the module source. */
    line: number;
    /** 0-based character where the identifier begins. */
    column: number;
    /** Length of the identifier. */
    length: number;
    /** 0-based start line of the procedure body / declaration. */
    startLine: number;
    /** 0-based end line of the procedure body (inclusive of End <kind>). */
    endLine: number;
    isPublic: boolean;
}

export interface VbaModuleSymbols {
    moduleName: string;
    symbols: VbaSymbol[];
    /** Cached module source used to build the symbols. */
    source: string;
}

interface CachedWorkbook {
    /** moduleName -> module symbols */
    modules: Map<string, VbaModuleSymbols>;
}

const PROC_RE = /^([ \t]*)(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([A-Za-z_][A-Za-z0-9_]*)/i;
const END_RE = /^[ \t]*End\s+(Sub|Function|Property)\b/i;

function kindFromRaw(raw: string): VbaSymbolKind {
    const normalized = raw.replace(/\s+/g, '').toLowerCase();
    if (normalized === 'sub') { return 'Sub'; }
    if (normalized === 'function') { return 'Function'; }
    if (normalized === 'propertyget') { return 'PropertyGet'; }
    if (normalized === 'propertylet') { return 'PropertyLet'; }
    return 'PropertySet';
}

/**
 * Parses VBA module source into a list of procedure symbols.
 * Lightweight regex-based parser; good enough for navigation/rename.
 */
export function parseVbaModule(source: string): VbaSymbol[] {
    const lines = source.split(/\r\n|\r|\n/);
    const symbols: VbaSymbol[] = [];
    let current: VbaSymbol | undefined;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const procMatch = PROC_RE.exec(line);
        if (procMatch) {
            // Close any previous symbol that didn't have a matching End
            if (current) {
                current.endLine = i - 1;
                symbols.push(current);
            }
            const visibility = (procMatch[2] ?? '').toLowerCase();
            const rawKind = procMatch[3];
            const name = procMatch[4];
            const nameIdx = line.indexOf(name, procMatch.index + procMatch[0].length - name.length);
            current = {
                name,
                kind: kindFromRaw(rawKind),
                line: i,
                column: nameIdx >= 0 ? nameIdx : 0,
                length: name.length,
                startLine: i,
                endLine: i,
                isPublic: visibility !== 'private',
            };
            continue;
        }
        if (current && END_RE.test(line)) {
            current.endLine = i;
            symbols.push(current);
            current = undefined;
        }
    }
    if (current) {
        current.endLine = lines.length - 1;
        symbols.push(current);
    }
    return symbols;
}

/**
 * Workbook-scoped VBA symbol index. Lazily loads modules on first query;
 * callers can invalidate single modules or whole workbooks after edits.
 */
export class VbaSymbolIndex implements vscode.Disposable {
    private _cache = new Map<string, CachedWorkbook>();
    private _emitter = new vscode.EventEmitter<{ xlsmPath: string; moduleName?: string }>();
    readonly onDidChange = this._emitter.event;

    constructor(private readonly _bridge: PythonBridge) {}

    /** Invalidate one module (or the whole workbook when moduleName is omitted). */
    invalidate(xlsmPath: string, moduleName?: string): void {
        const key = this._key(xlsmPath);
        const wb = this._cache.get(key);
        if (!wb) { return; }
        if (moduleName === undefined) {
            this._cache.delete(key);
        } else {
            wb.modules.delete(moduleName);
        }
        this._emitter.fire({ xlsmPath, moduleName });
    }

    invalidateAll(): void {
        this._cache.clear();
        this._emitter.fire({ xlsmPath: '' });
    }

    /** Returns the parsed symbols for a single module, loading on demand. */
    async getModule(xlsmPath: string, moduleName: string): Promise<VbaModuleSymbols> {
        const key = this._key(xlsmPath);
        let wb = this._cache.get(key);
        if (!wb) {
            wb = { modules: new Map() };
            this._cache.set(key, wb);
        }
        let mod = wb.modules.get(moduleName);
        if (!mod) {
            const result = await this._bridge.call<{ source: string }>(
                'readModule',
                { path: xlsmPath, module: moduleName },
            );
            mod = {
                moduleName,
                source: result.source,
                symbols: parseVbaModule(result.source),
            };
            wb.modules.set(moduleName, mod);
        }
        return mod;
    }

    /** Returns the parsed symbols for every module in the workbook. */
    async getAllModules(xlsmPath: string): Promise<VbaModuleSymbols[]> {
        const moduleList = await this._bridge.call<Array<{ name: string; type: string }>>(
            'listModules',
            { path: xlsmPath },
        );
        const out: VbaModuleSymbols[] = [];
        for (const entry of moduleList) {
            try {
                out.push(await this.getModule(xlsmPath, entry.name));
            } catch {
                // Skip modules that fail to read; index is best-effort.
            }
        }
        return out;
    }

    /**
     * Refreshes a single module's source from disk and re-parses it.
     * Useful immediately after a write so the index reflects the new content.
     */
    async refreshModule(xlsmPath: string, moduleName: string): Promise<VbaModuleSymbols> {
        this.invalidate(xlsmPath, moduleName);
        return this.getModule(xlsmPath, moduleName);
    }

    dispose(): void {
        this._cache.clear();
        this._emitter.dispose();
    }

    private _key(xlsmPath: string): string {
        return process.platform === 'win32' ? xlsmPath.toLowerCase() : xlsmPath;
    }
}
