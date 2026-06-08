import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import type { EventHandlerDocumentType } from './analyzer/completion/eventHandlers';
import { moduleIdentityKey, workbookIdentityKey } from './xlideFileSystem';

export type VbaSymbolKind =
    | 'Sub' | 'Function' | 'PropertyGet' | 'PropertyLet' | 'PropertySet'
    | 'Const' | 'Enum' | 'Type';

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
    /** Host module type from listModules (standard/class/document/userform). */
    type?: string;
    /** Excel document subtype from listModules when the bridge can prove it. */
    documentType?: EventHandlerDocumentType;
}

interface CachedWorkbook {
    /** moduleName -> module symbols */
    modules: Map<string, VbaModuleSymbols>;
    /** Cached workbook module list from the bridge. */
    moduleList?: VbaModuleEntry[];
    moduleListLoadedAt?: number;
}

interface VbaModuleEntry {
    name: string;
    type: string;
    documentType?: EventHandlerDocumentType;
}

interface VbaModuleSourceEntry extends VbaModuleEntry {
    source: string;
}

const MODULE_LIST_CACHE_TTL_MS = 5_000;
const WORKBOOK_INDEX_YIELD_EVERY_MODULES = 8;
const PROC_RE = /^([ \t]*)(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+([A-Za-z_][A-Za-z0-9_]*)/i;
const END_RE = /^[ \t]*End\s+(Sub|Function|Property)\b/i;
// Declarations that appear as single-line symbols (no End block).
const DECL_RE = /^[ \t]*(?:(Public|Private|Friend|Global)\s+)?(?:Const\s+([A-Za-z_][A-Za-z0-9_]*)|(?:Enum|Type)\s+([A-Za-z_][A-Za-z0-9_]*))/i;

function kindFromRaw(raw: string): VbaSymbolKind {
    const normalized = raw.replace(/\s+/g, '').toLowerCase();
    if (normalized === 'sub') { return 'Sub'; }
    if (normalized === 'function') { return 'Function'; }
    if (normalized === 'propertyget') { return 'PropertyGet'; }
    if (normalized === 'propertylet') { return 'PropertyLet'; }
    return 'PropertySet';
}

/**
 * Parses VBA module source into a list of procedure and declaration symbols.
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
            if (current) { current.endLine = i - 1; symbols.push(current); }
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
            continue;
        }
        // Single-line declarations: Const, Enum <name>, Type <name>
        if (!current) {
            const declMatch = DECL_RE.exec(line);
            if (declMatch) {
                const visibility = (declMatch[1] ?? '').toLowerCase();
                const constName = declMatch[2];   // Const path
                const blockName = declMatch[3];   // Enum / Type path
                // Determine kind by checking which keyword was matched
                let kind: VbaSymbolKind;
                let name: string;
                if (constName) {
                    kind = 'Const'; name = constName;
                } else {
                    const keyword = line.trim().replace(/^(?:Public|Private|Friend|Global)\s+/i, '').split(/\s+/)[0];
                    kind = /^Enum$/i.test(keyword) ? 'Enum' : 'Type';
                    name = blockName;
                }
                const col = line.indexOf(name);
                // Const is a single-line symbol; Enum/Type span to End Enum/Type.
                // Treat all as point symbols (endLine = startLine) — VS Code outline
                // only needs the declaration line for breadcrumb navigation.
                symbols.push({
                    name,
                    kind,
                    line: i, column: col >= 0 ? col : 0,
                    length: name.length,
                    startLine: i, endLine: i,
                    isPublic: visibility !== 'private',
                });
            }
        }
    }
    if (current) {
        current.endLine = lines.length - 1;
        symbols.push(current);
    }
    return symbols;
}

async function yieldToExtensionHost(): Promise<void> {
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
}

/**
 * Workbook-scoped VBA symbol index. Lazily loads modules on first query;
 * callers can invalidate single modules or whole workbooks after edits.
 */
export class VbaSymbolIndex implements vscode.Disposable {
    private _cache = new Map<string, CachedWorkbook>();
    private _moduleReads = new Map<string, Promise<VbaModuleSymbols>>();
    private _moduleListReads = new Map<string, Promise<VbaModuleEntry[]>>();
    private _allModuleReads = new Map<string, Promise<VbaModuleSymbols[]>>();
    private _moduleGenerations = new Map<string, number>();
    private _emitter = new vscode.EventEmitter<{ xlsmPath: string; moduleName?: string }>();
    readonly onDidChange = this._emitter.event;

    constructor(private readonly _bridge: PythonBridge) {}

    /** Invalidate one module (or the whole workbook when moduleName is omitted). */
    invalidate(xlsmPath: string, moduleName?: string): void {
        const key = workbookIdentityKey(xlsmPath);
        const wb = this._cache.get(key);
        if (moduleName === undefined) {
            this._cache.delete(key);
            this.deleteWorkbookInflight(key);
        } else {
            const moduleKey = moduleIdentityKey(moduleName);
            wb?.modules.delete(moduleKey);
            const requestKey = this.moduleRequestKey(key, moduleKey);
            this.bumpModuleGeneration(requestKey);
            this._moduleReads.delete(requestKey);
            this._allModuleReads.delete(key);
        }
        this._emitter.fire({ xlsmPath, moduleName });
    }

    invalidateAll(): void {
        this._cache.clear();
        this._moduleReads.clear();
        this._moduleListReads.clear();
        this._allModuleReads.clear();
        this._moduleGenerations.clear();
        this._emitter.fire({ xlsmPath: '' });
    }

    /** Returns the parsed symbols for a single module, loading on demand. */
    async getModule(xlsmPath: string, moduleName: string): Promise<VbaModuleSymbols> {
        const key = workbookIdentityKey(xlsmPath);
        const wb = this.workbook(key);
        const moduleKey = moduleIdentityKey(moduleName);
        const cached = wb.modules.get(moduleKey);
        if (cached) { return cached; }

        const requestKey = this.moduleRequestKey(key, moduleKey);
        const existingRead = this._moduleReads.get(requestKey);
        if (existingRead) { return existingRead; }

        const generation = this.moduleGeneration(requestKey);
        const promise = (async () => {
            const result = await this._bridge.call<{ source: string }>(
                'readModule',
                { path: xlsmPath, module: moduleName },
            );
            const mod: VbaModuleSymbols = {
                moduleName,
                source: result.source,
                symbols: parseVbaModule(result.source),
            };
            if (this.moduleGeneration(requestKey) !== generation) {
                return wb.modules.get(moduleKey) ?? mod;
            }
            wb.modules.set(moduleKey, mod);
            return mod;
        })();
        this._moduleReads.set(requestKey, promise);
        promise.then(
            () => {
                if (this._moduleReads.get(requestKey) === promise) {
                    this._moduleReads.delete(requestKey);
                }
            },
            () => {
                if (this._moduleReads.get(requestKey) === promise) {
                    this._moduleReads.delete(requestKey);
                }
            },
        );
        return promise;
    }

    /** Returns the parsed symbols for every module in the workbook. */
    async getAllModules(xlsmPath: string): Promise<VbaModuleSymbols[]> {
        const key = workbookIdentityKey(xlsmPath);
        const existingRead = this._allModuleReads.get(key);
        if (existingRead) { return existingRead; }

        const promise = (async () => {
            const cached = this.cachedAllModules(key);
            if (cached) {
                return cached;
            }
            const batch = await this.getAllModulesFromBatchRead(xlsmPath, key);
            if (batch) {
                return batch;
            }
            const moduleList = await this.getModuleList(xlsmPath, key);
            const out: VbaModuleSymbols[] = [];
            for (const entry of moduleList) {
                try {
                    const mod = await this.getModule(xlsmPath, entry.name);
                    mod.type = entry.type;
                    mod.documentType = entry.documentType;
                    out.push(mod);
                } catch {
                    // Skip modules that fail to read; index is best-effort.
                }
            }
            return out;
        })();
        this._allModuleReads.set(key, promise);
        promise.then(
            () => {
                if (this._allModuleReads.get(key) === promise) {
                    this._allModuleReads.delete(key);
                }
            },
            () => {
                if (this._allModuleReads.get(key) === promise) {
                    this._allModuleReads.delete(key);
                }
            },
        );
        return promise;
    }

    /**
     * Refreshes a single module's source from disk and re-parses it.
     * Useful immediately after a write so the index reflects the new content.
     */
    async refreshModule(xlsmPath: string, moduleName: string): Promise<VbaModuleSymbols> {
        this.invalidate(xlsmPath, moduleName);
        return this.getModule(xlsmPath, moduleName);
    }

    /**
     * Updates one cached module directly from an already-known source snapshot.
     * This avoids a bridge round-trip after saving a virtual VBA editor buffer.
     */
    updateModuleSource(
        xlsmPath: string,
        moduleName: string,
        source: string,
        metadata: { type?: string; documentType?: EventHandlerDocumentType } = {},
    ): VbaModuleSymbols {
        const key = workbookIdentityKey(xlsmPath);
        const moduleKey = moduleIdentityKey(moduleName);
        const requestKey = this.moduleRequestKey(key, moduleKey);
        this.bumpModuleGeneration(requestKey);
        this._moduleReads.delete(requestKey);
        this._allModuleReads.delete(key);

        const wb = this.workbook(key);
        const existing = wb.modules.get(moduleKey);
        const mod: VbaModuleSymbols = {
            moduleName,
            source,
            symbols: parseVbaModule(source),
            type: metadata.type ?? existing?.type,
            documentType: metadata.documentType ?? existing?.documentType,
        };
        wb.modules.set(moduleKey, mod);
        this._emitter.fire({ xlsmPath, moduleName });
        return mod;
    }

    dispose(): void {
        this._cache.clear();
        this._moduleReads.clear();
        this._moduleListReads.clear();
        this._allModuleReads.clear();
        this._moduleGenerations.clear();
        this._emitter.dispose();
    }

    private async getModuleList(xlsmPath: string, workbookKey: string): Promise<VbaModuleEntry[]> {
        const wb = this.workbook(workbookKey);
        const loadedAt = wb.moduleListLoadedAt ?? 0;
        if (wb.moduleList && Date.now() - loadedAt < MODULE_LIST_CACHE_TTL_MS) {
            return wb.moduleList;
        }

        const existingRead = this._moduleListReads.get(workbookKey);
        if (existingRead) { return existingRead; }

        const promise = (async () => {
            const moduleList = await this._bridge.call<VbaModuleEntry[]>(
                'listModules',
                { path: xlsmPath },
            );
            wb.moduleList = moduleList;
            wb.moduleListLoadedAt = Date.now();
            return moduleList;
        })();
        this._moduleListReads.set(workbookKey, promise);
        promise.then(
            () => {
                if (this._moduleListReads.get(workbookKey) === promise) {
                    this._moduleListReads.delete(workbookKey);
                }
            },
            () => {
                if (this._moduleListReads.get(workbookKey) === promise) {
                    this._moduleListReads.delete(workbookKey);
                }
            },
        );
        return promise;
    }

    private cachedAllModules(workbookKey: string): VbaModuleSymbols[] | undefined {
        const wb = this._cache.get(workbookKey);
        const loadedAt = wb?.moduleListLoadedAt ?? 0;
        if (!wb?.moduleList || Date.now() - loadedAt >= MODULE_LIST_CACHE_TTL_MS) {
            return undefined;
        }
        const modules: VbaModuleSymbols[] = [];
        for (const entry of wb.moduleList) {
            const mod = wb.modules.get(moduleIdentityKey(entry.name));
            if (!mod) {
                return undefined;
            }
            mod.type = entry.type;
            mod.documentType = entry.documentType;
            modules.push(mod);
        }
        return modules;
    }

    private async getAllModulesFromBatchRead(
        xlsmPath: string,
        workbookKey: string,
    ): Promise<VbaModuleSymbols[] | undefined> {
        try {
            const entries = await this._bridge.call<VbaModuleSourceEntry[]>(
                'readModules',
                { path: xlsmPath },
            );
            const wb = this.workbook(workbookKey);
            wb.moduleList = entries.map(({ name, type, documentType }) => ({ name, type, documentType }));
            wb.moduleListLoadedAt = Date.now();
            const out: VbaModuleSymbols[] = [];
            for (const [index, entry] of entries.entries()) {
                const moduleKey = moduleIdentityKey(entry.name);
                const requestKey = this.moduleRequestKey(workbookKey, moduleKey);
                const existing = wb.modules.get(moduleKey);
                if (existing && this.moduleGeneration(requestKey) > 0) {
                    existing.type = entry.type;
                    existing.documentType = entry.documentType;
                    out.push(existing);
                    if ((index + 1) % WORKBOOK_INDEX_YIELD_EVERY_MODULES === 0) {
                        await yieldToExtensionHost();
                    }
                    continue;
                }
                const mod: VbaModuleSymbols = {
                    moduleName: entry.name,
                    source: entry.source,
                    symbols: parseVbaModule(entry.source),
                    type: entry.type,
                    documentType: entry.documentType,
                };
                wb.modules.set(moduleKey, mod);
                out.push(mod);
                if ((index + 1) % WORKBOOK_INDEX_YIELD_EVERY_MODULES === 0) {
                    await yieldToExtensionHost();
                }
            }
            return out;
        } catch (err) {
            if (this.isReadModulesUnavailable(err)) {
                return undefined;
            }
            throw err;
        }
    }

    private isReadModulesUnavailable(err: unknown): boolean {
        const message = err instanceof Error ? err.message : String(err);
        return /Method not found:\s*readModules/i.test(message) ||
            /Unexpected bridge call readModules/i.test(message);
    }

    private workbook(workbookKey: string): CachedWorkbook {
        let wb = this._cache.get(workbookKey);
        if (!wb) {
            wb = { modules: new Map() };
            this._cache.set(workbookKey, wb);
        }
        return wb;
    }

    private moduleRequestKey(workbookKey: string, moduleKey: string): string {
        return `${workbookKey}\n${moduleKey}`;
    }

    private moduleGeneration(requestKey: string): number {
        return this._moduleGenerations.get(requestKey) ?? 0;
    }

    private bumpModuleGeneration(requestKey: string): void {
        this._moduleGenerations.set(requestKey, this.moduleGeneration(requestKey) + 1);
    }

    private deleteWorkbookInflight(workbookKey: string): void {
        const modulePrefix = `${workbookKey}\n`;
        for (const key of this._moduleReads.keys()) {
            if (key.startsWith(modulePrefix)) {
                this._moduleReads.delete(key);
                this.bumpModuleGeneration(key);
            }
        }
        this._moduleListReads.delete(workbookKey);
        this._allModuleReads.delete(workbookKey);
    }
}
