import * as vscode from 'vscode';
import { WorkbookEngine } from './workbookEngine';
import type { EventHandlerDocumentType } from './analyzer/completion/eventHandlers';
import { moduleIdentityKey, workbookIdentityKey } from './xlideFileSystem';
import { startPerformanceTrace } from './performanceTrace';
import { yieldToExtensionHost } from './util/async';

export interface VbaModuleSymbols {
    moduleName: string;
    /** Cached module source. */
    source: string;
    /** Host module type from listModules (standard/class/document/userform). */
    type?: string;
    /** Excel document subtype from listModules when the bridge can prove it. */
    documentType?: EventHandlerDocumentType;
    /** A form's designer-declared controls, read from the designer storage. */
    implicitMembers?: { name: string; type: string }[];
    /**
     * True when the module carries `Attribute VB_PredeclaredId = True`, giving
     * it a default instance so its own name is usable as a value. Absent means
     * the attribute header was not read, never "no".
     */
    predeclaredId?: boolean;
    /** The module's own file when the container's modules are files (VB6). */
    filePath?: string;
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
    /** A form's designer-declared controls, from the engine's designer read. */
    implicitMembers?: { name: string; type: string }[];
    /**
     * True when the module carries `Attribute VB_PredeclaredId = True`, giving
     * it a default instance so its own name is usable as a value. Absent means
     * the attribute header was not read, never "no".
     */
    predeclaredId?: boolean;
    /** The module's own file when the container's modules are files (VB6). */
    filePath?: string;
}

interface VbaModuleSourceEntry extends VbaModuleEntry {
    source: string;
}

const MODULE_LIST_CACHE_TTL_MS = 5_000;
const WORKBOOK_INDEX_YIELD_EVERY_MODULES = 8;

/**
 * Workbook-scoped VBA module source cache. Lazily loads modules on first
 * query; callers can invalidate single modules or whole workbooks after edits.
 */
export class VbaSymbolIndex implements vscode.Disposable {
    private _cache = new Map<string, CachedWorkbook>();
    private _moduleReads = new Map<string, Promise<VbaModuleSymbols>>();
    private _moduleListReads = new Map<string, Promise<VbaModuleEntry[]>>();
    private _allModuleReads = new Map<string, Promise<VbaModuleSymbols[]>>();
    private _moduleGenerations = new Map<string, number>();
    private _emitter = new vscode.EventEmitter<{ xlsmPath: string; moduleName?: string }>();
    readonly onDidChange = this._emitter.event;

    constructor(private readonly _bridge: WorkbookEngine) {}

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

    /** Returns the cached source for a single module, loading on demand. */
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

    /** Returns the cached source for every module in the workbook. */
    async getAllModules(xlsmPath: string): Promise<VbaModuleSymbols[]> {
        const key = workbookIdentityKey(xlsmPath);
        const existingRead = this._allModuleReads.get(key);
        if (existingRead) { return existingRead; }

        const promise = (async () => {
            const trace = startPerformanceTrace('workbookContext.getAllModules');
            try {
                const cached = this.cachedAllModules(key);
                if (cached) {
                    trace.end('ok', 'cached');
                    return cached;
                }
                const batch = await this.getAllModulesFromBatchRead(xlsmPath, key);
                trace.end('ok', 'batch');
                return batch;
            } catch (err) {
                trace.end('failed');
                throw err;
            }
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

    /** Returns the cached module entry without triggering a load. */
    peekModule(xlsmPath: string, moduleName: string): VbaModuleSymbols | undefined {
        return this._cache
            .get(workbookIdentityKey(xlsmPath))
            ?.modules.get(moduleIdentityKey(moduleName));
    }

    /**
     * Refreshes a single module's source from disk.
     * Useful immediately after a write so the cache reflects the new content.
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
            mod.predeclaredId = entry.predeclaredId;
            modules.push(mod);
        }
        return modules;
    }

    private async getAllModulesFromBatchRead(
        xlsmPath: string,
        workbookKey: string,
    ): Promise<VbaModuleSymbols[]> {
        const entries = await this._bridge.call<VbaModuleSourceEntry[]>(
            'readModules',
            { path: xlsmPath },
        );
        const wb = this.workbook(workbookKey);
        wb.moduleList = entries.map(({ name, type, documentType, predeclaredId }) =>
            ({ name, type, documentType, predeclaredId }));
        wb.moduleListLoadedAt = Date.now();
        const out: VbaModuleSymbols[] = [];
        for (const [index, entry] of entries.entries()) {
            const moduleKey = moduleIdentityKey(entry.name);
            const requestKey = this.moduleRequestKey(workbookKey, moduleKey);
            const existing = wb.modules.get(moduleKey);
            if (existing && this.moduleGeneration(requestKey) > 0) {
                existing.type = entry.type;
                existing.documentType = entry.documentType;
                existing.implicitMembers = entry.implicitMembers;
                existing.predeclaredId = entry.predeclaredId;
                existing.filePath = entry.filePath;
                out.push(existing);
                if ((index + 1) % WORKBOOK_INDEX_YIELD_EVERY_MODULES === 0) {
                    await yieldToExtensionHost();
                }
                continue;
            }
            const mod: VbaModuleSymbols = {
                moduleName: entry.name,
                source: entry.source,
                type: entry.type,
                documentType: entry.documentType,
                implicitMembers: entry.implicitMembers,
                predeclaredId: entry.predeclaredId,
                filePath: entry.filePath,
            };
            wb.modules.set(moduleKey, mod);
            out.push(mod);
            if ((index + 1) % WORKBOOK_INDEX_YIELD_EVERY_MODULES === 0) {
                await yieldToExtensionHost();
            }
        }
        return out;
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
