import * as vscode from 'vscode';
import { ProjectEngine } from './projectEngine';
import type { EventHandlerDocumentType } from './analyzer/completion/eventHandlers';
import { moduleIdentityKey, projectIdentityKey } from './xlideFileSystem';
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
    /** The project's conditional compilation arguments; same on every entry. */
    projectConditionalConstants?: string;
    /** A VB6 designer's class (`VB.Form`, `VB.MDIForm`), absent for Office forms. */
    designerClass?: string;
    /** The module's own file when the container's modules are files (VB6). */
    filePath?: string;
}

interface CachedProject {
    /** moduleName -> module symbols */
    modules: Map<string, VbaModuleSymbols>;
    /** Cached project module list from the bridge. */
    moduleList?: VbaModuleEntry[];
    moduleListLoadedAt?: number;
    /**
     * The project's conditional compilation arguments, as the VBE stores them.
     * A project-level fact that arrives with the module read, since the engine
     * has the dir stream open at that point.
     */
    conditionalConstants?: string;
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
    /** The project's conditional compilation arguments; same on every entry. */
    projectConditionalConstants?: string;
    /** A VB6 designer's class (`VB.Form`, `VB.MDIForm`), absent for Office forms. */
    designerClass?: string;
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
 * query; callers can invalidate single modules or whole projects after edits.
 */
export class VbaSymbolIndex implements vscode.Disposable {
    private _cache = new Map<string, CachedProject>();
    private _moduleReads = new Map<string, Promise<VbaModuleSymbols>>();
    private _moduleListReads = new Map<string, Promise<VbaModuleEntry[]>>();
    private _allModuleReads = new Map<string, Promise<VbaModuleSymbols[]>>();
    private _moduleGenerations = new Map<string, number>();
    private _emitter = new vscode.EventEmitter<{ projectPath: string; moduleName?: string }>();
    readonly onDidChange = this._emitter.event;

    constructor(private readonly _bridge: ProjectEngine) {}

    /** Invalidate one module (or the whole project when moduleName is omitted). */
    invalidate(projectPath: string, moduleName?: string): void {
        const key = projectIdentityKey(projectPath);
        const wb = this._cache.get(key);
        if (moduleName === undefined) {
            this._cache.delete(key);
            this.deleteProjectInflight(key);
        } else {
            const moduleKey = moduleIdentityKey(moduleName);
            wb?.modules.delete(moduleKey);
            const requestKey = this.moduleRequestKey(key, moduleKey);
            this.bumpModuleGeneration(requestKey);
            this._moduleReads.delete(requestKey);
            this._allModuleReads.delete(key);
        }
        this._emitter.fire({ projectPath, moduleName });
    }

    invalidateAll(): void {
        this._cache.clear();
        this._moduleReads.clear();
        this._moduleListReads.clear();
        this._allModuleReads.clear();
        this._moduleGenerations.clear();
        this._emitter.fire({ projectPath: '' });
    }

    /** Returns the cached source for a single module, loading on demand. */
    async getModule(projectPath: string, moduleName: string): Promise<VbaModuleSymbols> {
        const key = projectIdentityKey(projectPath);
        const wb = this.cachedProject(key);
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
                { path: projectPath, module: moduleName },
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

    /** Returns the cached source for every module in the project. */
    /**
     * The project's conditional compilation arguments, if a module read has
     * already fetched them. Empty until then, which leaves a custom `#If`
     * undecidable rather than guessing at it.
     */
    projectConditionalConstants(projectPath: string): string | undefined {
        return this.cachedProject(projectIdentityKey(projectPath)).conditionalConstants;
    }

    async getAllModules(projectPath: string): Promise<VbaModuleSymbols[]> {
        const key = projectIdentityKey(projectPath);
        const existingRead = this._allModuleReads.get(key);
        if (existingRead) { return existingRead; }

        const promise = (async () => {
            const trace = startPerformanceTrace('projectContext.getAllModules');
            try {
                const cached = this.cachedAllModules(key);
                if (cached) {
                    trace.end('ok', 'cached');
                    return cached;
                }
                const batch = await this.getAllModulesFromBatchRead(projectPath, key);
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
    peekModule(projectPath: string, moduleName: string): VbaModuleSymbols | undefined {
        return this._cache
            .get(projectIdentityKey(projectPath))
            ?.modules.get(moduleIdentityKey(moduleName));
    }

    /**
     * Refreshes a single module's source from disk.
     * Useful immediately after a write so the cache reflects the new content.
     */
    async refreshModule(projectPath: string, moduleName: string): Promise<VbaModuleSymbols> {
        this.invalidate(projectPath, moduleName);
        return this.getModule(projectPath, moduleName);
    }

    /**
     * Updates one cached module directly from an already-known source snapshot.
     * This avoids a bridge round-trip after saving a virtual VBA editor buffer.
     */
    updateModuleSource(
        projectPath: string,
        moduleName: string,
        source: string,
        metadata: { type?: string; documentType?: EventHandlerDocumentType } = {},
    ): VbaModuleSymbols {
        const key = projectIdentityKey(projectPath);
        const moduleKey = moduleIdentityKey(moduleName);
        const requestKey = this.moduleRequestKey(key, moduleKey);
        this.bumpModuleGeneration(requestKey);
        this._moduleReads.delete(requestKey);
        this._allModuleReads.delete(key);

        const wb = this.cachedProject(key);
        const existing = wb.modules.get(moduleKey);
        const mod: VbaModuleSymbols = {
            moduleName,
            source,
            type: metadata.type ?? existing?.type,
            documentType: metadata.documentType ?? existing?.documentType,
        };
        wb.modules.set(moduleKey, mod);
        this._emitter.fire({ projectPath, moduleName });
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

    private cachedAllModules(projectKey: string): VbaModuleSymbols[] | undefined {
        const wb = this._cache.get(projectKey);
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
            mod.designerClass = entry.designerClass;
            modules.push(mod);
        }
        return modules;
    }

    private async getAllModulesFromBatchRead(
        projectPath: string,
        projectKey: string,
    ): Promise<VbaModuleSymbols[]> {
        const entries = await this._bridge.call<VbaModuleSourceEntry[]>(
            'readModules',
            { path: projectPath },
        );
        const wb = this.cachedProject(projectKey);
        wb.moduleList = entries.map(({ name, type, documentType, predeclaredId, designerClass }) =>
            ({ name, type, documentType, predeclaredId, designerClass }));
        wb.moduleListLoadedAt = Date.now();
        wb.conditionalConstants = entries.find((e) => e.projectConditionalConstants)
            ?.projectConditionalConstants;
        const out: VbaModuleSymbols[] = [];
        for (const [index, entry] of entries.entries()) {
            const moduleKey = moduleIdentityKey(entry.name);
            const requestKey = this.moduleRequestKey(projectKey, moduleKey);
            const existing = wb.modules.get(moduleKey);
            if (existing && this.moduleGeneration(requestKey) > 0) {
                existing.type = entry.type;
                existing.documentType = entry.documentType;
                existing.implicitMembers = entry.implicitMembers;
                existing.predeclaredId = entry.predeclaredId;
                existing.designerClass = entry.designerClass;
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
                designerClass: entry.designerClass,
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

    private cachedProject(projectKey: string): CachedProject {
        let wb = this._cache.get(projectKey);
        if (!wb) {
            wb = { modules: new Map() };
            this._cache.set(projectKey, wb);
        }
        return wb;
    }

    private moduleRequestKey(projectKey: string, moduleKey: string): string {
        return `${projectKey}\n${moduleKey}`;
    }

    private moduleGeneration(requestKey: string): number {
        return this._moduleGenerations.get(requestKey) ?? 0;
    }

    private bumpModuleGeneration(requestKey: string): void {
        this._moduleGenerations.set(requestKey, this.moduleGeneration(requestKey) + 1);
    }

    private deleteProjectInflight(projectKey: string): void {
        const modulePrefix = `${projectKey}\n`;
        for (const key of this._moduleReads.keys()) {
            if (key.startsWith(modulePrefix)) {
                this._moduleReads.delete(key);
                this.bumpModuleGeneration(key);
            }
        }
        this._moduleListReads.delete(projectKey);
        this._allModuleReads.delete(projectKey);
    }
}
