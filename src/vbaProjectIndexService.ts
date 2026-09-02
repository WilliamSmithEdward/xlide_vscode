// Shared incremental per-project index (audit #22/#122/#49).
//
// Owns ONE long-lived ProjectIndex per project, built once from the symbol
// index and then kept current by folding in single-module deltas: open-editor
// text on access (per document version) and saved/index changes via
// VbaSymbolIndex.onDidChange. This replaces the per-consumer caches that each
// rebuilt the whole project index from scratch (navigation per open-document
// version, completion per document version, semantic tokens per 5s TTL).
//
// View semantics, reconciling the previously divergent consumers:
//  - 'live' (diagnostics, navigation reads, completion, semantic tokens):
//    a module whose latest source fails to index keeps its last successfully
//    indexed version in both the project and the byModule views, exactly like
//    the previous diagnostics context. Modules that have never indexed
//    successfully are absent, matching the previous fresh "live" builds.
//  - 'strict' (rename): the recorded per-module failure is rethrown, matching
//    buildVbaProjectIndexAsync without ignoreInvalidModules, which threw on
//    the first invalid module of a fresh build.
// The byModule views always carry the source text the project index actually
// parsed, so symbol spans resolved against them are self-consistent.

import * as vscode from 'vscode';
import * as path from 'path';
import {
    XLIDE_SCHEME,
    moduleIdentityKey,
    projectIdentityKey,
} from './xlideFileSystem';
import { VbaSymbolIndex, type VbaModuleSymbols } from './vbaSymbolIndex';
import { analysisSourceForDocument, moduleLocationOfDocument } from './vbaDocumentLocation';
import type {
    EventHandlerDocumentType,
    ModuleSymbolKind,
    ProjectIndex,
} from './analyzer';
import {
    buildLiveVbaProjectIndexAsync,
    moduleKindFromType,
    projectProcedureSignatures,
} from './vbaProjectAnalysis';

// Invalidation is event-driven (the symbol index fires onDidChange for module
// edits, adds, and removals), so the TTL is only a stale-context backstop.
const PROJECT_INDEX_CONTEXT_TTL_MS = 10 * 60_000;

export type VbaProjectViewMode = 'live' | 'strict';

export interface VbaProjectModuleMetadata {
    moduleName: string;
    moduleType?: string;
    moduleKind: ModuleSymbolKind;
    documentType?: EventHandlerDocumentType;
    /** A VB6 designer's class (`VB.Form`, `VB.MDIForm`), absent for Office forms. */
    designerClass?: string;
}

export interface VbaProjectContext {
    readonly projectPath: string;
    readonly project: ProjectIndex;
    /** Module views whose source is the text the project index last parsed. */
    readonly modules: VbaModuleSymbols[];
    readonly byModule: Map<string, VbaModuleSymbols>;
    /** Metadata for every known module, including currently invalid ones. */
    readonly moduleMetadata: Map<string, VbaProjectModuleMetadata>;
    /** Consumer-memoized signatures; the service resets this on any change. */
    projectProcedures?: ReturnType<typeof projectProcedureSignatures>;
    readonly loadedAt: number;
    /**
     * Monotonic counter of module-source changes applied to this record,
     * excluding changes to `moduleName` itself. Consumers analyzing one module
     * use it to detect when any OTHER module's content moved (cross-module
     * context changed) - e.g. to invalidate incremental analysis state or
     * reseed an analysis worker.
     */
    crossModuleGeneration(moduleName: string): number;
}

class ProjectRecord implements VbaProjectContext {
    readonly byModule = new Map<string, VbaModuleSymbols>();
    readonly moduleMetadata = new Map<string, VbaProjectModuleMetadata>();
    /** moduleKey -> error from the most recent failed module apply. */
    readonly invalidModules = new Map<string, unknown>();
    readonly appliedDocumentVersions = new Map<string, number>();
    projectProcedures?: ReturnType<typeof projectProcedureSignatures>;
    loadedAt = Date.now();
    private _changeCounter = 0;
    private readonly _changesByModule = new Map<string, number>();

    constructor(
        readonly projectPath: string,
        readonly project: ProjectIndex,
    ) {}

    get modules(): VbaModuleSymbols[] {
        return [...this.byModule.values()];
    }

    /** Folds one module source into the project, tracking views and validity. */
    applyModule(
        moduleName: string,
        source: string,
        metadata: { moduleType?: string; documentType?: EventHandlerDocumentType; designerClass?: string },
    ): void {
        const moduleKey = moduleIdentityKey(moduleName);
        const previous = this.moduleMetadata.get(moduleKey);
        const moduleType = metadata.moduleType ?? previous?.moduleType;
        const meta: VbaProjectModuleMetadata = {
            moduleName,
            moduleType,
            moduleKind: moduleKindFromType(moduleType),
            documentType: metadata.documentType ?? previous?.documentType,
            designerClass: metadata.designerClass ?? previous?.designerClass,
        };
        this.moduleMetadata.set(moduleKey, meta);
        this._recordChange(moduleName);
        try {
            this.project.setModule({
                moduleName,
                moduleKind: meta.moduleKind,
                source,
            });
        } catch (err) {
            // Keep the previous indexed version while the latest source is
            // parser-recovered/incomplete; remember the failure so strict
            // consumers (rename) surface it the way a fresh strict build did.
            this.invalidModules.set(moduleKey, err);
            this.markChanged();
            return;
        }
        this.invalidModules.delete(moduleKey);
        this.byModule.set(moduleKey, {
            moduleName,
            source,
            type: meta.moduleType,
            documentType: meta.documentType,
        });
        this.markChanged();
    }

    markChanged(): void {
        this.projectProcedures = undefined;
        this.loadedAt = Date.now();
    }

    crossModuleGeneration(moduleName: string): number {
        return this._changeCounter - (this._changesByModule.get(moduleIdentityKey(moduleName)) ?? 0);
    }

    private _recordChange(moduleName: string): void {
        this._changeCounter += 1;
        const key = moduleIdentityKey(moduleName);
        this._changesByModule.set(key, (this._changesByModule.get(key) ?? 0) + 1);
    }

    /** Throws the first recorded invalid-module error, in module-list order. */
    throwIfInvalid(): void {
        if (this.invalidModules.size === 0) {
            return;
        }
        for (const moduleKey of this.moduleMetadata.keys()) {
            if (this.invalidModules.has(moduleKey)) {
                throw this.invalidModules.get(moduleKey);
            }
        }
        throw [...this.invalidModules.values()][0];
    }
}

function projectKey(projectPath: string): string {
    return projectIdentityKey(path.resolve(projectPath));
}

/**
 * One long-lived incremental ProjectIndex per project, shared by diagnostics,
 * navigation, completion/hover/signature help, and semantic tokens.
 */
export class VbaProjectIndexService implements vscode.Disposable {
    private readonly _records = new Map<string, ProjectRecord>();
    private readonly _loads = new Map<string, Promise<ProjectRecord>>();
    private readonly _subscriptions: vscode.Disposable[];

    constructor(private readonly _index: VbaSymbolIndex) {
        this._subscriptions = [
            _index.onDidChange(({ projectPath, moduleName }) => {
                if (projectPath && moduleName && this._applyIndexModule(projectPath, moduleName)) {
                    return;
                }
                this.invalidate(projectPath || undefined);
            }),
            vscode.workspace.onDidCloseTextDocument((document) => {
                // A closed editor reverts to the indexed module content, so the
                // project context must be rebuilt from the symbol index.
                this._invalidateForDocument(document);
            }),
        ];
    }

    /**
     * Returns the shared project context, folding in any open-document edits
     * first. 'strict' rethrows the latest indexing failure when any module's
     * current source is invalid; 'live' serves last-good content instead.
     */
    async contextForProject(
        projectPath: string,
        mode: VbaProjectViewMode = 'live',
    ): Promise<VbaProjectContext> {
        const record = await this._recordForProject(projectPath);
        this._applyOpenDocumentSources(projectPath, record);
        if (mode === 'strict') {
            record.throwIfInvalid();
        }
        return record;
    }

    invalidate(projectPath?: string): void {
        if (!projectPath) {
            this._records.clear();
            this._loads.clear();
            return;
        }
        const key = projectKey(projectPath);
        this._records.delete(key);
        this._loads.delete(key);
    }

    dispose(): void {
        for (const subscription of this._subscriptions) {
            subscription.dispose();
        }
        this._records.clear();
        this._loads.clear();
    }

    private async _recordForProject(projectPath: string): Promise<ProjectRecord> {
        const key = projectKey(projectPath);
        const cached = this._records.get(key);
        if (cached && Date.now() - cached.loadedAt < PROJECT_INDEX_CONTEXT_TTL_MS) {
            return cached;
        }
        const existingLoad = this._loads.get(key);
        if (existingLoad) {
            return existingLoad;
        }
        const load = this._buildRecord(projectPath);
        this._loads.set(key, load);
        try {
            const record = await load;
            if (this._loads.get(key) === load) {
                this._records.set(key, record);
            }
            return record;
        } finally {
            if (this._loads.get(key) === load) {
                this._loads.delete(key);
            }
        }
    }

    private async _buildRecord(projectPath: string): Promise<ProjectRecord> {
        const modules = await this._index.getAllModules(projectPath);
        const invalid = new Map<string, unknown>();
        const project = await buildLiveVbaProjectIndexAsync(
            modules.map((mod) => ({
                moduleName: mod.moduleName,
                moduleKind: moduleKindFromType(mod.type),
                type: mod.type,
                documentType: mod.documentType,
                source: mod.source,
                implicitMembers: mod.implicitMembers,
                predeclaredId: mod.predeclaredId,
            })),
            undefined,
            {
                onInvalidModule: (moduleName, error) =>
                    invalid.set(moduleIdentityKey(moduleName), error),
            },
        );
        const record = new ProjectRecord(projectPath, project);
        for (const mod of modules) {
            const moduleKey = moduleIdentityKey(mod.moduleName);
            record.moduleMetadata.set(moduleKey, {
                moduleName: mod.moduleName,
                moduleType: mod.type,
                moduleKind: moduleKindFromType(mod.type),
                documentType: mod.documentType,
                designerClass: mod.designerClass,
            });
            if (invalid.has(moduleKey)) {
                record.invalidModules.set(moduleKey, invalid.get(moduleKey));
                continue;
            }
            record.byModule.set(moduleKey, {
                moduleName: mod.moduleName,
                source: mod.source,
                type: mod.type,
                documentType: mod.documentType,
                implicitMembers: mod.implicitMembers,
                predeclaredId: mod.predeclaredId,
                filePath: mod.filePath,
            });
        }
        return record;
    }

    /**
     * Folds the editor text of every open module of the container (by
     * version): a project's virtual documents, or a VB6 project's own files.
     */
    private _applyOpenDocumentSources(projectPath: string, record: ProjectRecord): void {
        const key = projectKey(projectPath);
        for (const openDocument of vscode.workspace.textDocuments) {
            const location = moduleLocationOfDocument(openDocument);
            if (!location || projectKey(location.projectPath) !== key) {
                continue;
            }
            const documentKey = openDocument.uri.toString();
            if (record.appliedDocumentVersions.get(documentKey) === openDocument.version) {
                continue;
            }
            record.applyModule(location.moduleName, analysisSourceForDocument(openDocument), {});
            record.appliedDocumentVersions.set(documentKey, openDocument.version);
        }
    }

    // Incrementally folds a single changed module into the cached record so a
    // save does not force a full-project rebuild. Returns false when the
    // update could not be applied and the caller must invalidate instead.
    private _applyIndexModule(projectPath: string, moduleName: string): boolean {
        const key = projectKey(projectPath);
        if (this._loads.has(key)) {
            return false;
        }
        const record = this._records.get(key);
        if (!record) {
            // Nothing cached, so there is nothing to refresh or invalidate.
            return true;
        }
        const mod = this._index.peekModule(projectPath, moduleName);
        if (!mod) {
            return false;
        }
        record.applyModule(mod.moduleName, mod.source, {
            moduleType: mod.type,
            documentType: mod.documentType,
            designerClass: mod.designerClass,
        });
        return true;
    }

    private _invalidateForDocument(document: vscode.TextDocument): void {
        const location = moduleLocationOfDocument(document);
        if (location) {
            this.invalidate(location.projectPath);
        } else if (document.uri.scheme === XLIDE_SCHEME) {
            // A virtual document whose URI no longer decodes: nothing says
            // which project it was, so every record is suspect.
            this._records.clear();
        }
    }
}
