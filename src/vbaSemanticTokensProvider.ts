// Semantic tokens for VBA type references (class/enum/struct/type names),
// with a TTL'd per-document project-types cache and debounced background
// refresh against the shared workbook ProjectIndex.
//
// Extracted verbatim from vbaLanguageProviders.ts (audit #21).

import * as vscode from 'vscode';
import { XLIDE_SCHEME, decodeModuleUri, isVbaDocument } from './xlideFileSystem';
import {
    analysisSourceForDocument,
    liveProjectIndexForDocument,
    moduleKindFromDocument,
    moduleNameFromDocument,
} from './vbaDocumentIdentity';
import { moduleLocationOfDocument } from './vbaDocumentLocation';
import {
    collectHostGlobalTokens,
    collectHostMemberMethodTokens,
    collectImplicitMemberMethodTokens,
    resolveTypeSemanticTokens,
    TypeSemanticTokenType,
} from './analyzer';
import { codeNameHostTypesForModules } from './vbaEditorProjectContext';
import {
    projectAnalysisOptionsForModule,
    type VbaProjectAnalysisOptions,
} from './vbaProjectAnalysis';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import {
    hostObjectModelForToken,
    hostTokenForFileName,
} from './analyzer/host/hostRegistry';
import type { HostObjectModel } from './analyzer/host/excelObjectModel';
import { moduleIdentityKey } from './workbookIdentity';
import { startPerformanceTrace } from './performanceTrace';

const TYPE_TOKEN_TYPES: TypeSemanticTokenType[] = [
    'class',
    'enum',
    'enumMember',
    'struct',
    'type',
    'variable',
    'function',
    'property',
];
// `defaultLibrary` marks host-injected globals (Application, ThisWorkbook, ...);
// most themes give it a subtle tint and themes that don't fall back cleanly to
// the identifier color.
const TYPE_TOKEN_MODIFIERS = ['defaultLibrary'];
export const TYPE_TOKEN_LEGEND = new vscode.SemanticTokensLegend(TYPE_TOKEN_TYPES, TYPE_TOKEN_MODIFIERS);
const TYPE_SEMANTIC_PROJECT_TYPES_CACHE_TTL_MS = 5000;
const TYPE_SEMANTIC_PROJECT_TYPES_REFRESH_DELAY_MS = 350;
const TYPE_SEMANTIC_CACHE_MAX_DOCUMENTS = 64;

interface CachedTypeSemanticProjectTypes {
    at: number;
    projectTypes: VbaProjectAnalysisOptions['projectTypes'];
    /** A form's designer-declared controls, when anything knows them. */
    implicitMembers?: VbaProjectAnalysisOptions['implicitMembers'];
    /** `MSForms.UserForm` when the module is a form, so `Me.Hide` paints. */
    meType?: string;
    /** The container's host model; absent keeps the Excel default. */
    hostModel?: HostObjectModel;
    /** Lowercased document code name -> host type, for member paint (issue #29). */
    codeNames?: Record<string, string>;
    /** Host type `Me` denotes in a document module, so `Me.Calculate` paints (issue #31). */
    meHostType?: string;
}

interface CachedTypeSemanticTokens {
    documentVersion: number;
    projectTypesLoadedAt: number;
    tokens: vscode.SemanticTokens;
}

export class VbaTypeSemanticTokensProvider implements vscode.DocumentSemanticTokensProvider, vscode.Disposable {
    private readonly _onDidChangeSemanticTokens = new vscode.EventEmitter<void>();
    private readonly _projectTypesCache = new Map<string, CachedTypeSemanticProjectTypes>();
    private readonly _semanticTokensCache = new Map<string, CachedTypeSemanticTokens>();
    private readonly _projectTypeRefreshes = new Set<string>();
    private readonly _projectTypeRefreshTimers = new Map<string, ReturnType<typeof setTimeout>>();

    readonly onDidChangeSemanticTokens = this._onDidChangeSemanticTokens.event;

    constructor(private readonly _projectIndexService: VbaProjectIndexService) {}

    dispose(): void {
        this._onDidChangeSemanticTokens.dispose();
        for (const timer of this._projectTypeRefreshTimers.values()) {
            clearTimeout(timer);
        }
        this._projectTypeRefreshTimers.clear();
        this._projectTypeRefreshes.clear();
        this._projectTypesCache.clear();
        this._semanticTokensCache.clear();
    }

    async provideDocumentSemanticTokens(
        document: vscode.TextDocument,
        token: vscode.CancellationToken,
    ): Promise<vscode.SemanticTokens> {
        const trace = startPerformanceTrace('semanticTokens', document.uri.scheme);
        const builder = new vscode.SemanticTokensBuilder(TYPE_TOKEN_LEGEND);
        try {
            if (!isVbaDocument(document)) { return builder.build(); }

            const source = analysisSourceForDocument(document);
            const moduleName = moduleNameFromDocument(document);
            const projectContext = document.uri.scheme === XLIDE_SCHEME
                ? this._cachedProjectTypesForDocument(document, { requireFresh: false })
                : await this._projectTypesForDocument(document, source, moduleName, token);
            const projectTypes = projectContext?.projectTypes ?? [];
            if (
                document.uri.scheme === XLIDE_SCHEME &&
                !this._cachedProjectTypesForDocument(document, { requireFresh: true })
            ) {
                this._scheduleProjectTypesRefresh(document);
            }
            const projectTypesLoadedAt = this._projectTypesCache.get(document.uri.toString())?.at ?? 0;
            const cachedTokens = this._semanticTokensCache.get(document.uri.toString());
            if (
                cachedTokens &&
                cachedTokens.documentVersion === document.version &&
                cachedTokens.projectTypesLoadedAt === projectTypesLoadedAt
            ) {
                return cachedTokens.tokens;
            }

            const items = [
                ...resolveTypeSemanticTokens(source, { projectTypes }),
                ...collectHostGlobalTokens(
                    source,
                    projectContext?.hostModel,
                    projectContext?.implicitMembers,
                ),
                ...collectImplicitMemberMethodTokens(source, {
                    implicitMembers: projectContext?.implicitMembers,
                    meType: projectContext?.meType,
                }),
                ...collectHostMemberMethodTokens(source, {
                    model: projectContext?.hostModel,
                    codeNames: projectContext?.codeNames,
                    implicitMembers: projectContext?.implicitMembers,
                    meType: projectContext?.meHostType,
                    projectTypes,
                }),
            ];
            for (const item of items) {
                if (token.isCancellationRequested) { break; }
                builder.push(
                    new vscode.Range(
                        document.positionAt(item.span.start),
                        document.positionAt(item.span.end),
                    ),
                    item.tokenType,
                    item.modifiers ?? [],
                );
            }
            const tokens = builder.build();
            if (!token.isCancellationRequested) {
                this._semanticTokensCache.set(document.uri.toString(), {
                    documentVersion: document.version,
                    projectTypesLoadedAt,
                    tokens,
                });
                this._pruneSemanticTokenCaches();
            }
            return tokens;
        } finally {
            trace.end(token.isCancellationRequested ? 'canceled' : 'ok', document.uri.scheme);
        }
    }

    private _cachedProjectTypesForDocument(
        document: vscode.TextDocument,
        options: { requireFresh?: boolean } = {},
    ): CachedTypeSemanticProjectTypes | undefined {
        const cached = this._projectTypesCache.get(document.uri.toString());
        if (!cached) {
            return undefined;
        }
        if (
            options.requireFresh &&
            Date.now() - cached.at >= TYPE_SEMANTIC_PROJECT_TYPES_CACHE_TTL_MS
        ) {
            return undefined;
        }
        return cached;
    }

    private _scheduleProjectTypesRefresh(document: vscode.TextDocument): void {
        const key = document.uri.toString();
        if (this._projectTypeRefreshes.has(key)) {
            return;
        }
        const existingTimer = this._projectTypeRefreshTimers.get(key);
        if (existingTimer) {
            clearTimeout(existingTimer);
        }
        const timer = setTimeout(() => {
            this._projectTypeRefreshTimers.delete(key);
            if (!vscode.workspace.textDocuments.includes(document)) {
                return;
            }
            this._refreshProjectTypesInBackground(
                document,
                analysisSourceForDocument(document),
                moduleNameFromDocument(document),
            );
        }, TYPE_SEMANTIC_PROJECT_TYPES_REFRESH_DELAY_MS);
        this._projectTypeRefreshTimers.set(key, timer);
    }

    private _refreshProjectTypesInBackground(
        document: vscode.TextDocument,
        source: string,
        moduleName: string,
    ): void {
        const key = document.uri.toString();
        if (this._projectTypeRefreshes.has(key)) {
            return;
        }
        this._projectTypeRefreshes.add(key);
        void this._projectTypesForDocument(document, source, moduleName)
            .then(() => {
                this._semanticTokensCache.delete(key);
                this._onDidChangeSemanticTokens.fire();
            })
            .finally(() => this._projectTypeRefreshes.delete(key));
    }

    private _pruneSemanticTokenCaches(): void {
        const openKeys = new Set(vscode.workspace.textDocuments.map((document) => document.uri.toString()));
        for (const key of this._semanticTokensCache.keys()) {
            if (!openKeys.has(key)) {
                this._semanticTokensCache.delete(key);
                this._projectTypesCache.delete(key);
            }
        }
        const overflow = this._semanticTokensCache.size - TYPE_SEMANTIC_CACHE_MAX_DOCUMENTS;
        if (overflow <= 0) {
            return;
        }
        for (const key of [...this._semanticTokensCache.keys()].slice(0, overflow)) {
            this._semanticTokensCache.delete(key);
            this._projectTypesCache.delete(key);
        }
    }

    private async _projectTypesForDocument(
        document: vscode.TextDocument,
        source: string,
        moduleName: string,
        token?: vscode.CancellationToken,
    ): Promise<CachedTypeSemanticProjectTypes | undefined> {
        const key = document.uri.toString();
        const cached = this._cachedProjectTypesForDocument(document, { requireFresh: true });
        if (cached) {
            return cached;
        }

        const previous = this._projectTypesCache.get(key);

        try {
            const project = await liveProjectIndexForDocument(
                this._projectIndexService,
                document,
                source,
                moduleName,
                token,
            );
            const options = projectAnalysisOptionsForModule(project, moduleName);
            const codeNames = await this._codeNamesForDocument(document);
            const entry: CachedTypeSemanticProjectTypes = {
                at: Date.now(),
                projectTypes: options.projectTypes ?? [],
                implicitMembers: options.implicitMembers,
                meType: await this._userFormMeType(document, moduleName),
                hostModel: hostModelForDocument(document),
                codeNames,
                // A document module's own code name IS what `Me` denotes there;
                // any other module kind has no entry and `Me.` stays unpainted.
                meHostType: codeNames?.[moduleName.toLowerCase()],
            };
            this._projectTypesCache.set(key, entry);
            return entry;
        } catch {
            return previous;
        }
    }

    /**
     * Lowercased code-name -> host type map for the document's container, so
     * `Sheet1.Calculate` and Word's `ThisDocument.Save` paint as method calls
     * (issue #29). Loose files have no sibling document modules; undefined.
     */
    private async _codeNamesForDocument(
        document: vscode.TextDocument,
    ): Promise<Record<string, string> | undefined> {
        const location = moduleLocationOfDocument(document);
        if (!location) {
            return undefined;
        }
        try {
            const xlsmPath = location.xlsmPath;
            // The same cached workbook context the project build above used.
            const context = await this._projectIndexService.contextForWorkbook(xlsmPath, 'live');
            return codeNameHostTypesForModules(
                [...context.moduleMetadata.values()].map((meta) => ({
                    name: meta.moduleName,
                    type: meta.moduleType ?? '',
                    documentType: meta.documentType,
                })),
                hostTokenForFileName(xlsmPath),
            );
        } catch {
            return undefined;
        }
    }

    /** `MSForms.UserForm` when the document is a form's code-behind. */
    // (see hostModelForDocument below for the host side)
    private async _userFormMeType(
        document: vscode.TextDocument,
        moduleName: string,
    ): Promise<string | undefined> {
        const location = moduleLocationOfDocument(document);
        if (!location) {
            return moduleKindFromDocument(document) === 'userform' ? 'MSForms.UserForm' : undefined;
        }
        if (hostTokenForFileName(location.xlsmPath) === 'vb6') {
            // A VB6 form is a VB.Form; its surface arrives with the vb6 model.
            return undefined;
        }
        try {
            // The same cached workbook context the project build above used.
            const context = await this._projectIndexService.contextForWorkbook(
                location.xlsmPath,
                'live',
            );
            return context.moduleMetadata.get(moduleIdentityKey(moduleName))?.moduleKind === 'userform'
                ? 'MSForms.UserForm'
                : undefined;
        } catch {
            return undefined;
        }
    }
}

/** The host model for the document's container; undefined keeps Excel defaults. */
function hostModelForDocument(document: vscode.TextDocument): HostObjectModel | undefined {
    const location = moduleLocationOfDocument(document);
    return location ? hostObjectModelForToken(hostTokenForFileName(location.xlsmPath)) : undefined;
}
