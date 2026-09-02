// Editor project-context service for the VBA completion stack.
//
// Builds and caches the per-document EditorProjectContext (module identity,
// code names, Me types, project types/members/procedures/symbols) on top of
// the shared VbaProjectIndexService, with in-flight build dedup, a TTL'd
// per-document cache, and budgeted builds for latency-sensitive callers.
// Extracted verbatim from vbaMemberCompletion.ts (audit #27).

import * as vscode from 'vscode';
import {
	projectIdentityKey,
} from './xlideFileSystem';
import {
	EventHandlerCompletionContext,
	EventHandlerDocumentType,
	eventHandlerDocumentTypeForContext,
	IdentifierCompletionContext,
	MemberCompletionContext,
	ModuleSymbolKind,
	VbaProcedureSignature,
	TypeCompletionContext,
} from './analyzer';
import {
	buildLiveVbaProjectIndex,
	buildLiveVbaProjectIndexAsync,
	moduleKindFromType,
	projectEditorSymbolContextForModule,
} from './vbaProjectAnalysis';
import {
	hostObjectModelForToken,
	hostTokenForFileName,
	type VbaHostToken,
} from './analyzer/host/hostRegistry';
import type { HostObjectModel } from './analyzer/host/excelObjectModel';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import { moduleLocationOfDocument, moduleLocationOfUri } from './vbaDocumentLocation';
import { blankDesignerHeader } from './vba/moduleSource';

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const CHART = 'Excel.Chart';
const USERFORM = 'MSForms.UserForm';
const WORD_DOCUMENT = 'Word.Document';
const VB6_FORM = 'VB.Form';
const VB6_MDI_FORM = 'VB.MDIForm';
const EDITOR_PROJECT_CONTEXT_CACHE_TTL_MS = 10_000;
const EDITOR_PROJECT_CONTEXT_CACHE_MAX_DOCUMENTS = 32;

export interface ModuleEntry {
	name: string;
	type: string;
	documentType?: EventHandlerDocumentType;
	/** A VB6 designer's class (`VB.Form`, `VB.MDIForm`), which decides what `Me` is. */
	designerClass?: string;
}

export interface EditorProjectContext {
	moduleName?: string;
	moduleKind?: ModuleSymbolKind;
	/** Office host of the module's container; absent means Excel. */
	host?: VbaHostToken;
	/** The host's object model, when the host is not Excel (issue #24:
	 * absent lets every resolver keep its Excel default). */
	hostModel?: HostObjectModel;
	documentType?: EventHandlerDocumentType;
	codeNameMap?: Record<string, string>;
	codeNameList?: string[];
	meType?: string;
	meProjectType?: string;
	projectTypes?: TypeCompletionContext['projectTypes'];
	projectClassMembers?: MemberCompletionContext['projectClassMembers'];
	/** A form's designer-declared controls, when the module text carries them. */
	implicitMembers?: MemberCompletionContext['implicitMembers'];
	projectProcedures?: readonly VbaProcedureSignature[];
	projectSymbols?: IdentifierCompletionContext['projectSymbols'];
}

interface CachedEditorProjectContext {
	documentVersion: number;
	loadedAt: number;
	context: EditorProjectContext;
}

interface EditorProjectContextBuild {
	documentVersion: number;
	promise: Promise<EditorProjectContext>;
}

/** Maps a document module to the host type that `Me` denotes inside it. */
function meTypeFor(entry: ModuleEntry | undefined, host?: VbaHostToken): string | undefined {
	if (host === 'vb6') {
		// A VB6 form is a VB.Form (or an MDIForm, as its designer says), not
		// an MSForms.UserForm: its `Me` reaches Caption, Show, Hide and the
		// rest of that surface from the vb6 model. A VB6 project has no
		// document modules, so nothing else carries a `Me` type here.
		if (entry?.type !== 'userform') {
			return undefined;
		}
		return entry.designerClass === VB6_MDI_FORM ? VB6_MDI_FORM : VB6_FORM;
	}
	if (entry?.type === 'userform') {
		// A form IS an MSForms.UserForm, so `Me.` reaches Caption, Controls and
		// the rest of that surface as well as the form's own code. Forms are
		// host-independent.
		return USERFORM;
	}
	if (!entry || entry.type !== 'document') {
		return undefined;
	}
	if (host === 'word') {
		return WORD_DOCUMENT;
	}
	if (host !== undefined && host !== 'excel') {
		// PowerPoint has no document modules and other hosts' document
		// surfaces are unmodelled; silence beats a wrong Excel type.
		return undefined;
	}
	switch (documentTypeFor(entry)) {
		case 'workbook':
			return WORKBOOK;
		case 'chart':
			return CHART;
		default:
			return WORKSHEET;
	}
}

/** Maps `Me` to the source-backed current object module when applicable. */
function meProjectTypeFor(entry: ModuleEntry | undefined): string | undefined {
	if (!entry || !['class', 'document', 'userform'].includes(entry.type)) {
		return undefined;
	}
	return entry.name;
}

/**
 * Builds the lowercased code-name -> host type map for a project. Exported so
 * the semantic-token provider can type code-name receivers (issue #29) from
 * the same host mapping completion and hover use.
 */
export function codeNameHostTypesForModules(
	entries: readonly ModuleEntry[],
	host?: VbaHostToken,
): Record<string, string> {
	const out: Record<string, string> = {};
	for (const entry of entries) {
		if (entry.type !== 'document') {
			continue;
		}
		const meType = meTypeFor(entry, host)
			?? (host === undefined || host === 'excel' ? WORKSHEET : undefined);
		if (meType) {
			out[entry.name.toLowerCase()] = meType;
		}
	}
	return out;
}

function codeNameListFor(entries: ModuleEntry[]): string[] {
	return entries
		.filter((entry) => entry.type === 'document')
		.map((entry) => entry.name);
}

function documentTypeFor(entry: ModuleEntry | undefined): EventHandlerDocumentType | undefined {
	if (!entry || entry.type !== 'document') {
		return undefined;
	}
	return eventHandlerDocumentTypeForContext({
		moduleName: entry.name,
		moduleKind: 'document',
		documentType: entry.documentType,
	});
}

function localDocumentTypeFromModuleName(
	moduleName: string,
	host?: VbaHostToken,
): EventHandlerDocumentType | undefined {
	if (host === 'word') {
		return /^thisdocument$/i.test(moduleName) ? 'document' : undefined;
	}
	if (host !== undefined && host !== 'excel') {
		return undefined;
	}
	if (/^thisworkbook$/i.test(moduleName)) {
		return 'workbook';
	}
	if (/^chart\d*$/i.test(moduleName)) {
		return 'chart';
	}
	if (/^sheet\d+$/i.test(moduleName)) {
		return 'worksheet';
	}
	return undefined;
}

export function toMemberCompletionContext(ctx: EditorProjectContext): MemberCompletionContext {
	return {
		codeNames: ctx.codeNameMap,
		meType: ctx.meType,
		meProjectType: ctx.meProjectType,
		projectClassMembers: ctx.projectClassMembers,
		implicitMembers: ctx.implicitMembers,
		model: ctx.hostModel,
	};
}

export function toTypeCompletionContext(ctx: EditorProjectContext): TypeCompletionContext {
	return {
		projectTypes: ctx.projectTypes,
		model: ctx.hostModel,
	};
}

export function toIdentifierCompletionContext(ctx: EditorProjectContext): IdentifierCompletionContext {
	return {
		codeNames: ctx.codeNameList,
		codeNameTypes: ctx.codeNameMap,
		moduleName: ctx.moduleName,
		moduleKind: ctx.moduleKind,
		projectMemberSurfaces: ctx.projectClassMembers,
		projectProcedures: ctx.projectProcedures,
		projectSymbols: ctx.projectSymbols,
		model: ctx.hostModel,
	};
}

export function toEventHandlerCompletionContext(ctx: EditorProjectContext): EventHandlerCompletionContext {
	return {
		moduleName: ctx.moduleName,
		moduleKind: ctx.moduleKind,
		documentType: ctx.documentType,
		// A VB6 form's handlers come from its model, not from a table: the
		// form's own class and its controls decide which stubs exist.
		host: ctx.host,
		model: ctx.hostModel,
		meType: ctx.meType,
		implicitMembers: ctx.implicitMembers,
	};
}

export class VbaEditorProjectContextService {
	private readonly _projectContextCache = new Map<string, CachedEditorProjectContext>();
	private readonly _projectContextBuilds = new Map<string, EditorProjectContextBuild>();

	constructor(private readonly _projectIndexService: VbaProjectIndexService) {}

	/** Drop derived editor contexts for a project (e.g. after a project change). */
	invalidate(projectPath?: string): void {
		if (projectPath === undefined) {
			this._projectContextCache.clear();
			this._projectContextBuilds.clear();
		} else {
			this._clearProjectContextCacheForPath(projectPath);
		}
	}

	cachedEditorProjectContext(document: vscode.TextDocument): EditorProjectContext | undefined {
		const cached = this._projectContextCache.get(document.uri.toString());
		if (
			!cached ||
			cached.documentVersion !== document.version ||
			Date.now() - cached.loadedAt > EDITOR_PROJECT_CONTEXT_CACHE_TTL_MS
		) {
			return undefined;
		}
		return cached.context;
	}

	private _storeEditorProjectContext(
		document: vscode.TextDocument,
		context: EditorProjectContext,
		documentVersion = document.version,
	): EditorProjectContext {
		const key = document.uri.toString();
		const existing = this._projectContextCache.get(key);
		if (existing && existing.documentVersion > documentVersion) {
			return existing.context;
		}
		this._projectContextCache.set(key, {
			documentVersion,
			loadedAt: Date.now(),
			context,
		});
		this._pruneEditorProjectContextCache();
		return context;
	}

	private _isCurrentProjectContextBuild(document: vscode.TextDocument, documentVersion: number): boolean {
		const build = this._projectContextBuilds.get(document.uri.toString());
		return !build || build.documentVersion === documentVersion;
	}

	private _pruneEditorProjectContextCache(): void {
		const openKeys = new Set(vscode.workspace.textDocuments.map((document) => document.uri.toString()));
		for (const key of this._projectContextCache.keys()) {
			if (!openKeys.has(key)) {
				this._projectContextCache.delete(key);
			}
		}
		const overflow = this._projectContextCache.size - EDITOR_PROJECT_CONTEXT_CACHE_MAX_DOCUMENTS;
		if (overflow <= 0) {
			return;
		}
		for (const key of [...this._projectContextCache.keys()].slice(0, overflow)) {
			this._projectContextCache.delete(key);
		}
	}

	private _clearProjectContextCacheForPath(projectPath: string): void {
		const projectKey = projectIdentityKey(projectPath);
		for (const key of [...this._projectContextCache.keys()]) {
			const location = moduleLocationOfUri(vscode.Uri.parse(key));
			if (!location || projectIdentityKey(location.projectPath) === projectKey) {
				this._projectContextCache.delete(key);
			}
		}
		for (const key of [...this._projectContextBuilds.keys()]) {
			const location = moduleLocationOfUri(vscode.Uri.parse(key));
			if (!location || projectIdentityKey(location.projectPath) === projectKey) {
				this._projectContextBuilds.delete(key);
			}
		}
	}

	private async _buildEditorProjectContext(
		document: vscode.TextDocument,
		source: string,
	): Promise<EditorProjectContext> {
		const cached = this.cachedEditorProjectContext(document);
		if (cached) {
			return cached;
		}
		const buildKey = document.uri.toString();
		const existingBuild = this._projectContextBuilds.get(buildKey);
		if (existingBuild?.documentVersion === document.version) {
			return existingBuild.promise;
		}
		const documentVersion = document.version;
		const build = this._computeEditorProjectContext(document, source, documentVersion)
			.finally(() => {
				if (this._projectContextBuilds.get(buildKey)?.promise === build) {
					this._projectContextBuilds.delete(buildKey);
				}
			});
		this._projectContextBuilds.set(buildKey, { documentVersion, promise: build });
		return build;
	}

	private async _computeEditorProjectContext(
		document: vscode.TextDocument,
		source: string,
		documentVersion: number,
	): Promise<EditorProjectContext> {
		const location = moduleLocationOfDocument(document);
		if (!location) {
			try {
				const project = await buildLiveVbaProjectIndexAsync(
					[{ moduleName: 'Module', moduleKind: 'standard', source: blankDesignerHeader(source) }],
				);
				if (!this._isCurrentProjectContextBuild(document, documentVersion)) {
					return this.cachedEditorProjectContext(document) ?? {};
				}
				const context = projectEditorSymbolContextForModule(project, 'Module');
				return this._storeEditorProjectContext(document, {
					moduleName: 'Module',
					moduleKind: 'standard',
					projectTypes: context.analysisOptions.projectTypes,
					projectClassMembers: context.analysisOptions.projectClassMembers,
					implicitMembers: context.analysisOptions.implicitMembers,
					projectProcedures: context.externalProjectProcedures,
					projectSymbols: context.externalProjectSymbols,
				}, documentVersion);
			} catch {
				return {};
			}
		}

		try {
			const decoded = location;
			const host = hostTokenForFileName(decoded.projectPath);
			// The shared project context already folds in the open editors'
			// text (including this document) one changed module at a time.
			const projectContext = await this._projectIndexService.contextForProject(
				decoded.projectPath,
				'live',
			);
			if (!this._isCurrentProjectContextBuild(document, documentVersion)) {
				return this.cachedEditorProjectContext(document) ?? {};
			}
			const allEntries: ModuleEntry[] = [...projectContext.moduleMetadata.values()].map(
				(metadata) => ({
					name: metadata.moduleName,
					type: metadata.moduleType ?? 'standard',
					documentType: metadata.documentType,
					designerClass: metadata.designerClass,
				}),
			);
			const current = allEntries.find(
				(entry) => entry.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			const moduleKind = moduleKindFromType(current?.type);
			const context = projectEditorSymbolContextForModule(
				projectContext.project,
				decoded.moduleName,
			);
			return this._storeEditorProjectContext(document, {
				moduleName: decoded.moduleName,
				moduleKind,
				host,
				hostModel: hostObjectModelForToken(host),
				documentType: documentTypeFor(current),
				codeNameMap: codeNameHostTypesForModules(allEntries, host),
				codeNameList: codeNameListFor(allEntries),
				meType: meTypeFor(current, host),
				meProjectType: meProjectTypeFor(current),
				projectTypes: context.analysisOptions.projectTypes,
				projectClassMembers: context.analysisOptions.projectClassMembers,
				implicitMembers: context.analysisOptions.implicitMembers,
				projectProcedures: context.externalProjectProcedures,
				projectSymbols: context.externalProjectSymbols,
			}, documentVersion);
		} catch {
			return {};
		}
	}

	warmEditorProjectContext(document: vscode.TextDocument, source: string): void {
		if (this._projectContextBuilds.has(document.uri.toString())) {
			return;
		}
		void this._buildEditorProjectContext(document, source).catch(() => {
			/* best-effort cache warm */
		});
	}

	async buildEditorProjectContextWithin(
		document: vscode.TextDocument,
		source: string,
		timeoutMs: number,
	): Promise<EditorProjectContext | undefined> {
		const build = this._buildEditorProjectContext(document, source);
		let timeout: ReturnType<typeof setTimeout> | undefined;
		try {
			return await new Promise<EditorProjectContext | undefined>((resolve) => {
				let settled = false;
				const finish = (value: EditorProjectContext | undefined): void => {
					if (settled) {
						return;
					}
					settled = true;
					if (timeout) {
						clearTimeout(timeout);
					}
					resolve(value);
				};
				timeout = setTimeout(() => finish(undefined), timeoutMs);
				build.then(finish, () => finish(undefined));
			});
		} finally {
			if (timeout) {
				clearTimeout(timeout);
			}
		}
	}

	cheapEditorProjectContext(document: vscode.TextDocument): EditorProjectContext {
		return this._localModuleIdentity(document);
	}

	localEditorProjectContext(
		document: vscode.TextDocument,
		source: string,
	): EditorProjectContext {
		const identity = this._localModuleIdentity(document);
		try {
			const project = buildLiveVbaProjectIndex([{
				moduleName: identity.moduleName,
				moduleKind: identity.moduleKind,
				source,
			}]);
			const context = projectEditorSymbolContextForModule(project, identity.moduleName);
			return {
				moduleName: identity.moduleName,
				moduleKind: identity.moduleKind,
				host: identity.host,
				hostModel: identity.hostModel,
				documentType: identity.documentType,
				meType: identity.meType,
				meProjectType: identity.meProjectType,
				projectTypes: context.analysisOptions.projectTypes,
				projectClassMembers: context.analysisOptions.projectClassMembers,
				implicitMembers: context.analysisOptions.implicitMembers,
				projectProcedures: context.externalProjectProcedures,
				projectSymbols: context.externalProjectSymbols,
			};
		} catch {
			return {
				moduleName: identity.moduleName,
				moduleKind: identity.moduleKind,
				host: identity.host,
				hostModel: identity.hostModel,
				documentType: identity.documentType,
				meType: identity.meType,
				meProjectType: identity.meProjectType,
			};
		}
	}

	private _localModuleIdentity(document: vscode.TextDocument): {
		moduleName: string;
		moduleKind: ModuleSymbolKind;
		host?: VbaHostToken;
		hostModel?: HostObjectModel;
		documentType?: EventHandlerDocumentType;
		meType?: string;
		meProjectType?: string;
	} {
		let moduleName = 'Module';
		let host: VbaHostToken | undefined;
		const location = moduleLocationOfDocument(document);
		if (location) {
			moduleName = location.moduleName;
			host = hostTokenForFileName(location.projectPath);
		}
		const documentType = localDocumentTypeFromModuleName(moduleName, host);
		const moduleKind: ModuleSymbolKind = documentType ? 'document' : 'standard';
		return {
			moduleName,
			moduleKind,
			host,
			hostModel: hostObjectModelForToken(host),
			documentType,
			meType: documentType === 'workbook'
				? WORKBOOK
				: documentType === 'chart'
					? CHART
					: documentType === 'worksheet'
						? WORKSHEET
						: documentType === 'document'
							? WORD_DOCUMENT
							: undefined,
			meProjectType: moduleKind === 'document' ? moduleName : undefined,
		};
	}
}
