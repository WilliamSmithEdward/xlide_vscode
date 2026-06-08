// Host-context member completion provider.
//
// Registers a CompletionItemProvider triggered on '.' that resolves the type of
// the receiver expression (ThisWorkbook, Application, ActiveSheet, a worksheet
// code name, Me, or a typed local/module variable) and offers the verified
// Excel object-model members for that type. The same provider also offers
// type-name completions in a declaration type position (after `As` / `As New`)
// and a HoverProvider that describes the identifier under the cursor. See the
// Host-Context Member Completion addendum and Phases 6-7 in
// docs/xlide_vba_language_service_roadmap.md.

import * as vscode from 'vscode';
import { PythonBridge } from './pythonBridge';
import {
	XLIDE_SCHEME,
	XLIDE_VBA_LANGUAGE_ID,
	decodeModuleUri,
	moduleIdentityKey,
	workbookIdentityKey,
} from './xlideFileSystem';
import { openModuleSourceForWorkbook } from './vbaOpenDocuments';
import { leadingWhitespace, procedureHeaderParensEdit } from './vbaStructuralAnalysis';
import { xlideEditorBlockLayoutFromConfig } from './globalSettings';
import {
	DocRegistry,
	EventHandlerCompletion,
	EventHandlerCompletionContext,
	EventHandlerDocumentType,
	eventHandlerDocumentTypeForContext,
	getHostType,
	HoverContext,
	IdentifierCompletion,
	IdentifierCompletionContext,
	KeywordCompletion,
	MemberCompletion,
	MemberCompletionContext,
	ModuleSymbolKind,
	materializeKeywordSnippet,
	VbaProcedureSignature,
	callableCompletionShouldInsertParens,
	type CanonicalCaseContext,
	type CanonicalCaseEdit,
	canonicalCaseBoundaryKind,
	resolveCanonicalCaseEdit,
	resolveCanonicalCaseEdits,
	resolveEventHandlerCompletions,
	resolveHover,
	resolveIdentifierCompletions,
	resolveKeywordCompletions,
	resolveMemberCompletions,
	resolveProcedureLabelCompletions,
	resolveSignatureHelp,
	resolveTypeCompletions,
	SignatureHelpContext,
	TypeCompletion,
	TypeCompletionContext,
	type VbaProcedureLabelCompletion,
} from './analyzer';
import {
	buildLiveVbaProjectIndex,
	buildVbaProjectIndex,
	moduleKindFromType,
	projectEditorSymbolContextForModule,
	type VbaProjectLiveOverride,
	type VbaProjectModuleInput,
} from './vbaProjectAnalysis';
import {
	resolveVbaTestDirectiveCompletions,
	type VbaTestDirectiveCompletion,
} from './vbaTestDirectiveCompletion';

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const CHART = 'Excel.Chart';
const MODULE_CACHE_TTL_MS = 5000;
const MODULE_SOURCE_CACHE_TTL_MS = 10_000;
const KEYWORD_SNIPPET_ACCEPTED_COMMAND = 'xlide.vba.keywordSnippetAccepted';
const KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS = 150;
const MAX_PENDING_CANONICAL_CASE_REQUESTS = 16;
const EDITOR_PROJECT_CONTEXT_CACHE_TTL_MS = 10_000;
const CANONICAL_LINE_IDLE_DELAY_MS = 200;
const HOVER_PROJECT_CONTEXT_BUDGET_MS = 120;
const SIGNATURE_HELP_PROJECT_CONTEXT_BUDGET_MS = 150;

interface ModuleEntry {
	name: string;
	type: string;
	documentType?: EventHandlerDocumentType;
}

interface ModuleSourceEntry extends ModuleEntry {
	source: string;
}

interface CachedModules {
	entries: ModuleEntry[];
	loadedAt: number;
}

interface CachedModuleSource {
	source: string;
	loadedAt: number;
}

interface EditorProjectContext {
	moduleName?: string;
	moduleKind?: ModuleSymbolKind;
	documentType?: EventHandlerDocumentType;
	codeNameMap?: Record<string, string>;
	codeNameList?: string[];
	meType?: string;
	meProjectType?: string;
	projectTypes?: TypeCompletionContext['projectTypes'];
	projectClassMembers?: MemberCompletionContext['projectClassMembers'];
	projectProcedures?: readonly VbaProcedureSignature[];
	projectSymbols?: IdentifierCompletionContext['projectSymbols'];
}

type CanonicalCaseRequest = {
	document: vscode.TextDocument;
	editorHint?: vscode.TextEditor;
	resolveEdits: (source: string, ctx: CanonicalCaseContext) => CanonicalCaseEdit[];
};

interface CachedEditorProjectContext {
	documentVersion: number;
	loadedAt: number;
	context: EditorProjectContext;
}

interface CanonicalLineOptions {
	completeProcedureHeader?: boolean;
}

/** Maps a document module to the host type that `Me` denotes inside it. */
function meTypeFor(entry: ModuleEntry | undefined): string | undefined {
	if (!entry || entry.type !== 'document') {
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

/** Builds the lowercased code-name -> host type map for a workbook project. */
function codeNamesFor(entries: ModuleEntry[]): Record<string, string> {
	const out: Record<string, string> = {};
	for (const entry of entries) {
		if (entry.type !== 'document') {
			continue;
		}
		out[entry.name.toLowerCase()] = meTypeFor(entry) ?? WORKSHEET;
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

function localDocumentTypeFromModuleName(moduleName: string): EventHandlerDocumentType | undefined {
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

class VbaMemberCompletionProvider
	implements
		vscode.CompletionItemProvider,
		vscode.HoverProvider,
		vscode.SignatureHelpProvider
{
	private readonly _cache = new Map<string, CachedModules>();
	private readonly _moduleListReads = new Map<string, Promise<ModuleEntry[] | undefined>>();
	private readonly _moduleSourceCache = new Map<string, CachedModuleSource>();
	private readonly _moduleSourceReads = new Map<string, Promise<string | undefined>>();
	private readonly _workbookSourceReads = new Map<string, Promise<ModuleSourceEntry[] | undefined>>();
	private readonly _projectContextCache = new Map<string, CachedEditorProjectContext>();
	private readonly _projectContextBuilds = new Map<string, Promise<EditorProjectContext>>();
	private readonly _pendingCanonicalCaseRequests: CanonicalCaseRequest[] = [];
	private _applyingCanonicalCase = false;

	constructor(
		private readonly _bridge: PythonBridge,
		private readonly _docs?: DocRegistry,
	) {}

	/** Drop cached module lists for a workbook (e.g. after a project change). */
	invalidate(xlsmPath?: string): void {
		if (xlsmPath === undefined) {
			this._cache.clear();
			this._moduleListReads.clear();
			this._moduleSourceCache.clear();
			this._moduleSourceReads.clear();
			this._workbookSourceReads.clear();
			this._projectContextCache.clear();
			this._projectContextBuilds.clear();
		} else {
			const key = workbookIdentityKey(xlsmPath);
			this._cache.delete(key);
			this._moduleListReads.delete(key);
			this._workbookSourceReads.delete(key);
			this._clearModuleSourceCacheForWorkbook(key);
			this._clearProjectContextCacheForWorkbook(xlsmPath);
		}
	}

	async provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.CompletionItem[]> {
		const directiveCompletions = this._testDirectiveCompletions(document, position);
		const directiveItems = directiveCompletions.map(
			(completion) => this._toTestDirectiveItem(completion, position.line),
		);
		if (directiveCompletions.some((completion) => completion.exclusive)) {
			return directiveItems;
		}

		const source = document.getText();
		const offset = document.offsetAt(position);
		const range = this._completionRange(document, position);

		const quickTypes = resolveTypeCompletions(source, offset, {});
		if (quickTypes.length > 0) {
			return quickTypes.map((t) => this._toTypeItem(t, range));
		}

		const cachedProjectCtx = this._cachedEditorProjectContext(document);
		const fastProjectCtx = cachedProjectCtx ?? this._localEditorProjectContext(document, source);
		if (!cachedProjectCtx && document.uri.scheme === XLIDE_SCHEME) {
			this._warmEditorProjectContext(document, source);
		}

		const fastTypes = resolveTypeCompletions(source, offset, this._typeContext(fastProjectCtx));
		if (fastTypes.length > 0) {
			return fastTypes.map((t) => this._toTypeItem(t, range));
		}

		const fastMembers = resolveMemberCompletions(source, offset, this._memberContext(fastProjectCtx));
		if (fastMembers.length > 0) {
			return fastMembers.map((mem) => this._toItem(mem, range, source, offset));
		}

		const fastEvents = resolveEventHandlerCompletions(source, offset, this._eventHandlerContext(fastProjectCtx));
		if (fastEvents.length > 0) {
			return fastEvents.map((event) => this._toEventHandlerItem(event, range));
		}

		const projectCtx = cachedProjectCtx ?? await this._buildEditorProjectContext(document, source);
		const typeCtx = this._typeContext(projectCtx);
		const types = resolveTypeCompletions(source, offset, typeCtx);
		if (types.length > 0) {
			return types.map((t) => this._toTypeItem(t, range));
		}

		const memberCtx = this._memberContext(projectCtx);
		const members = resolveMemberCompletions(source, offset, memberCtx);
		if (members.length > 0) {
			return members.map((mem) => this._toItem(mem, range, source, offset));
		}

		const eventCtx = this._eventHandlerContext(projectCtx);
		const events = resolveEventHandlerCompletions(source, offset, eventCtx);
		if (events.length > 0) {
			return events.map((event) => this._toEventHandlerItem(event, range));
		}

		const labels = resolveProcedureLabelCompletions(source, offset);
		if (labels.length > 0) {
			return labels.map((label) => this._toProcedureLabelItem(label, range));
		}

		const keywords = resolveKeywordCompletions(source, offset, {
			blockLayout: xlideEditorBlockLayoutFromConfig(vscode.workspace.getConfiguration('xlide')).value,
		});
		if (keywords.exclusive) {
			return [
				...directiveItems,
				...keywords.items.map((item) => this._toKeywordItem(item, range, document)),
			];
		}

		const identCtx = this._identifierContext(projectCtx);
		const idents = resolveIdentifierCompletions(source, offset, identCtx);
		return [
			...directiveItems,
			...idents.map((id) => this._toIdentItem(id, range, source, offset)),
			...keywords.items.map((item) => this._toKeywordItem(item, range, document)),
		];
	}

	async applyCanonicalCase(
		document: vscode.TextDocument,
		candidateEnd: vscode.Position,
		editorHint?: vscode.TextEditor,
	): Promise<void> {
		await this._applyCanonicalCaseEdits(document, editorHint, (source, ctx) => {
			const offset = document.offsetAt(candidateEnd);
			const edit = resolveCanonicalCaseEdit(source, offset, ctx);
			return edit ? [edit] : [];
		});
	}

	async applyCanonicalCaseForLine(
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
		options: CanonicalLineOptions = {},
	): Promise<void> {
		if (lineNumber < 0 || lineNumber >= document.lineCount) {
			return;
		}
		await this._applyCanonicalCaseEdits(document, editorHint, (source, ctx) => {
			if (lineNumber >= document.lineCount) {
				return [];
			}
			const line = document.lineAt(lineNumber);
			const start = document.offsetAt(line.range.start);
			const end = document.offsetAt(line.range.end);
			const edits = resolveCanonicalCaseEdits(source, { start, end }, ctx);
			if (options.completeProcedureHeader) {
				const headerEdit = procedureHeaderParensEdit(line.text);
				if (headerEdit) {
					edits.push({
						start: start + headerEdit.startCol,
						end: start + headerEdit.endCol,
						text: headerEdit.newText,
					});
				}
			}
			return edits;
		});
	}

	private async _applyCanonicalCaseEdits(
		document: vscode.TextDocument,
		editorHint: vscode.TextEditor | undefined,
		resolveEdits: (source: string, ctx: CanonicalCaseContext) => CanonicalCaseEdit[],
	): Promise<void> {
		if (this._applyingCanonicalCase) {
			this._enqueueCanonicalCaseRequest({ document, editorHint, resolveEdits });
			return;
		}
		this._applyingCanonicalCase = true;
		try {
			const editor = editorHint?.document === document
				? editorHint
				: vscode.window.visibleTextEditors.find((candidate) => candidate.document === document);
			if (!editor || editor.document !== document) {
				return;
			}
			const source = document.getText();
			const projectCtx = this._cachedEditorProjectContext(document) ?? {};
			const edits = resolveEdits(source, {
				member: this._memberContext(projectCtx),
				identifier: this._identifierContext(projectCtx),
				type: this._typeContext(projectCtx),
			}).filter((edit) => {
				const range = new vscode.Range(
					document.positionAt(edit.start),
					document.positionAt(edit.end),
				);
				return document.getText(range) !== edit.text;
			});
			if (edits.length === 0) {
				return;
			}
			await editor.edit((builder) => {
				for (const edit of edits) {
					builder.replace(
						new vscode.Range(
							document.positionAt(edit.start),
							document.positionAt(edit.end),
						),
						edit.text,
					);
				}
			}, {
				undoStopBefore: false,
				undoStopAfter: false,
			});
		} finally {
			this._applyingCanonicalCase = false;
			const next = this._pendingCanonicalCaseRequests.shift();
			if (next) {
				void this._applyCanonicalCaseEdits(next.document, next.editorHint, next.resolveEdits);
			}
		}
	}

	private _enqueueCanonicalCaseRequest(request: CanonicalCaseRequest): void {
		this._pendingCanonicalCaseRequests.push(request);
		const overflow = this._pendingCanonicalCaseRequests.length - MAX_PENDING_CANONICAL_CASE_REQUESTS;
		if (overflow > 0) {
			this._pendingCanonicalCaseRequests.splice(0, overflow);
		}
	}

	private _cachedEditorProjectContext(document: vscode.TextDocument): EditorProjectContext | undefined {
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
		this._projectContextCache.set(document.uri.toString(), {
			documentVersion,
			loadedAt: Date.now(),
			context,
		});
		return context;
	}

	private _clearProjectContextCacheForWorkbook(xlsmPath: string): void {
		const workbookKey = workbookIdentityKey(xlsmPath);
		for (const key of [...this._projectContextCache.keys()]) {
			try {
				const decoded = decodeModuleUri(vscode.Uri.parse(key));
				if (workbookIdentityKey(decoded.xlsmPath) === workbookKey) {
					this._projectContextCache.delete(key);
				}
			} catch {
				this._projectContextCache.delete(key);
			}
		}
		for (const key of [...this._projectContextBuilds.keys()]) {
			try {
				const uriKey = key.slice(0, key.lastIndexOf(':'));
				const decoded = decodeModuleUri(vscode.Uri.parse(uriKey));
				if (workbookIdentityKey(decoded.xlsmPath) === workbookKey) {
					this._projectContextBuilds.delete(key);
				}
			} catch {
				this._projectContextBuilds.delete(key);
			}
		}
	}

	async provideHover(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.Hover | undefined> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const cached = this._cachedEditorProjectContext(document);
		const fastCtx = cached ?? this._cheapEditorProjectContext(document);
		if (!cached && document.uri.scheme === XLIDE_SCHEME) {
			this._warmEditorProjectContext(document, source);
		}
		let info = resolveHover(source, offset, this._hoverContext(fastCtx));
		if (!info && !cached) {
			info = resolveHover(source, offset, this._hoverContext(
				this._localEditorProjectContext(document, source),
			));
		}
		if (!info && !cached && document.uri.scheme === XLIDE_SCHEME) {
			const projectCtx = await this._buildEditorProjectContextWithin(
				document,
				source,
				HOVER_PROJECT_CONTEXT_BUDGET_MS,
			);
			if (projectCtx) {
				info = resolveHover(source, offset, this._hoverContext(projectCtx));
			}
		}
		if (!info) {
			return undefined;
		}
		return this._toHover(info, document);
	}

	async provideSignatureHelp(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.SignatureHelp | undefined> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const cached = this._cachedEditorProjectContext(document);
		const fastCtx = cached ?? this._cheapEditorProjectContext(document);
		if (!cached && document.uri.scheme === XLIDE_SCHEME) {
			this._warmEditorProjectContext(document, source);
		}
		let info = resolveSignatureHelp(source, offset, this._signatureHelpContext(fastCtx, source));
		if (!info && !cached) {
			info = resolveSignatureHelp(
				source,
				offset,
				this._signatureHelpContext(this._localEditorProjectContext(document, source), source),
			);
		}
		if (!info && !cached && document.uri.scheme === XLIDE_SCHEME) {
			const projectCtx = await this._buildEditorProjectContextWithin(
				document,
				source,
				SIGNATURE_HELP_PROJECT_CONTEXT_BUDGET_MS,
			);
			if (projectCtx) {
				info = resolveSignatureHelp(source, offset, this._signatureHelpContext(projectCtx, source));
			}
		}
		if (!info) {
			return undefined;
		}
		const sig = new vscode.SignatureInformation(info.label);
		sig.parameters = info.parameters.map((p) => {
			const pi = new vscode.ParameterInformation(p.label);
			if (p.documentation) {
				pi.documentation = new vscode.MarkdownString(p.documentation);
			}
			return pi;
		});
		if (info.documentation) {
			sig.documentation = new vscode.MarkdownString(info.documentation);
		}
		const help = new vscode.SignatureHelp();
		help.signatures = [sig];
		help.activeSignature = 0;
		help.activeParameter = info.activeParameter;
		return help;
	}

	private _toHover(
		info: NonNullable<ReturnType<typeof resolveHover>>,
		document: vscode.TextDocument,
	): vscode.Hover {
		const md = new vscode.MarkdownString();
		md.appendCodeblock(info.signature, 'vba');
		if (info.documentation) {
			md.appendMarkdown('\n\n');
			md.appendMarkdown(info.documentation);
		}
		if (info.details.length > 0) {
			md.appendMarkdown('\n\n');
			md.appendMarkdown(info.details.join('  \n'));
		}
		const range = new vscode.Range(
			document.positionAt(info.span.start),
			document.positionAt(info.span.end),
		);
		return new vscode.Hover(md, range);
	}

	private _memberContext(ctx: EditorProjectContext): MemberCompletionContext {
		return {
			codeNames: ctx.codeNameMap,
			meType: ctx.meType,
			meProjectType: ctx.meProjectType,
			projectClassMembers: ctx.projectClassMembers,
		};
	}

	private _typeContext(ctx: EditorProjectContext): TypeCompletionContext {
		return {
			projectTypes: ctx.projectTypes,
		};
	}

	private _identifierContext(ctx: EditorProjectContext): IdentifierCompletionContext {
		return {
			codeNames: ctx.codeNameList,
			moduleName: ctx.moduleName,
			moduleKind: ctx.moduleKind,
			projectMemberSurfaces: ctx.projectClassMembers,
			projectProcedures: ctx.projectProcedures,
			projectSymbols: ctx.projectSymbols,
		};
	}

	private _hoverContext(ctx: EditorProjectContext): HoverContext {
		return {
			...this._memberContext(ctx),
			moduleName: ctx.moduleName,
			moduleKind: ctx.moduleKind,
			projectTypes: ctx.projectTypes,
			projectProcedures: ctx.projectProcedures,
			docRegistry: this._docs,
		};
	}

	private _signatureHelpContext(
		ctx: EditorProjectContext,
		source: string,
	): SignatureHelpContext {
		return {
			...this._memberContext(ctx),
			moduleName: ctx.moduleName,
			moduleSource: source,
			projectProcedures: ctx.projectProcedures,
			docRegistry: this._docs,
		};
	}

	private _eventHandlerContext(ctx: EditorProjectContext): EventHandlerCompletionContext {
		return {
			moduleName: ctx.moduleName,
			moduleKind: ctx.moduleKind,
			documentType: ctx.documentType,
		};
	}

	private async _buildEditorProjectContext(
		document: vscode.TextDocument,
		source: string,
	): Promise<EditorProjectContext> {
		const cached = this._cachedEditorProjectContext(document);
		if (cached) {
			return cached;
		}
		const buildKey = `${document.uri.toString()}:${document.version}`;
		const existingBuild = this._projectContextBuilds.get(buildKey);
		if (existingBuild) {
			return existingBuild;
		}
		const documentVersion = document.version;
		const build = this._computeEditorProjectContext(document, source, documentVersion)
			.finally(() => {
				if (this._projectContextBuilds.get(buildKey) === build) {
					this._projectContextBuilds.delete(buildKey);
				}
			});
		this._projectContextBuilds.set(buildKey, build);
		return build;
	}

	private async _computeEditorProjectContext(
		document: vscode.TextDocument,
		source: string,
		documentVersion: number,
	): Promise<EditorProjectContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			try {
				const project = buildLiveVbaProjectIndex(
					[{ moduleName: 'Module', moduleKind: 'standard', source }],
				);
				const context = projectEditorSymbolContextForModule(project, 'Module');
				return this._storeEditorProjectContext(document, {
					moduleName: 'Module',
					moduleKind: 'standard',
					projectTypes: context.analysisOptions.projectTypes,
					projectClassMembers: context.analysisOptions.projectClassMembers,
					projectProcedures: context.externalProjectProcedures,
					projectSymbols: context.externalProjectSymbols,
				}, documentVersion);
			} catch {
				return {};
			}
		}

		try {
			const decoded = decodeModuleUri(document.uri);
			const sourceEntries = await this._loadWorkbookModuleSources(decoded.xlsmPath);
			const entries = sourceEntries ?? await this._loadModules(decoded.xlsmPath);
			const allEntries = entries ?? [];
			const current = allEntries.find(
				(entry) => entry.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			const moduleKind = moduleKindFromType(current?.type);
			const liveOverride = {
				moduleName: decoded.moduleName,
				moduleKind,
				source,
			};
			const project = sourceEntries
				? this._buildProjectIndexFromSourceEntries(sourceEntries, { liveOverride })
				: await this._buildProjectIndexFromEntries(
					decoded.xlsmPath,
					allEntries,
					{
						liveOverride,
					},
				);
			const context = projectEditorSymbolContextForModule(project, decoded.moduleName);
			return this._storeEditorProjectContext(document, {
				moduleName: decoded.moduleName,
				moduleKind,
				documentType: documentTypeFor(current),
				codeNameMap: codeNamesFor(allEntries),
				codeNameList: codeNameListFor(allEntries),
				meType: meTypeFor(current),
				meProjectType: meProjectTypeFor(current),
				projectTypes: context.analysisOptions.projectTypes,
				projectClassMembers: context.analysisOptions.projectClassMembers,
				projectProcedures: context.externalProjectProcedures,
				projectSymbols: context.externalProjectSymbols,
			}, documentVersion);
		} catch {
			return {};
		}
	}

	private _warmEditorProjectContext(document: vscode.TextDocument, source: string): void {
		void this._buildEditorProjectContext(document, source).catch(() => {
			/* best-effort cache warm */
		});
	}

	private async _buildEditorProjectContextWithin(
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

	private _cheapEditorProjectContext(document: vscode.TextDocument): EditorProjectContext {
		return this._localModuleIdentity(document);
	}

	private _localEditorProjectContext(
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
				documentType: identity.documentType,
				meType: identity.meType,
				meProjectType: identity.meProjectType,
				projectTypes: context.analysisOptions.projectTypes,
				projectClassMembers: context.analysisOptions.projectClassMembers,
				projectProcedures: context.externalProjectProcedures,
				projectSymbols: context.externalProjectSymbols,
			};
		} catch {
			return {
				moduleName: identity.moduleName,
				moduleKind: identity.moduleKind,
				documentType: identity.documentType,
				meType: identity.meType,
				meProjectType: identity.meProjectType,
			};
		}
	}

	private _localModuleIdentity(document: vscode.TextDocument): {
		moduleName: string;
		moduleKind: ModuleSymbolKind;
		documentType?: EventHandlerDocumentType;
		meType?: string;
		meProjectType?: string;
	} {
		let moduleName = 'Module';
		if (document.uri.scheme === XLIDE_SCHEME) {
			try {
				moduleName = decodeModuleUri(document.uri).moduleName;
			} catch {
				moduleName = 'Module';
			}
		}
		const documentType = localDocumentTypeFromModuleName(moduleName);
		const moduleKind: ModuleSymbolKind = documentType ? 'document' : 'standard';
		return {
			moduleName,
			moduleKind,
			documentType,
			meType: documentType === 'workbook'
				? WORKBOOK
				: documentType === 'chart'
					? CHART
					: documentType === 'worksheet'
						? WORKSHEET
						: undefined,
			meProjectType: moduleKind === 'document' ? moduleName : undefined,
		};
	}

	private async _loadModules(xlsmPath: string): Promise<ModuleEntry[] | undefined> {
		const key = workbookIdentityKey(xlsmPath);
		const cached = this._cache.get(key);
		if (cached && Date.now() - cached.loadedAt < MODULE_CACHE_TTL_MS) {
			return cached.entries;
		}
		const existingRead = this._moduleListReads.get(key);
		if (existingRead) {
			return existingRead;
		}
		const promise = this._loadModulesFromBridge(xlsmPath, key, cached);
		this._moduleListReads.set(key, promise);
		promise.then(
			() => {
				if (this._moduleListReads.get(key) === promise) {
					this._moduleListReads.delete(key);
				}
			},
			() => {
				if (this._moduleListReads.get(key) === promise) {
					this._moduleListReads.delete(key);
				}
			},
		);
		return promise;
	}

	private async _loadModulesFromBridge(
		xlsmPath: string,
		key: string,
		cached: CachedModules | undefined,
	): Promise<ModuleEntry[] | undefined> {
		try {
			const entries = await this._bridge.call<ModuleEntry[]>('listModules', {
				path: xlsmPath,
			});
			this._cache.set(key, { entries, loadedAt: Date.now() });
			return entries;
		} catch {
			return cached?.entries;
		}
	}

	private async _buildProjectIndexFromEntries(
		xlsmPath: string,
		entries: ModuleEntry[],
		options: {
			liveOverride?: VbaProjectLiveOverride;
			include?: (kind: ModuleSymbolKind) => boolean;
		} = {},
	): Promise<ReturnType<typeof buildVbaProjectIndex>> {
		const modules: VbaProjectModuleInput[] = [];
		const live = options.liveOverride;
		for (const entry of entries) {
			const isOverride =
				live &&
				entry.name.toLowerCase() === live.moduleName.toLowerCase();
			const entryKind = moduleKindFromType(entry.type);
			const moduleKind = isOverride ? live.moduleKind : entryKind;
			if (options.include && !options.include(moduleKind)) {
				continue;
			}
			if (isOverride) {
				continue;
			}
			let source = await this._moduleSource(xlsmPath, entry);
			if (source === undefined) {
				if (moduleKind !== 'class' && moduleKind !== 'document' && moduleKind !== 'userform') {
					continue;
				}
				// The object type name is the component name, so it remains useful
				// even when the module body cannot be read for members/docs.
				source = '';
			}
			modules.push({
				moduleName: entry.name,
				type: entry.type,
				moduleKind,
				documentType: entry.documentType,
				source,
			});
		}
		const liveOverride = live && (!options.include || options.include(live.moduleKind))
			? live
			: undefined;
		return buildLiveVbaProjectIndex(modules, liveOverride);
	}

	private _buildProjectIndexFromSourceEntries(
		entries: ModuleSourceEntry[],
		options: {
			liveOverride?: VbaProjectLiveOverride;
			include?: (kind: ModuleSymbolKind) => boolean;
		} = {},
	): ReturnType<typeof buildVbaProjectIndex> {
		const modules: VbaProjectModuleInput[] = [];
		const live = options.liveOverride;
		for (const entry of entries) {
			const isOverride =
				live &&
				entry.name.toLowerCase() === live.moduleName.toLowerCase();
			const entryKind = moduleKindFromType(entry.type);
			const moduleKind = isOverride ? live.moduleKind : entryKind;
			if (options.include && !options.include(moduleKind)) {
				continue;
			}
			if (isOverride) {
				continue;
			}
			modules.push({
				moduleName: entry.name,
				type: entry.type,
				moduleKind,
				documentType: entry.documentType,
				source: entry.source,
			});
		}
		const liveOverride = live && (!options.include || options.include(live.moduleKind))
			? live
			: undefined;
		return buildLiveVbaProjectIndex(modules, liveOverride);
	}

	private async _loadWorkbookModuleSources(xlsmPath: string): Promise<ModuleSourceEntry[] | undefined> {
		const workbookKey = workbookIdentityKey(xlsmPath);
		const cached = this._cachedWorkbookModuleSources(xlsmPath, workbookKey);
		if (cached) {
			return cached;
		}
		const existingRead = this._workbookSourceReads.get(workbookKey);
		if (existingRead) {
			return existingRead;
		}
		const promise = this._loadWorkbookModuleSourcesFromBridge(xlsmPath, workbookKey);
		this._workbookSourceReads.set(workbookKey, promise);
		promise.then(
			() => {
				if (this._workbookSourceReads.get(workbookKey) === promise) {
					this._workbookSourceReads.delete(workbookKey);
				}
			},
			() => {
				if (this._workbookSourceReads.get(workbookKey) === promise) {
					this._workbookSourceReads.delete(workbookKey);
				}
			},
		);
		return promise;
	}

	private _cachedWorkbookModuleSources(
		xlsmPath: string,
		workbookKey: string,
	): ModuleSourceEntry[] | undefined {
		const cachedModules = this._cache.get(workbookKey);
		if (!cachedModules || Date.now() - cachedModules.loadedAt >= MODULE_CACHE_TTL_MS) {
			return undefined;
		}
		const openDocuments = vscode.workspace.textDocuments ?? [];
		const entries: ModuleSourceEntry[] = [];
		for (const entry of cachedModules.entries) {
			const sourceKey = this._moduleSourceKey(workbookKey, entry.name);
			const cachedSource = this._moduleSourceCache.get(sourceKey);
			if (!cachedSource || Date.now() - cachedSource.loadedAt >= MODULE_SOURCE_CACHE_TTL_MS) {
				return undefined;
			}
			entries.push({
				...entry,
				source: openModuleSourceForWorkbook(xlsmPath, entry.name, openDocuments) ?? cachedSource.source,
			});
		}
		return entries;
	}

	private async _loadWorkbookModuleSourcesFromBridge(
		xlsmPath: string,
		workbookKey: string,
	): Promise<ModuleSourceEntry[] | undefined> {
		try {
			const entries = await this._bridge.call<ModuleSourceEntry[]>('readModules', {
				path: xlsmPath,
			});
			const now = Date.now();
			this._cache.set(workbookKey, {
				entries: entries.map(({ name, type, documentType }) => ({ name, type, documentType })),
				loadedAt: now,
			});
			const openDocuments = vscode.workspace.textDocuments ?? [];
			const withOpenSources = entries.map((entry) => {
				const open = openModuleSourceForWorkbook(xlsmPath, entry.name, openDocuments);
				const source = open ?? entry.source;
				this._moduleSourceCache.set(this._moduleSourceKey(workbookKey, entry.name), {
					source,
					loadedAt: now,
				});
				return { ...entry, source };
			});
			return withOpenSources;
		} catch (err) {
			const message = err instanceof Error ? err.message : String(err);
			if (/Method not found:\s*readModules/i.test(message) ||
				/Unexpected bridge call readModules/i.test(message)) {
				return undefined;
			}
			throw err;
		}
	}

	private async _moduleSource(
		xlsmPath: string,
		entry: ModuleEntry,
	): Promise<string | undefined> {
		const open = openModuleSourceForWorkbook(xlsmPath, entry.name);
		if (open !== undefined) {
			return open;
		}
		const workbookKey = workbookIdentityKey(xlsmPath);
		const sourceKey = this._moduleSourceKey(workbookKey, entry.name);
		const cached = this._moduleSourceCache.get(sourceKey);
		if (cached && Date.now() - cached.loadedAt < MODULE_SOURCE_CACHE_TTL_MS) {
			return cached.source;
		}
		const existingRead = this._moduleSourceReads.get(sourceKey);
		if (existingRead) {
			return existingRead;
		}
		const promise = this._readModuleSourceFromBridge(xlsmPath, entry, sourceKey, cached);
		this._moduleSourceReads.set(sourceKey, promise);
		promise.then(
			() => {
				if (this._moduleSourceReads.get(sourceKey) === promise) {
					this._moduleSourceReads.delete(sourceKey);
				}
			},
			() => {
				if (this._moduleSourceReads.get(sourceKey) === promise) {
					this._moduleSourceReads.delete(sourceKey);
				}
			},
		);
		return promise;
	}

	private async _readModuleSourceFromBridge(
		xlsmPath: string,
		entry: ModuleEntry,
		sourceKey: string,
		cached: CachedModuleSource | undefined,
	): Promise<string | undefined> {
		try {
			const res = await this._bridge.call<{ source: string }>('readModule', {
				path: xlsmPath,
				module: entry.name,
			});
			this._moduleSourceCache.set(sourceKey, {
				source: res.source,
				loadedAt: Date.now(),
			});
			return res.source;
		} catch {
			return cached?.source;
		}
	}

	private _moduleSourceKey(workbookKey: string, moduleName: string): string {
		return `${workbookKey}\n${moduleIdentityKey(moduleName)}`;
	}

	private _clearModuleSourceCacheForWorkbook(workbookKey: string): void {
		const prefix = `${workbookKey}\n`;
		for (const key of this._moduleSourceCache.keys()) {
			if (key.startsWith(prefix)) {
				this._moduleSourceCache.delete(key);
			}
		}
		for (const key of this._moduleSourceReads.keys()) {
			if (key.startsWith(prefix)) {
				this._moduleSourceReads.delete(key);
			}
		}
	}

	private _toItem(
		mem: MemberCompletion,
		range: vscode.Range,
		source: string,
		offset: number,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(mem.name, this._memberItemKind(mem));
		const ownerName = getHostType(mem.owner)?.displayName ?? mem.owner;
		const kindLabel = mem.kind;
		if (mem.signature) {
			item.detail = `${ownerName}.${mem.signature}`;
		} else {
			item.detail = `${ownerName} ${kindLabel}`;
		}
		if (!mem.signature && mem.returns) {
			const returnName = getHostType(mem.returns)?.displayName ?? mem.returns;
			item.detail += ` -> ${returnName}`;
		}
		if (mem.documentation) {
			item.documentation = new vscode.MarkdownString(mem.documentation);
		}
		this._applyCompletionInsert(
			item,
			mem.name,
			range,
			mem.kind === 'method',
			callableCompletionShouldInsertParens(source, offset),
		);
		return item;
	}

	private _memberItemKind(mem: MemberCompletion): vscode.CompletionItemKind {
		switch (mem.kind) {
			case 'method':
				return vscode.CompletionItemKind.Method;
			case 'event':
				return vscode.CompletionItemKind.Event;
			default:
				return vscode.CompletionItemKind.Property;
		}
	}

	private _toTypeItem(t: TypeCompletion, range: vscode.Range): vscode.CompletionItem {
		const item = new vscode.CompletionItem(t.name, this._typeItemKind(t));
		item.detail = t.detail;
		if (t.documentation) {
			item.documentation = new vscode.MarkdownString(t.documentation);
		}
		this._applyCompletionInsert(item, t.name, range, false);
		return item;
	}

	private _typeItemKind(t: TypeCompletion): vscode.CompletionItemKind {
		switch (t.kind) {
			case 'enum':
				return vscode.CompletionItemKind.Enum;
			case 'external':
				return vscode.CompletionItemKind.Interface;
			case 'host':
			case 'class':
			case 'document':
			case 'userform':
				return vscode.CompletionItemKind.Class;
			case 'module':
				return vscode.CompletionItemKind.Module;
			default:
				return vscode.CompletionItemKind.Struct;
		}
	}

	private _toEventHandlerItem(
		event: EventHandlerCompletion,
		range: vscode.Range,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(event.name, vscode.CompletionItemKind.Event);
		item.detail = event.detail;
		const documentation = new vscode.MarkdownString();
		documentation.appendCodeblock(event.signature, 'vba');
		documentation.appendMarkdown('\n\n');
		documentation.appendMarkdown(event.documentation);
		item.documentation = documentation;
		item.range = range;
		item.filterText = event.name;
		item.sortText = `0:${event.name}`;
		item.insertText = new vscode.SnippetString(event.insertText);
		return item;
	}

	private _toProcedureLabelItem(
		label: VbaProcedureLabelCompletion,
		range: vscode.Range,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(label.label, vscode.CompletionItemKind.Reference);
		item.detail = label.detail;
		item.range = range;
		item.filterText = label.label;
		item.sortText = `1:${label.label}`;
		item.insertText = label.label;
		return item;
	}

	private _toIdentItem(
		id: IdentifierCompletion,
		range: vscode.Range,
		source: string,
		offset: number,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(id.name, this._identItemKind(id));
		item.detail = id.detail;
		if (id.documentation) {
			item.documentation = new vscode.MarkdownString(id.documentation);
		}
		const callable = id.kind === 'runtime' || id.kind === 'procedure';
		this._applyCompletionInsert(
			item,
			id.name,
			range,
			callable,
			callableCompletionShouldInsertParens(source, offset),
		);
		return item;
	}

	private _toKeywordItem(
		keyword: KeywordCompletion,
		range: vscode.Range,
		document: vscode.TextDocument,
	): vscode.CompletionItem {
		const kind = keyword.kind === 'snippet'
			? vscode.CompletionItemKind.Snippet
			: vscode.CompletionItemKind.Keyword;
		const item = new vscode.CompletionItem(keyword.label, kind);
		item.detail = keyword.detail;
		if (keyword.documentation) {
			item.documentation = new vscode.MarkdownString(keyword.documentation);
		}
		item.range = range;
		item.filterText = keyword.filterText ?? keyword.label;
		item.sortText = keyword.sortText ?? `9:${keyword.label}`;
		item.insertText = keyword.kind === 'snippet'
			? new vscode.SnippetString(materializeKeywordSnippet(
				keyword.insertText,
				this._lineIndent(document, range.start.line),
			))
			: keyword.insertText;
		if (keyword.kind === 'snippet') {
			item.keepWhitespace = true;
			item.command = {
				command: KEYWORD_SNIPPET_ACCEPTED_COMMAND,
				title: 'Track VBA Snippet',
			};
		}
		return item;
	}

	private _testDirectiveCompletions(
		document: vscode.TextDocument,
		position: vscode.Position,
	): VbaTestDirectiveCompletion[] {
		return resolveVbaTestDirectiveCompletions(
			document.lineAt(position.line).text,
			position.character,
		);
	}

	private _toTestDirectiveItem(
		completion: VbaTestDirectiveCompletion,
		line: number,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(
			completion.label,
			vscode.CompletionItemKind.Snippet,
		);
		item.detail = completion.detail;
		item.documentation = new vscode.MarkdownString(completion.documentation);
		item.range = new vscode.Range(
			line,
			completion.range.start,
			line,
			completion.range.end,
		);
		item.filterText = `${completion.label} ${completion.label.replace(/^@/, '')}`;
		item.sortText = completion.sortText;
		item.insertText = new vscode.SnippetString(completion.insertText);
		return item;
	}

	private _applyCompletionInsert(
		item: vscode.CompletionItem,
		name: string,
		range: vscode.Range,
		callable: boolean,
		insertParens: boolean = callable,
	): void {
		item.range = range;
		item.filterText = name;
		if (!callable || !insertParens) {
			item.insertText = name;
			return;
		}
		item.insertText = new vscode.SnippetString(`${name}($0)`);
		item.command = {
			command: 'editor.action.triggerParameterHints',
			title: 'Trigger Parameter Hints',
		};
	}

	private _completionRange(
		document: vscode.TextDocument,
		position: vscode.Position,
	): vscode.Range {
		const line = document.lineAt(position.line).text;
		let start = position.character;
		while (start > 0 && /[A-Za-z0-9_]/.test(line[start - 1])) {
			start -= 1;
		}
		if (start === position.character && start > 0 && line[start - 1] === '#') {
			start -= 1;
		}
		let end = position.character;
		while (end < line.length && /[A-Za-z0-9_]/.test(line[end])) {
			end += 1;
		}
		return new vscode.Range(position.line, start, position.line, end);
	}

	private _lineIndent(document: vscode.TextDocument, line: number): string {
		return leadingWhitespace(document.lineAt(line).text);
	}

	private _identItemKind(id: IdentifierCompletion): vscode.CompletionItemKind {
		switch (id.kind) {
			case 'procedure':
				return vscode.CompletionItemKind.Method;
			case 'module':
				return vscode.CompletionItemKind.Module;
			case 'runtime':
				return vscode.CompletionItemKind.Function;
			case 'constant':
				return vscode.CompletionItemKind.Constant;
			case 'value':
				return vscode.CompletionItemKind.Value;
			case 'enum':
				return vscode.CompletionItemKind.Enum;
			case 'enumMember':
				return vscode.CompletionItemKind.EnumMember;
			case 'type':
				return vscode.CompletionItemKind.Struct;
			case 'global':
			case 'codeName':
				return vscode.CompletionItemKind.Variable;
			default:
				return vscode.CompletionItemKind.Variable;
		}
	}
}

const ACTIVE_MEMBER_COMPLETION_PROVIDERS = new Set<VbaMemberCompletionProvider>();

export function invalidateVbaMemberCompletionCache(xlsmPath?: string): void {
	for (const provider of ACTIVE_MEMBER_COMPLETION_PROVIDERS) {
		provider.invalidate(xlsmPath);
	}
}

/** Registers the host-context member completion provider. */
export function registerVbaMemberCompletion(
	context: vscode.ExtensionContext,
	bridge: PythonBridge,
	selector: vscode.DocumentSelector,
	docs?: DocRegistry,
): void {
	const provider = new VbaMemberCompletionProvider(bridge, docs);
	ACTIVE_MEMBER_COMPLETION_PROVIDERS.add(provider);
	context.subscriptions.push({
		dispose: () => ACTIVE_MEMBER_COMPLETION_PROVIDERS.delete(provider),
	});
	let lastCanonicalCandidate = canonicalCandidateFromEditor(vscode.window.activeTextEditor);
	let activeKeywordSnippet:
		| { editor: vscode.TextEditor; documentKey: string; textChangeSerialAtAccept: number }
		| undefined;
	let textChangeSerial = 0;
	const lastTextChange = new Map<string, { at: number; serial: number }>();
	const canonicalLineTimers = new Map<string, ReturnType<typeof setTimeout>>();
	const editorHintFor = (document: vscode.TextDocument): vscode.TextEditor | undefined => {
		const active = vscode.window.activeTextEditor;
		return active?.document === document ? active : undefined;
	};
	const canonicalLineKey = (document: vscode.TextDocument, lineNumber: number): string =>
		`${document.uri.toString()}\n${lineNumber}`;
	const applyCanonicalLine = (
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
		options: CanonicalLineOptions = {},
	): void => {
		void provider.applyCanonicalCaseForLine(document, lineNumber, editorHint, options);
	};
	const applyCanonicalPosition = (
		document: vscode.TextDocument,
		position: vscode.Position,
		editorHint?: vscode.TextEditor,
	): void => {
		void provider.applyCanonicalCase(document, position, editorHint);
	};
	const scheduleCanonicalLine = (
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
		options: CanonicalLineOptions = {},
		delayMs = 0,
	): void => {
		const key = canonicalLineKey(document, lineNumber);
		const existing = canonicalLineTimers.get(key);
		if (existing) {
			clearTimeout(existing);
		}
		const timer = setTimeout(() => {
			canonicalLineTimers.delete(key);
			if (!isVbaDocument(document)) {
				return;
			}
			applyCanonicalLine(document, lineNumber, editorHint, options);
		}, delayMs);
		canonicalLineTimers.set(key, timer);
	};
	const clearCanonicalLineTimers = (document: vscode.TextDocument): void => {
		const prefix = `${document.uri.toString()}\n`;
		for (const [key, timer] of canonicalLineTimers) {
			if (!key.startsWith(prefix)) {
				continue;
			}
			clearTimeout(timer);
			canonicalLineTimers.delete(key);
		}
	};
	const scheduleCanonicalLineIdle = (
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
	): void => {
		scheduleCanonicalLine(document, lineNumber, editorHint, {}, CANONICAL_LINE_IDLE_DELAY_MS);
	};
	const flushCanonicalLine = (): void => {
		const candidate = lastCanonicalCandidate;
		if (!candidate || !isVbaDocument(candidate.editor.document)) {
			return;
		}
		applyCanonicalLine(
			candidate.editor.document,
			candidate.position.line,
			candidate.editor,
			{ completeProcedureHeader: true },
		);
	};
	const markTextChange = (document: vscode.TextDocument): void => {
		if (isVbaDocument(document)) {
			textChangeSerial += 1;
			lastTextChange.set(document.uri.toString(), {
				at: Date.now(),
				serial: textChangeSerial,
			});
		}
	};
	const maybeLeaveKeywordSnippet = (event: vscode.TextEditorSelectionChangeEvent): void => {
		if (!activeKeywordSnippet || event.textEditor !== activeKeywordSnippet.editor) {
			return;
		}
		if (!isVbaDocument(event.textEditor.document)) {
			activeKeywordSnippet = undefined;
			return;
		}
		if (
			event.kind !== vscode.TextEditorSelectionChangeKind.Mouse &&
			event.kind !== vscode.TextEditorSelectionChangeKind.Keyboard
		) {
			return;
		}
		if (event.kind === vscode.TextEditorSelectionChangeKind.Keyboard) {
			const changed = lastTextChange.get(activeKeywordSnippet.documentKey);
			if (
				changed &&
				changed.serial > activeKeywordSnippet.textChangeSerialAtAccept &&
				Date.now() - changed.at <= KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS
			) {
				return;
			}
		}
		activeKeywordSnippet = undefined;
		void vscode.commands.executeCommand('leaveSnippet');
	};

	context.subscriptions.push(
		vscode.commands.registerCommand(KEYWORD_SNIPPET_ACCEPTED_COMMAND, () => {
			const editor = vscode.window.activeTextEditor;
			if (!editor || !isVbaDocument(editor.document)) {
				return;
			}
			activeKeywordSnippet = {
				editor,
				documentKey: editor.document.uri.toString(),
				textChangeSerialAtAccept: textChangeSerial,
			};
		}),
		vscode.languages.registerCompletionItemProvider(
			selector,
			provider,
			'.',
			' ',
			'#',
			'@',
		),
		vscode.workspace.onDidChangeTextDocument((event) => {
			markTextChange(event.document);
			if (!isVbaDocument(event.document)) {
				return;
			}
			const editorHint = editorHintFor(event.document);
			const touchedLines = new Set<number>();
			const immediateLines = new Set<number>();
			for (const change of event.contentChanges) {
				touchedLines.add(Math.min(change.range.start.line, Math.max(0, event.document.lineCount - 1)));
				if (!change.range.isEmpty) {
					continue;
				}
				const boundary = canonicalCaseBoundaryKind(change.text);
				if (!boundary) {
					continue;
				}
				if (boundary === 'line') {
					immediateLines.add(change.range.start.line);
					scheduleCanonicalLine(
						event.document,
						change.range.start.line,
						editorHint,
					);
				} else {
					applyCanonicalPosition(event.document, change.range.start, editorHint);
				}
			}
			for (const lineNumber of touchedLines) {
				if (immediateLines.has(lineNumber)) {
					continue;
				}
				scheduleCanonicalLineIdle(event.document, lineNumber, editorHint);
			}
		}),
		vscode.window.onDidChangeTextEditorSelection((event) => {
			maybeLeaveKeywordSnippet(event);
			const previous = lastCanonicalCandidate;
			if (previous && previous.editor !== event.textEditor) {
				applyCanonicalLine(
					previous.editor.document,
					previous.position.line,
					previous.editor,
					{ completeProcedureHeader: true },
				);
			} else if (previous?.editor === event.textEditor) {
				const nextPosition = event.textEditor.selection.active;
				if (previous.position.line !== nextPosition.line) {
					applyCanonicalLine(
						previous.editor.document,
						previous.position.line,
						previous.editor,
						{ completeProcedureHeader: true },
					);
				} else {
					applyCanonicalPosition(
						previous.editor.document,
						previous.position,
						previous.editor,
					);
				}
			}
			lastCanonicalCandidate = canonicalCandidateFromEditor(event.textEditor);
		}),
		vscode.window.onDidChangeActiveTextEditor((editor) => {
			if (activeKeywordSnippet && editor !== activeKeywordSnippet.editor) {
				activeKeywordSnippet = undefined;
			}
			flushCanonicalLine();
			lastCanonicalCandidate = canonicalCandidateFromEditor(editor);
		}),
		vscode.window.onDidChangeWindowState((state) => {
			if (!state.focused) {
				flushCanonicalLine();
			}
		}),
		vscode.languages.registerHoverProvider(selector, provider),
		vscode.languages.registerSignatureHelpProvider(
			selector,
			provider,
			'(',
			',',
			' ',
		),
		vscode.workspace.onDidSaveTextDocument((doc) => {
			if (doc.uri.scheme !== XLIDE_SCHEME) {
				return;
			}
			try {
				const { xlsmPath } = decodeModuleUri(doc.uri);
				provider.invalidate(xlsmPath);
			} catch {
				// Ignore URIs we cannot decode.
			}
		}),
		vscode.workspace.onDidCloseTextDocument(clearCanonicalLineTimers),
	);
}

function canonicalCandidateFromEditor(
	editor: vscode.TextEditor | undefined,
): { editor: vscode.TextEditor; position: vscode.Position } | undefined {
	if (!editor || !isVbaDocument(editor.document)) {
		return undefined;
	}
	return { editor, position: editor.selection.active };
}

function isVbaDocument(document: vscode.TextDocument): boolean {
	return document.languageId === 'vba'
		|| document.languageId === XLIDE_VBA_LANGUAGE_ID
		|| document.uri.scheme === XLIDE_SCHEME;
}
