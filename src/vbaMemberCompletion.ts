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
import { XLIDE_SCHEME, decodeModuleUri } from './xlideFileSystem';
import { leadingWhitespace } from './vbaLinter';
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
	ProjectIndex,
	VbaProcedureSignature,
	callableCompletionShouldInsertParens,
	resolveCanonicalCaseEdit,
	resolveEventHandlerCompletions,
	resolveHover,
	resolveIdentifierCompletions,
	resolveKeywordCompletions,
	resolveMemberCompletions,
	resolveSignatureHelp,
	resolveTypeCompletions,
	SignatureHelpContext,
	TypeCompletion,
	TypeCompletionContext,
	VbaProjectClassMembers,
} from './analyzer';

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const CHART = 'Excel.Chart';
const MODULE_CACHE_TTL_MS = 5000;
const KEYWORD_SNIPPET_ACCEPTED_COMMAND = 'xlide.vba.keywordSnippetAccepted';
const KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS = 150;

interface ModuleEntry {
	name: string;
	type: string;
	documentType?: EventHandlerDocumentType;
}

interface CachedModules {
	entries: ModuleEntry[];
	loadedAt: number;
}

interface CachedProjectProcedures {
	procedures: VbaProcedureSignature[];
	loadedAt: number;
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

class VbaMemberCompletionProvider
	implements
		vscode.CompletionItemProvider,
		vscode.HoverProvider,
		vscode.SignatureHelpProvider
{
	private readonly _cache = new Map<string, CachedModules>();
	private readonly _procedureCache = new Map<string, CachedProjectProcedures>();
	private _applyingCanonicalCase = false;

	constructor(
		private readonly _bridge: PythonBridge,
		private readonly _docs?: DocRegistry,
	) {}

	/** Drop cached module lists for a workbook (e.g. after a project change). */
	invalidate(xlsmPath?: string): void {
		if (xlsmPath === undefined) {
			this._cache.clear();
			this._procedureCache.clear();
		} else {
			const key = this._key(xlsmPath);
			this._cache.delete(key);
			this._procedureCache.delete(key);
		}
	}

	async provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.CompletionItem[]> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const range = this._completionRange(document, position);

		const memberCtx = await this._buildContext(document);
		const members = resolveMemberCompletions(source, offset, memberCtx);
		if (members.length > 0) {
			return members.map((mem) => this._toItem(mem, range, source, offset));
		}

		const typeCtx = await this._buildTypeContext(document, source);
		const types = resolveTypeCompletions(source, offset, typeCtx);
		if (types.length > 0) {
			return types.map((t) => this._toTypeItem(t, range));
		}

		const eventCtx = await this._buildEventHandlerContext(document);
		const events = resolveEventHandlerCompletions(source, offset, eventCtx);
		if (events.length > 0) {
			return events.map((event) => this._toEventHandlerItem(event, range));
		}

		const keywords = resolveKeywordCompletions(source, offset);
		if (keywords.exclusive) {
			return keywords.items.map((item) => this._toKeywordItem(item, range, document));
		}

		const identCtx = await this._buildIdentifierContext(document);
		const idents = resolveIdentifierCompletions(source, offset, identCtx);
		return [
			...idents.map((id) => this._toIdentItem(id, range, source, offset)),
			...keywords.items.map((item) => this._toKeywordItem(item, range, document)),
		];
	}

	async applyCanonicalCase(
		document: vscode.TextDocument,
		candidateEnd: vscode.Position,
		editorHint?: vscode.TextEditor,
	): Promise<void> {
		if (this._applyingCanonicalCase) {
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
			const offset = document.offsetAt(candidateEnd);
			const edit = resolveCanonicalCaseEdit(source, offset, {
				member: await this._buildContext(document),
				identifier: await this._buildIdentifierContext(document),
				type: await this._buildTypeContext(document, source),
			});
			if (!edit) {
				return;
			}
			const range = new vscode.Range(
				document.positionAt(edit.start),
				document.positionAt(edit.end),
			);
			if (document.getText(range) === edit.text) {
				return;
			}
			const selections = editor.selections.map((selection) =>
				new vscode.Selection(selection.anchor, selection.active),
			);
			const restoreSelection = vscode.window.activeTextEditor === editor;
			const applied = await editor.edit((builder) => builder.replace(range, edit.text), {
				undoStopBefore: false,
				undoStopAfter: false,
			});
			if (applied && restoreSelection && vscode.window.activeTextEditor === editor) {
				editor.selections = selections;
			}
		} finally {
			this._applyingCanonicalCase = false;
		}
	}

	async provideHover(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.Hover | undefined> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const info = resolveHover(source, offset, await this._buildHoverContext(document));
		if (!info) {
			return undefined;
		}
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

	async provideSignatureHelp(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.SignatureHelp | undefined> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const ctx: SignatureHelpContext = {
			...(await this._buildContext(document)),
			moduleSource: source,
			projectProcedures: await this._buildProjectProcedureContext(document),
			docRegistry: this._docs,
		};
		const info = resolveSignatureHelp(source, offset, ctx);
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

	private async _buildHoverContext(
		document: vscode.TextDocument,
	): Promise<HoverContext> {
		const source = document.getText();
		const typeContext = await this._buildTypeContext(document, source);
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return { ...typeContext, docRegistry: this._docs };
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			const current = entries?.find(
				(e) => e.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			return {
				codeNames: codeNamesFor(entries ?? []),
				meType: meTypeFor(current),
				meProjectType: meProjectTypeFor(current),
				moduleName: decoded.moduleName,
				moduleKind: current ? this._moduleKind(current.type) : 'standard',
				projectTypes: typeContext.projectTypes,
				projectClassMembers: entries
					? await this._loadProjectMemberSurfaces(
						decoded.xlsmPath,
						entries,
						decoded.moduleName,
						{
							moduleKind: current ? this._moduleKind(current.type) : 'standard',
							source,
						},
					)
					: undefined,
				projectProcedures: entries
					? await this._loadCrossModuleProcedureSignatures(
						decoded.xlsmPath,
						entries,
						decoded.moduleName.toLowerCase(),
					)
					: undefined,
				docRegistry: this._docs,
			};
		} catch {
			return { ...typeContext, docRegistry: this._docs };
		}
	}

	private async _buildContext(
		document: vscode.TextDocument,
	): Promise<MemberCompletionContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return {};
		}
		let decoded: { xlsmPath: string; moduleName: string };
		try {
			decoded = decodeModuleUri(document.uri);
		} catch {
			return {};
		}
		const entries = await this._loadModules(decoded.xlsmPath);
		if (!entries) {
			return {};
		}
		const current = entries.find(
			(e) => e.name.toLowerCase() === decoded.moduleName.toLowerCase(),
		);
		const moduleKind = current ? this._moduleKind(current.type) : 'standard';
		const projectClassMembers = await this._loadProjectMemberSurfaces(
			decoded.xlsmPath,
			entries,
			decoded.moduleName,
			{ moduleKind, source: document.getText() },
		);
		return {
			codeNames: codeNamesFor(entries),
			meType: meTypeFor(current),
			meProjectType: meProjectTypeFor(current),
			projectClassMembers,
		};
	}

	private async _loadModules(xlsmPath: string): Promise<ModuleEntry[] | undefined> {
		const key = this._key(xlsmPath);
		const cached = this._cache.get(key);
		if (cached && Date.now() - cached.loadedAt < MODULE_CACHE_TTL_MS) {
			return cached.entries;
		}
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

	private async _loadProjectMemberSurfaces(
		xlsmPath: string,
		entries: ModuleEntry[],
		moduleName: string,
		liveOverride?: { moduleKind: ModuleSymbolKind; source: string },
	): Promise<VbaProjectClassMembers[]> {
		const project = await this._buildProjectIndexFromEntries(
			xlsmPath,
			entries,
			liveOverride
				? {
					liveOverride: {
						moduleName,
						moduleKind: liveOverride.moduleKind,
						source: liveOverride.source,
					},
				}
				: {},
		);
		return project.projectMemberSurfaces(moduleName);
	}

	private async _loadCrossModuleProcedureSignatures(
		xlsmPath: string,
		entries: ModuleEntry[],
		currentLower: string,
	): Promise<VbaProcedureSignature[]> {
		const allProcedures = await this._loadProjectProcedureSignatures(xlsmPath, entries);
		return allProcedures.filter(
			(procedure) => procedure.moduleName.toLowerCase() !== currentLower,
		);
	}

	private async _loadProjectProcedureSignatures(
		xlsmPath: string,
		entries: ModuleEntry[],
	): Promise<VbaProcedureSignature[]> {
		const key = this._key(xlsmPath);
		const cached = this._procedureCache.get(key);
		if (cached && Date.now() - cached.loadedAt < MODULE_CACHE_TTL_MS) {
			return cached.procedures;
		}
		const project = await this._buildProjectIndexFromEntries(xlsmPath, entries, {
			include: (kind) => kind === 'standard',
		});
		const procedures = project.visibleProcedureSignatures('__xlide_external_module__');
		this._procedureCache.set(key, { procedures, loadedAt: Date.now() });
		return procedures;
	}

	private async _buildProjectIndexFromEntries(
		xlsmPath: string,
		entries: ModuleEntry[],
		options: {
			liveOverride?: { moduleName: string; moduleKind: ModuleSymbolKind; source: string };
			include?: (kind: ModuleSymbolKind) => boolean;
		} = {},
	): Promise<ProjectIndex> {
		const project = new ProjectIndex();
		let appliedOverride = false;
		for (const entry of entries) {
			const entryKind = this._moduleKind(entry.type);
			const isOverride =
				options.liveOverride &&
				entry.name.toLowerCase() === options.liveOverride.moduleName.toLowerCase();
			const moduleKind = isOverride ? options.liveOverride!.moduleKind : entryKind;
			if (options.include && !options.include(moduleKind)) {
				continue;
			}
			let source = isOverride
				? options.liveOverride!.source
				: await this._moduleSource(xlsmPath, entry);
			if (source === undefined) {
				if (moduleKind !== 'class' && moduleKind !== 'document' && moduleKind !== 'userform') {
					continue;
				}
				// The object type name is the component name, so it remains useful
				// even when the module body cannot be read for members/docs.
				source = '';
			}
			try {
				project.setModule({
					moduleName: isOverride ? options.liveOverride!.moduleName : entry.name,
					moduleKind,
					source,
				});
				appliedOverride = appliedOverride || !!isOverride;
			} catch {
				// Keep IntelliSense alive if one workbook module is temporarily invalid.
			}
		}
		const live = options.liveOverride;
		if (
			live &&
			!appliedOverride &&
			(!options.include || options.include(live.moduleKind))
		) {
			try {
				project.setModule(live);
			} catch {
				// Same recovery rule as above for unsaved/incomplete live text.
			}
		}
		return project;
	}

	private async _moduleSource(
		xlsmPath: string,
		entry: ModuleEntry,
	): Promise<string | undefined> {
		const open = this._openModuleSource(xlsmPath, entry.name);
		if (open !== undefined) {
			return open;
		}
		try {
			const res = await this._bridge.call<{ source: string }>('readModule', {
				path: xlsmPath,
				module: entry.name,
			});
			return res.source;
		} catch {
			return undefined;
		}
	}

	private _openModuleSource(
		xlsmPath: string,
		moduleName: string,
	): string | undefined {
		const workbookKey = this._key(xlsmPath);
		const moduleKey = moduleName.toLowerCase();
		for (const doc of vscode.workspace.textDocuments) {
			if (doc.uri.scheme !== XLIDE_SCHEME) {
				continue;
			}
			try {
				const decoded = decodeModuleUri(doc.uri);
				if (
					this._key(decoded.xlsmPath) === workbookKey &&
					decoded.moduleName.toLowerCase() === moduleKey
				) {
					return doc.getText();
				}
			} catch {
				// Ignore virtual documents outside XLIDE's module URI shape.
			}
		}
		return undefined;
	}

	private async _buildTypeContext(
		document: vscode.TextDocument,
		source: string,
	): Promise<TypeCompletionContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			const project = new ProjectIndex();
			try {
				project.setModule({
					moduleName: 'Module',
					moduleKind: 'standard',
					source,
				});
			} catch {
				return {};
			}
			return { projectTypes: project.visibleTypeNames('Module') };
		}

		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			const current = entries?.find(
				(entry) => entry.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			const moduleKind = current ? this._moduleKind(current.type) : 'standard';
			const project = await this._buildProjectIndexFromEntries(
				decoded.xlsmPath,
				entries ?? [],
				{
					liveOverride: {
						moduleName: decoded.moduleName,
						moduleKind,
						source,
					},
				},
			);
			return { projectTypes: project.visibleTypeNames(decoded.moduleName) };
		} catch {
			return {};
		}
	}

	/**
	 * Builds the event-handler completion context for the current module.
	 * Event procedures are scoped by host document module type and intentionally
	 * stay outside object-member completion.
	 */
	private async _buildEventHandlerContext(
		document: vscode.TextDocument,
	): Promise<EventHandlerCompletionContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return {};
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			const current = entries?.find(
				(e) => e.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			return {
				moduleName: decoded.moduleName,
				moduleKind: current ? this._moduleKind(current.type) : 'standard',
				documentType: documentTypeFor(current),
			};
		} catch {
			return {};
		}
	}

	/**
	 * Builds the identifier-completion context: worksheet/document code names of
	 * the workbook project plus the module currently being edited.
	 */
	private async _buildIdentifierContext(
		document: vscode.TextDocument,
	): Promise<IdentifierCompletionContext> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return {};
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const source = document.getText();
			const entries = await this._loadModules(decoded.xlsmPath);
			const codeNames: string[] = [];
			let moduleKind: ModuleSymbolKind = 'standard';
			const currentLower = decoded.moduleName.toLowerCase();
			for (const entry of entries ?? []) {
				if (entry.type === 'document') {
					codeNames.push(entry.name);
				}
				if (entry.name.toLowerCase() === currentLower) {
					moduleKind = this._moduleKind(entry.type);
				}
			}
			let projectSymbols: IdentifierCompletionContext['projectSymbols'];
			let projectProcedures: IdentifierCompletionContext['projectProcedures'];
			if (entries) {
				const project = await this._buildProjectIndexFromEntries(
					decoded.xlsmPath,
					entries,
					{
						liveOverride: {
							moduleName: decoded.moduleName,
							moduleKind,
							source,
						},
					},
				);
				projectSymbols = project.visibleIdentifierSymbols(decoded.moduleName)
					.filter((symbol) => symbol.moduleName.toLowerCase() !== currentLower);
				projectProcedures = project.visibleProcedureSignatures(decoded.moduleName)
					.filter((procedure) => procedure.moduleName.toLowerCase() !== currentLower);
			}
			return {
				codeNames,
				moduleName: decoded.moduleName,
				moduleKind,
				projectProcedures,
				projectSymbols,
			};
		} catch {
			return {};
		}
	}

	private async _buildProjectProcedureContext(
		document: vscode.TextDocument,
	): Promise<VbaProcedureSignature[] | undefined> {
		if (document.uri.scheme !== XLIDE_SCHEME) {
			return undefined;
		}
		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			return entries
				? await this._loadCrossModuleProcedureSignatures(
					decoded.xlsmPath,
					entries,
					decoded.moduleName.toLowerCase(),
				)
				: undefined;
		} catch {
			return undefined;
		}
	}

	/** Maps a host module-type string to a {@link ModuleSymbolKind}. */
	private _moduleKind(type: string): ModuleSymbolKind {
		switch (type) {
			case 'class':
				return 'class';
			case 'document':
				return 'document';
			case 'userform':
				return 'userform';
			default:
				return 'standard';
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
			case 'host':
			case 'class':
			case 'document':
			case 'userform':
				return vscode.CompletionItemKind.Class;
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
			case 'runtime':
				return vscode.CompletionItemKind.Function;
			case 'constant':
				return vscode.CompletionItemKind.Constant;
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

	private _key(xlsmPath: string): string {
		return process.platform === 'win32' ? xlsmPath.toLowerCase() : xlsmPath;
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
	let lastCanonicalCandidate = canonicalCandidateFromEditor(vscode.window.activeTextEditor);
	let activeKeywordSnippet:
		| { editor: vscode.TextEditor; documentKey: string; textChangeSerialAtAccept: number }
		| undefined;
	let textChangeSerial = 0;
	const lastTextChange = new Map<string, { at: number; serial: number }>();
	const flushCanonicalCandidate = (): void => {
		const candidate = lastCanonicalCandidate;
		if (!candidate || !isVbaDocument(candidate.editor.document)) {
			return;
		}
		void provider.applyCanonicalCase(
			candidate.editor.document,
			candidate.position,
			candidate.editor,
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
		),
		vscode.workspace.onDidChangeTextDocument((event) => {
			markTextChange(event.document);
			if (!isVbaDocument(event.document) || event.contentChanges.length !== 1) {
				return;
			}
			const change = event.contentChanges[0];
			if (!change.range.isEmpty || !isCanonicalCaseBoundary(change.text)) {
				return;
			}
			void provider.applyCanonicalCase(event.document, change.range.start);
		}),
		vscode.window.onDidChangeTextEditorSelection((event) => {
			maybeLeaveKeywordSnippet(event);
			const previous = lastCanonicalCandidate;
			if (previous?.editor === event.textEditor) {
				void provider.applyCanonicalCase(
					previous.editor.document,
					previous.position,
					previous.editor,
				);
			}
			lastCanonicalCandidate = canonicalCandidateFromEditor(event.textEditor);
		}),
		vscode.window.onDidChangeActiveTextEditor((editor) => {
			if (activeKeywordSnippet && editor !== activeKeywordSnippet.editor) {
				activeKeywordSnippet = undefined;
			}
			flushCanonicalCandidate();
			lastCanonicalCandidate = canonicalCandidateFromEditor(editor);
		}),
		vscode.window.onDidChangeWindowState((state) => {
			if (!state.focused) {
				flushCanonicalCandidate();
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
	return document.languageId === 'vba' || document.uri.scheme === XLIDE_SCHEME;
}

function isCanonicalCaseBoundary(text: string): boolean {
	return (
		/^[ \t]$/.test(text) ||
		/^\r?\n[ \t]*$/.test(text) ||
		['(', ')', '.', ',', '=', ':', '+', '-', '*', '/', '\\', '&', '<', '>'].includes(text)
	);
}
