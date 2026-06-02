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
	decodeModuleUri,
	workbookIdentityKey,
} from './xlideFileSystem';
import { openModuleSourceForWorkbook } from './vbaOpenDocuments';
import { leadingWhitespace, normalizeSmartBlockLayout } from './vbaStructuralAnalysis';
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
	resolveSignatureHelp,
	resolveTypeCompletions,
	SignatureHelpContext,
	TypeCompletion,
	TypeCompletionContext,
} from './analyzer';
import {
	buildLiveVbaProjectIndex,
	buildVbaProjectIndex,
	moduleKindFromType,
	projectEditorSymbolContextForModule,
	type VbaProjectLiveOverride,
	type VbaProjectModuleInput,
} from './vbaProjectAnalysis';

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

class VbaMemberCompletionProvider
	implements
		vscode.CompletionItemProvider,
		vscode.HoverProvider,
		vscode.SignatureHelpProvider
{
	private readonly _cache = new Map<string, CachedModules>();
	private _applyingCanonicalCase = false;

	constructor(
		private readonly _bridge: PythonBridge,
		private readonly _docs?: DocRegistry,
	) {}

	/** Drop cached module lists for a workbook (e.g. after a project change). */
	invalidate(xlsmPath?: string): void {
		if (xlsmPath === undefined) {
			this._cache.clear();
		} else {
			const key = workbookIdentityKey(xlsmPath);
			this._cache.delete(key);
		}
	}

	async provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
	): Promise<vscode.CompletionItem[]> {
		const source = document.getText();
		const offset = document.offsetAt(position);
		const range = this._completionRange(document, position);
		const projectCtx = await this._buildEditorProjectContext(document, source);

		const memberCtx = this._memberContext(projectCtx);
		const members = resolveMemberCompletions(source, offset, memberCtx);
		if (members.length > 0) {
			return members.map((mem) => this._toItem(mem, range, source, offset));
		}

		const typeCtx = this._typeContext(projectCtx);
		const types = resolveTypeCompletions(source, offset, typeCtx);
		if (types.length > 0) {
			return types.map((t) => this._toTypeItem(t, range));
		}

		const eventCtx = this._eventHandlerContext(projectCtx);
		const events = resolveEventHandlerCompletions(source, offset, eventCtx);
		if (events.length > 0) {
			return events.map((event) => this._toEventHandlerItem(event, range));
		}

		const keywords = resolveKeywordCompletions(source, offset, {
			blockLayout: normalizeSmartBlockLayout(
				vscode.workspace.getConfiguration('xlide').get<string>('editor.blockLayout'),
			),
		});
		if (keywords.exclusive) {
			return keywords.items.map((item) => this._toKeywordItem(item, range, document));
		}

		const identCtx = this._identifierContext(projectCtx);
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
			return resolveCanonicalCaseEdits(source, { start, end }, ctx);
		});
	}

	private async _applyCanonicalCaseEdits(
		document: vscode.TextDocument,
		editorHint: vscode.TextEditor | undefined,
		resolveEdits: (source: string, ctx: CanonicalCaseContext) => CanonicalCaseEdit[],
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
			const projectCtx = await this._buildEditorProjectContext(document, source);
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
			const selections = editor.selections.map((selection) =>
				new vscode.Selection(selection.anchor, selection.active),
			);
			const restoreSelection = vscode.window.activeTextEditor === editor;
			const applied = await editor.edit((builder) => {
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
		const projectCtx = await this._buildEditorProjectContext(document, source);
		const info = resolveHover(source, offset, this._hoverContext(projectCtx));
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
		const projectCtx = await this._buildEditorProjectContext(document, source);
		const info = resolveSignatureHelp(source, offset, this._signatureHelpContext(projectCtx, source));
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
		if (document.uri.scheme !== XLIDE_SCHEME) {
			try {
				const project = buildLiveVbaProjectIndex(
					[{ moduleName: 'Module', moduleKind: 'standard', source }],
				);
				const context = projectEditorSymbolContextForModule(project, 'Module');
				return {
					moduleName: 'Module',
					moduleKind: 'standard',
					projectTypes: context.analysisOptions.projectTypes,
					projectClassMembers: context.analysisOptions.projectClassMembers,
					projectProcedures: context.externalProjectProcedures,
					projectSymbols: context.externalProjectSymbols,
				};
			} catch {
				return {};
			}
		}

		try {
			const decoded = decodeModuleUri(document.uri);
			const entries = await this._loadModules(decoded.xlsmPath);
			const allEntries = entries ?? [];
			const current = allEntries.find(
				(entry) => entry.name.toLowerCase() === decoded.moduleName.toLowerCase(),
			);
			const moduleKind = moduleKindFromType(current?.type);
			const project = await this._buildProjectIndexFromEntries(
				decoded.xlsmPath,
				allEntries,
				{
					liveOverride: {
						moduleName: decoded.moduleName,
						moduleKind,
						source,
					},
				},
			);
			const context = projectEditorSymbolContextForModule(project, decoded.moduleName);
			return {
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
			};
		} catch {
			return {};
		}
	}

	private async _loadModules(xlsmPath: string): Promise<ModuleEntry[] | undefined> {
		const key = workbookIdentityKey(xlsmPath);
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

	private async _moduleSource(
		xlsmPath: string,
		entry: ModuleEntry,
	): Promise<string | undefined> {
		const open = openModuleSourceForWorkbook(xlsmPath, entry.name);
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
	const editorHintFor = (document: vscode.TextDocument): vscode.TextEditor | undefined => {
		const active = vscode.window.activeTextEditor;
		return active?.document === document ? active : undefined;
	};
	const applyCanonicalLine = (
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
	): void => {
		void provider.applyCanonicalCaseForLine(document, lineNumber, editorHint);
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
	): void => {
		setTimeout(() => {
			if (!isVbaDocument(document)) {
				return;
			}
			applyCanonicalLine(document, lineNumber, editorHint);
		}, 0);
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
			if (!change.range.isEmpty) {
				return;
			}
			const boundary = canonicalCaseBoundaryKind(change.text);
			if (!boundary) {
				return;
			}
			const editorHint = editorHintFor(event.document);
			if (boundary === 'line') {
				scheduleCanonicalLine(event.document, change.range.start.line, editorHint);
			} else {
				applyCanonicalPosition(event.document, change.range.start, editorHint);
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
				);
			} else if (previous?.editor === event.textEditor) {
				const nextPosition = event.textEditor.selection.active;
				if (previous.position.line !== nextPosition.line) {
					applyCanonicalLine(
						previous.editor.document,
						previous.position.line,
						previous.editor,
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
