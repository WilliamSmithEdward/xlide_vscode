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
import {
	XLIDE_SCHEME,
	decodeModuleUri,
	isVbaDocument,
} from './xlideFileSystem';
import { leadingWhitespace } from './vbaSourceScan';
import { procedureHeaderParensEdit } from './vbaSmartEnter';
import { xlideEditorBlockLayoutFromConfig } from './globalSettings';
import {
	DocRegistry,
	EventHandlerCompletion,
	getHostType,
	HoverContext,
	IdentifierCompletion,
	KeywordCompletion,
	MemberCompletion,
	materializeKeywordSnippet,
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
	spaceTriggerMayComplete,
	SignatureHelpContext,
	TypeCompletion,
	type VbaProcedureLabelCompletion,
} from './analyzer';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import {
	VbaEditorProjectContextService,
	toEventHandlerCompletionContext,
	toIdentifierCompletionContext,
	toMemberCompletionContext,
	toTypeCompletionContext,
	type EditorProjectContext,
} from './vbaEditorProjectContext';
import {
	resolveVbaTestDirectiveCompletions,
	type VbaTestDirectiveCompletion,
} from './vbaTestDirectiveCompletion';
import { startPerformanceTrace } from './performanceTrace';

const KEYWORD_SNIPPET_ACCEPTED_COMMAND = 'xlide.vba.keywordSnippetAccepted';
const KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS = 150;
const MAX_PENDING_CANONICAL_CASE_REQUESTS = 16;
const CANONICAL_LINE_IDLE_DELAY_MS = 200;
const COMPLETION_PROJECT_CONTEXT_BUDGET_MS = 150;
const HOVER_PROJECT_CONTEXT_BUDGET_MS = 120;
const SIGNATURE_HELP_PROJECT_CONTEXT_BUDGET_MS = 150;

type CanonicalCaseRequest = {
	document: vscode.TextDocument;
	editorHint?: vscode.TextEditor;
	resolveEdits: (source: string, ctx: CanonicalCaseContext) => CanonicalCaseEdit[];
};

interface CanonicalLineOptions {
	completeProcedureHeader?: boolean;
}

class VbaMemberCompletionProvider
	implements
		vscode.CompletionItemProvider,
		vscode.HoverProvider,
		vscode.SignatureHelpProvider
{
	private readonly _pendingCanonicalCaseRequests: CanonicalCaseRequest[] = [];
	private _applyingCanonicalCase = false;

	constructor(
		private readonly _projectContext: VbaEditorProjectContextService,
		private readonly _docs?: DocRegistry,
	) {}

	/** Drop derived editor contexts for a workbook (e.g. after a project change). */
	invalidate(xlsmPath?: string): void {
		this._projectContext.invalidate(xlsmPath);
	}

	async provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
		context?: vscode.CompletionContext,
	): Promise<vscode.CompletionItem[]> {
		const trace = startPerformanceTrace('completion', document.uri.scheme);
		try {
			return await this._provideCompletionItems(document, position, token, context);
		} finally {
			trace.end(token?.isCancellationRequested ? 'canceled' : 'ok', document.uri.scheme);
		}
	}

	private async _provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
		context?: vscode.CompletionContext,
	): Promise<vscode.CompletionItem[]> {
		if (token?.isCancellationRequested) {
			return [];
		}
		const requestVersion = document.version;
		const directiveCompletions = this._testDirectiveCompletions(document, position);
		const directiveItems = directiveCompletions.map(
			(completion) => this._toTestDirectiveItem(completion, position.line),
		);
		if (directiveCompletions.some((completion) => completion.exclusive)) {
			return directiveItems;
		}

		// A space is typed far more often than it opens a grammar position
		// (after As/New, End/Exit, a member dot, ...); bail on ordinary code
		// before any full-source resolver runs.
		if (
			context?.triggerKind === vscode.CompletionTriggerKind.TriggerCharacter &&
			context.triggerCharacter === ' ' &&
			!this._spaceTriggerMayComplete(document, position)
		) {
			return directiveItems;
		}

		const source = document.getText();
		const offset = document.offsetAt(position);
		const range = this._completionRange(document, position);

		const quickTypes = resolveTypeCompletions(source, offset, {});
		if (quickTypes.length > 0) {
			return quickTypes.map((t) => this._toTypeItem(t, range));
		}

		const cachedProjectCtx = this._projectContext.cachedEditorProjectContext(document);
		const fastProjectCtx = cachedProjectCtx ?? this._projectContext.localEditorProjectContext(document, source);
		if (!cachedProjectCtx && document.uri.scheme === XLIDE_SCHEME) {
			this._projectContext.warmEditorProjectContext(document, source);
		}

		const fastTypes = resolveTypeCompletions(source, offset, toTypeCompletionContext(fastProjectCtx));
		if (fastTypes.length > 0) {
			return fastTypes.map((t) => this._toTypeItem(t, range));
		}

		const fastMembers = resolveMemberCompletions(source, offset, toMemberCompletionContext(fastProjectCtx));
		if (fastMembers.length > 0) {
			return fastMembers.map((mem) => this._toItem(mem, range, source, offset));
		}

		const fastEvents = resolveEventHandlerCompletions(source, offset, toEventHandlerCompletionContext(fastProjectCtx));
		if (fastEvents.length > 0) {
			return fastEvents.map((event) => this._toEventHandlerItem(event, range));
		}

		let projectCtx = cachedProjectCtx;
		if (!projectCtx) {
			projectCtx = await this._projectContext.buildEditorProjectContextWithin(
				document,
				source,
				COMPLETION_PROJECT_CONTEXT_BUDGET_MS,
			) ?? fastProjectCtx;
			if (token?.isCancellationRequested || document.version !== requestVersion) {
				return [];
			}
		}
		const typeCtx = toTypeCompletionContext(projectCtx);
		const types = resolveTypeCompletions(source, offset, typeCtx);
		if (types.length > 0) {
			return types.map((t) => this._toTypeItem(t, range));
		}

		const memberCtx = toMemberCompletionContext(projectCtx);
		const members = resolveMemberCompletions(source, offset, memberCtx);
		if (members.length > 0) {
			return members.map((mem) => this._toItem(mem, range, source, offset));
		}

		const eventCtx = toEventHandlerCompletionContext(projectCtx);
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

		const identCtx = toIdentifierCompletionContext(projectCtx);
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
			const projectCtx = this._projectContext.cachedEditorProjectContext(document) ?? {};
			const edits = resolveEdits(source, {
				member: toMemberCompletionContext(projectCtx),
				identifier: toIdentifierCompletionContext(projectCtx),
				type: toTypeCompletionContext(projectCtx),
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

	async provideHover(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
	): Promise<vscode.Hover | undefined> {
		const trace = startPerformanceTrace('hover', document.uri.scheme);
		try {
			return await this._provideHover(document, position, token);
		} finally {
			trace.end(token?.isCancellationRequested ? 'canceled' : 'ok', document.uri.scheme);
		}
	}

	private async _provideHover(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
	): Promise<vscode.Hover | undefined> {
		if (token?.isCancellationRequested) {
			return undefined;
		}
		const requestVersion = document.version;
		const source = document.getText();
		const offset = document.offsetAt(position);
		const cached = this._projectContext.cachedEditorProjectContext(document);
		const fastCtx = cached ?? this._projectContext.cheapEditorProjectContext(document);
		if (!cached && document.uri.scheme === XLIDE_SCHEME) {
			this._projectContext.warmEditorProjectContext(document, source);
		}
		let info = resolveHover(source, offset, this._hoverContext(fastCtx));
		if (!info && !cached) {
			info = resolveHover(source, offset, this._hoverContext(
				this._projectContext.localEditorProjectContext(document, source),
			));
		}
		if (!info && !cached && document.uri.scheme === XLIDE_SCHEME) {
			const projectCtx = await this._projectContext.buildEditorProjectContextWithin(
				document,
				source,
				HOVER_PROJECT_CONTEXT_BUDGET_MS,
			);
			if (token?.isCancellationRequested || document.version !== requestVersion) {
				return undefined;
			}
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
		token?: vscode.CancellationToken,
	): Promise<vscode.SignatureHelp | undefined> {
		const trace = startPerformanceTrace('signatureHelp', document.uri.scheme);
		try {
			return await this._provideSignatureHelp(document, position, token);
		} finally {
			trace.end(token?.isCancellationRequested ? 'canceled' : 'ok', document.uri.scheme);
		}
	}

	private async _provideSignatureHelp(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
	): Promise<vscode.SignatureHelp | undefined> {
		if (token?.isCancellationRequested) {
			return undefined;
		}
		const requestVersion = document.version;
		const source = document.getText();
		const offset = document.offsetAt(position);
		const cached = this._projectContext.cachedEditorProjectContext(document);
		const fastCtx = cached ?? this._projectContext.cheapEditorProjectContext(document);
		if (!cached && document.uri.scheme === XLIDE_SCHEME) {
			this._projectContext.warmEditorProjectContext(document, source);
		}
		let info = resolveSignatureHelp(source, offset, this._signatureHelpContext(fastCtx, source));
		if (!info && !cached) {
			info = resolveSignatureHelp(
				source,
				offset,
				this._signatureHelpContext(this._projectContext.localEditorProjectContext(document, source), source),
			);
		}
		if (!info && !cached && document.uri.scheme === XLIDE_SCHEME) {
			const projectCtx = await this._projectContext.buildEditorProjectContextWithin(
				document,
				source,
				SIGNATURE_HELP_PROJECT_CONTEXT_BUDGET_MS,
			);
			if (token?.isCancellationRequested || document.version !== requestVersion) {
				return undefined;
			}
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

	private _hoverContext(ctx: EditorProjectContext): HoverContext {
		return {
			...toMemberCompletionContext(ctx),
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
			...toMemberCompletionContext(ctx),
			moduleName: ctx.moduleName,
			moduleSource: source,
			projectProcedures: ctx.projectProcedures,
			docRegistry: this._docs,
		};
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

	private _spaceTriggerMayComplete(
		document: vscode.TextDocument,
		position: vscode.Position,
	): boolean {
		const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
		const continued = position.line > 0 &&
			/\s_[ \t]*$/.test(document.lineAt(position.line - 1).text);
		return spaceTriggerMayComplete(linePrefix, continued);
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
	projectIndexService: VbaProjectIndexService,
	selector: vscode.DocumentSelector,
	docs?: DocRegistry,
): void {
	const provider = new VbaMemberCompletionProvider(
		new VbaEditorProjectContextService(projectIndexService),
		docs,
	);
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
	const userTouchedCanonicalLines = new Set<string>();
	const editorHintFor = (document: vscode.TextDocument): vscode.TextEditor | undefined => {
		const active = vscode.window.activeTextEditor;
		return active?.document === document ? active : undefined;
	};
	const canonicalLineKey = (document: vscode.TextDocument, lineNumber: number): string =>
		`${document.uri.toString()}\n${lineNumber}`;
	const canonicalLineWasTouched = (document: vscode.TextDocument, lineNumber: number): boolean =>
		userTouchedCanonicalLines.has(canonicalLineKey(document, lineNumber));
	const applyCanonicalLine = (
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
		options: CanonicalLineOptions = {},
	): void => {
		if (!canonicalLineWasTouched(document, lineNumber)) {
			return;
		}
		void provider.applyCanonicalCaseForLine(document, lineNumber, editorHint, options);
	};
	const applyCanonicalPosition = (
		document: vscode.TextDocument,
		position: vscode.Position,
		editorHint?: vscode.TextEditor,
	): void => {
		if (!canonicalLineWasTouched(document, position.line)) {
			return;
		}
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
		for (const key of [...userTouchedCanonicalLines]) {
			if (key.startsWith(prefix)) {
				userTouchedCanonicalLines.delete(key);
			}
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
				const lineNumber = Math.min(change.range.start.line, Math.max(0, event.document.lineCount - 1));
				touchedLines.add(lineNumber);
				userTouchedCanonicalLines.add(canonicalLineKey(event.document, lineNumber));
				if (!change.range.isEmpty) {
					continue;
				}
				// Token-boundary characters (space, '(', '=', ...) no longer
				// resolve casing synchronously inside the change event; the
				// idle line pass below covers every touched line.
				if (canonicalCaseBoundaryKind(change.text) === 'line') {
					immediateLines.add(change.range.start.line);
					scheduleCanonicalLine(
						event.document,
						change.range.start.line,
						editorHint,
					);
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

