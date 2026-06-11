// Host-context completion provider for VBA documents.
//
// A CompletionItemProvider triggered on '.' that resolves the type of the
// receiver expression (ThisWorkbook, Application, ActiveSheet, a worksheet
// code name, Me, or a typed local/module variable) and offers the verified
// Excel object-model members for that type. The same provider also offers
// type-name completions in a declaration type position (after `As` / `As New`),
// event-handler stubs, procedure labels, keywords/snippets, and identifiers.
// See the Host-Context Member Completion addendum and Phases 6-7 in
// docs/xlide_vba_language_service_roadmap.md.
//
// Extracted verbatim from vbaMemberCompletion.ts (audit #27).

import * as vscode from 'vscode';
import { XLIDE_SCHEME } from './xlideFileSystem';
import { leadingWhitespace } from './vbaSourceScan';
import { xlideEditorBlockLayoutFromConfig } from './globalSettings';
import {
	EventHandlerCompletion,
	getHostType,
	IdentifierCompletion,
	KeywordCompletion,
	MemberCompletion,
	materializeKeywordSnippet,
	callableCompletionShouldInsertParens,
	resolveEventHandlerCompletions,
	resolveIdentifierCompletions,
	resolveKeywordCompletions,
	resolveMemberCompletions,
	resolveProcedureLabelCompletions,
	resolveTypeCompletions,
	spaceTriggerMayComplete,
	TypeCompletion,
	type VbaProcedureLabelCompletion,
} from './analyzer';
import {
	VbaEditorProjectContextService,
	toEventHandlerCompletionContext,
	toIdentifierCompletionContext,
	toMemberCompletionContext,
	toTypeCompletionContext,
} from './vbaEditorProjectContext';
import {
	resolveVbaTestDirectiveCompletions,
	type VbaTestDirectiveCompletion,
} from './vbaTestDirectiveCompletion';
import { startPerformanceTrace } from './performanceTrace';

export const KEYWORD_SNIPPET_ACCEPTED_COMMAND = 'xlide.vba.keywordSnippetAccepted';
const COMPLETION_PROJECT_CONTEXT_BUDGET_MS = 150;

export class VbaMemberCompletionProvider implements vscode.CompletionItemProvider {
	constructor(
		private readonly _projectContext: VbaEditorProjectContextService,
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
