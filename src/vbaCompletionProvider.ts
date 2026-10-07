// Host-context completion provider for VBA documents.
//
// A CompletionItemProvider triggered on '.' that resolves the type of the
// receiver expression (ThisWorkbook, Application, ActiveSheet, a worksheet
// code name, Me, or a typed local/module variable) and offers the verified
// Excel object-model members for that type. The same provider also offers
// type-name completions in a declaration type position (after `As` / `As New`),
// event-handler stubs, procedure labels, keywords/snippets, and identifiers.
// VbaKeywordSnippetTracker owns the companion leave-detection state machine
// for accepted keyword snippets. See the Host-Context Member Completion
// addendum and Phases 6-7 in docs/xlide_vba_language_service_roadmap.md.
//
// Extracted verbatim from vbaMemberCompletion.ts (audit #27).

import * as vscode from 'vscode';
import { assignmentValueTriggerMayComplete, completionLineCursorContext } from './analyzer/completion/cursorContext';
import { assignmentTargetAt, resolveAssignmentValueCompletion } from './analyzer/completion/assignmentValueCompletion';
import { macroNameStringMayResolveAt } from './analyzer/completion/macroNames';
import { hasDocContent, renderDocMarkdown } from './analyzer/docs/docModel';
import { isVbaDocument } from './xlideFileSystem';
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
	resolveArgumentValueCompletion,
	resolveMacroNameCompletions,
	type MacroNameCandidate,
	resolveIdentifierCompletions,
	type ArgumentValueCompletion,
	type HostConstant,
	resolveKeywordCompletions,
	resolveMemberCompletions,
	memberCompletionStatus,
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
import { vbaColorNameShadowedAt } from './vbaColors';

export const KEYWORD_SNIPPET_ACCEPTED_COMMAND = 'xlide.vba.keywordSnippetAccepted';
const KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS = 150;

/**
 * A member's name as code has to write it. A name that is not an identifier,
 * such as the field `Unit Price` of an Access form's record source, is only
 * reachable in brackets: `Me.[Unit Price]` (issue #206).
 */
export function memberNameAsWritten(name: string): string {
	return /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.test(name) ? name : `[${name}]`;
}

/**
 * Tracks the keyword snippet most recently accepted from completion (via
 * KEYWORD_SNIPPET_ACCEPTED_COMMAND) and forces `leaveSnippet` when the user
 * navigates away from it by mouse or keyboard, so stale tab stops do not
 * capture Tab/Enter. Keyboard moves within the grace window of a text change
 * are treated as typing (snippet navigation), not leaving.
 */
export class VbaKeywordSnippetTracker {
	private _activeKeywordSnippet:
		| { editor: vscode.TextEditor; documentKey: string; textChangeSerialAtAccept: number }
		| undefined;
	private _textChangeSerial = 0;
	private readonly _lastTextChange = new Map<string, { at: number; serial: number }>();

	/** Command handler attached to accepted keyword-snippet completion items. */
	handleSnippetAccepted(): void {
		const editor = vscode.window.activeTextEditor;
		if (!editor || !isVbaDocument(editor.document)) {
			return;
		}
		this._activeKeywordSnippet = {
			editor,
			documentKey: editor.document.uri.toString(),
			textChangeSerialAtAccept: this._textChangeSerial,
		};
	}

	handleTextDocumentChange(event: vscode.TextDocumentChangeEvent): void {
		const document = event.document;
		if (isVbaDocument(document)) {
			this._textChangeSerial += 1;
			this._lastTextChange.set(document.uri.toString(), {
				at: Date.now(),
				serial: this._textChangeSerial,
			});
		}
	}

	/** Drop per-document state on close so _lastTextChange stays bounded. */
	handleDocumentClose(document: vscode.TextDocument): void {
		const key = document.uri.toString();
		this._lastTextChange.delete(key);
		if (this._activeKeywordSnippet?.documentKey === key) {
			this._activeKeywordSnippet = undefined;
		}
	}

	handleSelectionChange(event: vscode.TextEditorSelectionChangeEvent): void {
		if (!this._activeKeywordSnippet || event.textEditor !== this._activeKeywordSnippet.editor) {
			return;
		}
		if (!isVbaDocument(event.textEditor.document)) {
			this._activeKeywordSnippet = undefined;
			return;
		}
		if (
			event.kind !== vscode.TextEditorSelectionChangeKind.Mouse &&
			event.kind !== vscode.TextEditorSelectionChangeKind.Keyboard
		) {
			return;
		}
		if (event.kind === vscode.TextEditorSelectionChangeKind.Keyboard) {
			const changed = this._lastTextChange.get(this._activeKeywordSnippet.documentKey);
			if (
				changed &&
				changed.serial > this._activeKeywordSnippet.textChangeSerialAtAccept &&
				Date.now() - changed.at <= KEYBOARD_NAV_TEXT_CHANGE_GRACE_MS
			) {
				return;
			}
		}
		this._activeKeywordSnippet = undefined;
		void vscode.commands.executeCommand('leaveSnippet');
	}

	handleActiveEditorChange(editor: vscode.TextEditor | undefined): void {
		if (this._activeKeywordSnippet && editor !== this._activeKeywordSnippet.editor) {
			this._activeKeywordSnippet = undefined;
		}
	}
}

export class VbaMemberCompletionProvider implements vscode.CompletionItemProvider, vscode.Disposable {
	private _disposed = false;
	private _pendingRecovery: { document: vscode.TextDocument; stop: () => void } | undefined;
	constructor(
		private readonly _projectContext: VbaEditorProjectContextService,
	) {}

	dispose(): void {
		this._disposed = true;
		this._pendingRecovery?.stop();
	}

	/** Drop derived editor contexts for a project (e.g. after a project change). */
	invalidate(projectPath?: string): void {
		this._projectContext.invalidate(projectPath);
	}

	/** Reopen member suggestions when Backspace widens a prefix after a miss. */
	handleTextDocumentChange(event: vscode.TextDocumentChangeEvent): void {
		const document = event.document;
		if (event.contentChanges.length > 0 && this._pendingRecovery?.document === document) {
			this._pendingRecovery.stop();
		}
		const editor = vscode.window.activeTextEditor;
		if (this._disposed || !isVbaDocument(document) || !editor || editor.document !== document ||
			event.reason !== undefined || event.contentChanges.length !== 1 ||
			editor.selections.length !== 1) {
			return;
		}
		const change = event.contentChanges[0];
		if (change.text !== '' || change.rangeLength !== 1 ||
			change.range.start.line !== change.range.end.line ||
			change.range.end.character !== change.range.start.character + 1) {
			return;
		}
		const expectedCaret = change.range.start;
		const prefix = document.lineAt(expectedCaret.line).text.slice(0, expectedCaret.character);
		if (!/\.(?:[\p{L}\p{M}0-9_]*|\[[^\]\r\n]*\]?)$/u.test(prefix)) {
			return;
		}
		const version = document.version;
		// The document event can precede the updated selection and suggest
		// widget filtering. Wait for both, and discard superseded keystrokes.
		const isCurrent = (): boolean => {
			const caret = editor.selection.active;
			return !this._disposed && !document.isClosed && document.version === version &&
				vscode.window.activeTextEditor === editor && editor.document === document &&
				editor.selections.length === 1 && editor.selection.isEmpty &&
				caret.line === expectedCaret.line && caret.character === expectedCaret.character;
		};
		const recover = async (): Promise<void> => {
			if (!isCurrent()) { return; }
			const source = document.getText();
			const offset = document.offsetAt(expectedCaret);
			const cursor = completionLineCursorContext(source, offset);
			if (cursor.inComment || cursor.inString) { return; }
			const cachedProjectCtx = this._projectContext.cachedEditorProjectContext(document)
				?? this._projectContext.readyEditorProjectContext?.(document);
			const projectCtx = cachedProjectCtx ?? this._projectContext.cheapEditorProjectContext(document);
			// Probe host members without rebuilding module symbols or rendering
			// every completion's documentation just to test whether a match exists.
			let hasMatches = memberCompletionStatus(source, offset, toMemberCompletionContext(projectCtx));
			// Without project context, a named root can be extended or shadowed
			// by source declarations (Me, ThisWorkbook, or a class named Range).
			// Only a resolved call chain can rule out such missing named members.
			const knownCallResult = /\)\.(?:[\p{L}\p{M}0-9_]*|\[[^\]\r\n]*\]?)$/u.test(prefix);
			if (hasMatches === undefined || (hasMatches === false && !cachedProjectCtx && !knownCallResult)) {
				// Source-backed receivers need the full cross-module context, which
				// a keystroke's version bump can make temporarily unavailable.
				// Recovery has no partial widget to serve. Keep waiting asynchronously
				// for the full context rather than abandoning it at the request budget.
				const built = await this._projectContext.buildEditorProjectContext(document, source)
					.catch(() => undefined);
				if (!isCurrent()) { return; }
				hasMatches = Boolean(built && memberCompletionStatus(source, offset, toMemberCompletionContext(built)));
			}
			if (hasMatches && isCurrent()) {
				// Recovery is keyboard-driven, so a subsequent miss should dismiss
				// suggestions just as it does after typing the member dot.
				void vscode.commands.executeCommand('editor.action.triggerSuggest', { auto: true });
			}
		};
		// A busy host can deliver the native caret update after a wall-clock
		// timeout. Keep one pending recovery until that update or a later action;
		// unrelated dirty-state notifications do not supersede the deletion.
		this._pendingRecovery?.stop();
		let settled = false;
		const subscriptions: vscode.Disposable[] = [];
		let initialTry: ReturnType<typeof setTimeout> | undefined;
		const stopWaiting = (): void => {
			if (settled) { return; }
			settled = true;
			for (const subscription of subscriptions) { subscription.dispose(); }
			if (initialTry !== undefined) { clearTimeout(initialTry); }
			if (this._pendingRecovery?.stop === stopWaiting) { this._pendingRecovery = undefined; }
		};
		this._pendingRecovery = { document, stop: stopWaiting };
		const tryRecover = (): void => {
			if (settled) { return; }
			if (this._disposed || document.isClosed || document.version !== version || vscode.window.activeTextEditor !== editor) {
				stopWaiting();
				return;
			}
			if (!isCurrent()) { return; }
			stopWaiting();
			void recover();
		};
		subscriptions.push(
			vscode.window.onDidChangeTextEditorSelection(event => {
				if (event.textEditor !== editor) { return; }
				if (!isCurrent()) { stopWaiting(); return; }
				tryRecover();
			}),
			vscode.window.onDidChangeActiveTextEditor(active => {
				if (active !== editor) { stopWaiting(); }
			}),
			vscode.workspace.onDidCloseTextDocument(closed => {
				if (closed === document) { stopWaiting(); }
			}),
		);
		initialTry = setTimeout(tryRecover, 0);
	}

	async provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
		context?: vscode.CompletionContext,
	): Promise<vscode.CompletionList> {
		const trace = startPerformanceTrace('completion', document.uri.scheme);
		const requestVersion = document.version;
		try {
			const result = await this._provideCompletionItems(document, position, token, context);
			return token?.isCancellationRequested || document.isClosed || document.version !== requestVersion
				? new vscode.CompletionList([], false) : result;
		} finally {
			trace.end(token?.isCancellationRequested ? 'canceled' : 'ok', document.uri.scheme);
		}
	}

	private async _provideCompletionItems(
		document: vscode.TextDocument,
		position: vscode.Position,
		token?: vscode.CancellationToken,
		context?: vscode.CompletionContext,
	): Promise<vscode.CompletionList> {
		if (token?.isCancellationRequested || document.isClosed) {
			return new vscode.CompletionList([], false);
		}
		const directiveCompletions = this._testDirectiveCompletions(document, position);
		const directiveItems = directiveCompletions.map(
			(completion) => this._toTestDirectiveItem(completion, position.line),
		);
		if (directiveCompletions.some((completion) => completion.exclusive)) {
			return new vscode.CompletionList(directiveItems, false);
		}

		// A space is typed far more often than it opens a grammar position
		// (after As/New, End/Exit, a member dot, ...); bail on ordinary code
		// before any full-source resolver runs.
		if (
			context?.triggerKind === vscode.CompletionTriggerKind.TriggerCharacter &&
			context.triggerCharacter === ' ' &&
			!this._spaceTriggerMayComplete(document, position)
		) {
			return new vscode.CompletionList(directiveItems, false);
		}

		if (context?.triggerKind === vscode.CompletionTriggerKind.TriggerCharacter && context.triggerCharacter === '='
			&& !assignmentValueTriggerMayComplete(document.lineAt(position.line).text.slice(0, position.character),
				position.line > 0 && /_\s*$/.test(document.lineAt(position.line - 1).text))) {
			return new vscode.CompletionList([], false);
		}
		const source = document.getText();
		const offset = document.offsetAt(position);
		if (completionLineCursorContext(source, offset).inComment) {
			return new vscode.CompletionList(directiveItems, false);
		}
		if (completionLineCursorContext(source, offset).inString && !macroNameStringMayResolveAt(source, offset)) {
			return new vscode.CompletionList([], false);
		}
		const range = this._completionRange(document, position, source, offset);
		const bracketedMember = document.lineAt(range.start.line).text[range.start.character] === '['
			&& completionLineCursorContext(source, offset).significantTokens.at(-2)?.rawText === '.';
		let insertParens: boolean | undefined;
		const shouldInsertParens = (): boolean =>
			insertParens ??= !/^[ \t]*\(/.test(document.lineAt(range.end.line).text.slice(range.end.character))
				&& callableCompletionShouldInsertParens(source, offset);

		const cachedProjectCtx = this._projectContext.cachedEditorProjectContext(document)
			?? this._projectContext.readyEditorProjectContext?.(document);
		const bareIdentifierStatement = /^[ \t]*[\p{L}_][\p{L}\p{M}\p{N}_]*[$%&!#@^]?$/u.test(
			document.lineAt(position.line).text.slice(0, position.character),
		) && completionLineCursorContext(source, offset).statementStart === document.offsetAt(new vscode.Position(position.line, 0));
		const fastProjectCtx = cachedProjectCtx ?? this._projectContext.localEditorProjectContext(document, source, bareIdentifierStatement);
		if (!cachedProjectCtx) {
			this._projectContext.warmEditorProjectContext(document, source);
		}

		// While the cross-module project context is still loading, results are
		// served from the synchronous intra-module (local) context. They are marked
		// incomplete so VS Code keeps requesting (and refreshing the list) as the
		// context warms, instead of caching an early intra-only/empty result and
		// never asking again. Once the full context is available the list is
		// complete and VS Code filters it client-side.
		const contextComplete = Boolean(cachedProjectCtx);
		const list = (items: vscode.CompletionItem[]): vscode.CompletionList =>
			new vscode.CompletionList(items, !contextComplete);

		// A string that names a procedure: `Application.Run "`, ReDim's
		// `.OnClick "` (issue #217). Nothing else is offered inside it.
		const macroNames = resolveMacroNameCompletions(source, offset, {
			...toMemberCompletionContext(fastProjectCtx),
			moduleName: fastProjectCtx.moduleName,
			moduleSource: source,
			projectProcedures: fastProjectCtx.projectProcedures,
			macroProcedures: fastProjectCtx.macroProcedures,
		});
		if (macroNames) {
			const replace = new vscode.Range(document.positionAt(macroNames.contentSpan.start), document.positionAt(macroNames.contentSpan.end));
			return list(macroNames.candidates.map((candidate) => this._toMacroNameItem(candidate, replace)));
		}
		// Ordinary strings are never code completion positions, including
		// manual requests and typing within an already-open string.
		if (completionLineCursorContext(source, offset).inString) {
			return new vscode.CompletionList([], false);
		}
		// A quote opens or closes any other string, where nothing is offered.
		if (context?.triggerKind === vscode.CompletionTriggerKind.TriggerCharacter && context.triggerCharacter === '"') {
			return new vscode.CompletionList(directiveItems, false);
		}

		const fastTypes = resolveTypeCompletions(source, offset, toTypeCompletionContext(fastProjectCtx));
		if (fastTypes.length > 0) {
			return list(fastTypes.map((t) => this._toTypeItem(t, range)));
		}

		const fastMembers = resolveMemberCompletions(source, offset, toMemberCompletionContext(fastProjectCtx));
		if (fastMembers.length > 0) {
			return list(fastMembers.map((mem) => this._toItem(mem, range, shouldInsertParens, bracketedMember)));
		}

		// Unmatched bracketed names remain member positions, without bare globals.
		if (bracketedMember) { return list([]); }

		const fastEvents = resolveEventHandlerCompletions(source, offset, toEventHandlerCompletionContext(fastProjectCtx));
		if (fastEvents.length > 0) {
			return list(fastEvents.map((event) => this._toEventHandlerItem(event, range)));
		}

		// Menu updates never wait for project loading. Serve local facts now;
		// the incomplete list asks VS Code to requery as its background cache warms.
		const projectCtx = fastProjectCtx;
		const memberCtx = toMemberCompletionContext(projectCtx);

		const labels = resolveProcedureLabelCompletions(source, offset);
		if (labels.length > 0) {
			return list(labels.map((label) => this._toProcedureLabelItem(label, range)));
		}

		const keywords = resolveKeywordCompletions(source, offset, {
			blockLayout: xlideEditorBlockLayoutFromConfig(vscode.workspace.getConfiguration('xlide')).value,
		});
		if (keywords.exclusive) {
			return list([
				...directiveItems,
				...keywords.items.map((item) => this._toKeywordItem(item, range, document)),
			]);
		}

		const identCtx = toIdentifierCompletionContext(projectCtx);
		// An assignment or argument with a known enum type has a known set of
		// legal values, and they sort above the general list rather than
		// replacing it: `Type:=someVariable` is legal too.
		const argumentValues = resolveAssignmentValueCompletion(source, offset, {
			...memberCtx, moduleName: projectCtx.moduleName, moduleKind: projectCtx.moduleKind, projectSymbols: projectCtx.projectSymbols,
		}) ?? resolveArgumentValueCompletion(
			source,
			offset,
			{
				...memberCtx,
				moduleName: projectCtx.moduleName,
				moduleSource: source,
				projectProcedures: projectCtx.projectProcedures,
			},
		);
		if (!argumentValues && (context?.triggerCharacter === '=' ||
			(context?.triggerCharacter === ' ' && assignmentTargetAt(source, offset)))) { return list([]); }
		const idents = resolveIdentifierCompletions(source, offset, identCtx);
		if (argumentValues?.enumName === 'ColorConstants') {
		const identifiersByName = new Map((argumentValues ? idents : []).map(id => [id.name.toLowerCase(), id]));
		const prioritizedNames = new Set((argumentValues?.constants ?? []).map(c => c.name.toLowerCase()));
		const colorValues = argumentValues?.enumName === 'ColorConstants';
		const shadowedColors = new Set(colorValues ? argumentValues.constants
			.filter(c => vbaColorNameShadowedAt(source, c.name, offset)).map(c => c.name.toLowerCase()) : []);
		const rgbShadowed = colorValues && vbaColorNameShadowedAt(source, 'rgb', offset);
		const rgbInsertion = rgbShadowed
			? (vbaColorNameShadowedAt(source, 'vba', offset) ? '&H000000&' : 'VBA.RGB(0, 0, 0)') : 'RGB(0, 0, 0)';
		return list([
			...directiveItems,
			...(colorValues ? [this._toRgbItem(range, rgbInsertion)] : []),
			...(argumentValues?.constants ?? []).map(
				(constant) => {
					const item = this._toArgumentValueItem(constant, argumentValues!, range);
					const binding = identifiersByName.get(constant.name.toLowerCase());
					if ((binding && binding.kind !== 'constant' && binding.kind !== 'enumMember')
						|| shadowedColors.has(constant.name.toLowerCase())) {
						item.insertText = colorValues && vbaColorNameShadowedAt(source, 'colorconstants', offset)
							? String(constant.value) : `${argumentValues!.enumName}.${constant.name}`;
					}
					return item;
				},
			),
			...idents.filter(id => !(colorValues && id.name.toLowerCase() === 'rgb' && !rgbShadowed)
				&& (!prioritizedNames.has(id.name.toLowerCase()) || shadowedColors.has(id.name.toLowerCase())
					|| (id.kind !== 'constant' && id.kind !== 'enumMember')))
				.map((id) => this._toIdentItem(id, range, shouldInsertParens)),
			...keywords.items.map((item) => this._toKeywordItem(item, range, document)),
		]);
		}
		const preferred = new Set((argumentValues?.constants ?? []).map(c => c.name.toLowerCase()));
		const identifiers = new Map(idents.map(id => [id.name.toLowerCase(), id]));
		let unfilteredIdentifiers: Map<string, IdentifierCompletion> | undefined;
		const qualifierIsShadowed = (name: string) => {
			unfilteredIdentifiers ??= new Map(resolveIdentifierCompletions(source, document.offsetAt(range.start), identCtx)
				.map(id => [id.name.toLowerCase(), id]));
			const binding = unfilteredIdentifiers.get(name.toLowerCase());
			return Boolean(binding && !(argumentValues?.origin === 'source' && (binding.kind === 'module'
				|| binding.kind === 'enum' && binding.enumOwner?.toLowerCase() === argumentValues.qualifiedEnumName?.toLowerCase())));
		};
		const matchesEnum = (id: IdentifierCompletion) => argumentValues?.origin !== 'host' && argumentValues?.origin !== 'runtime' && id.detail === `${argumentValues?.enumName} member`
			&& id.enumOwner?.toLowerCase() === argumentValues?.qualifiedEnumName?.toLowerCase()
			|| (id.kind === 'constant' && /^(VBA|.+\/Office) constant As /.test(id.detail) && id.detail.endsWith(` As ${argumentValues?.enumName}`))
			|| id.detail === 'Boolean literal';
		return list([
			...directiveItems,
			...(argumentValues?.constants ?? []).flatMap(
				(constant) => {
					const item = this._toArgumentValueItem(constant, argumentValues!, range);
					const binding = identifiers.get(constant.name.toLowerCase());
					if (binding && !matchesEnum(binding)) {
						const qualified = argumentValues!.qualifiedEnumName ?? argumentValues!.enumName;
						const qualifier = qualified.includes('.') && qualifierIsShadowed(qualified.split('.')[0])
							? argumentValues!.enumName : qualified;
						if (qualifier !== qualified && qualifierIsShadowed(qualifier)) { return []; }
						item.insertText = `${qualifier}.${constant.name}`;
					}
					return [item];
				},
			),
			...idents.filter(id => !preferred.has(id.name.toLowerCase()) || !matchesEnum(id))
				.map((id) => this._toIdentItem(id, range, shouldInsertParens)),
			...keywords.items.map((item) => this._toKeywordItem(item, range, document)),
		]);
	}

	private _toRgbItem(range: vscode.Range, insertText: string): vscode.CompletionItem {
		const item = new vscode.CompletionItem(insertText === 'RGB(0, 0, 0)' ? 'RGB' : 'RGB color', vscode.CompletionItemKind.Color);
		item.filterText = 'RGB';
		item.detail = 'Insert a color; hover the swatch to open the color picker';
		item.insertText = insertText;
		item.range = range;
		item.sortText = '0:RGB';
		return item;
	}

	/**
	 * A constant offered as the value of an argument whose parameter declares
	 * that enumeration. Sorted ahead of everything else, because at `Type:=`
	 * these are the only values the parameter accepts.
	 */
	private _toArgumentValueItem(
		constant: HostConstant,
		accepted: ArgumentValueCompletion,
		range: vscode.Range,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(constant.name, vscode.CompletionItemKind.EnumMember);
		item.detail = constant.value === undefined
			? accepted.enumName
			: `${accepted.enumName} = ${constant.value}`;
		if (hasDocContent(constant.doc)) {
			item.documentation = new vscode.MarkdownString(renderDocMarkdown(constant.doc));
		}
		item.range = range;
		item.sortText = `0:${constant.name}`;
		return item;
	}

	/** A procedure offered inside a string that names one, `Module.Proc` (issue #217). */
	private _toMacroNameItem(candidate: MacroNameCandidate, range: vscode.Range): vscode.CompletionItem {
		const procedure = candidate.procedure;
		const item = new vscode.CompletionItem(candidate.name, procedure.kind === 'function' ? vscode.CompletionItemKind.Function : vscode.CompletionItemKind.Method);
		item.detail = procedure.signature ?? `${procedure.kind === 'function' ? 'Function' : 'Sub'} ${procedure.name}`;
		if (hasDocContent(procedure.doc)) {
			item.documentation = new vscode.MarkdownString(renderDocMarkdown(procedure.doc));
		}
		item.range = range;
		return item;
	}

	private _toItem(
		mem: MemberCompletion,
		range: vscode.Range,
		shouldInsertParens: () => boolean,
		bracketedMember: boolean,
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
		const writtenName = bracketedMember ? `[${mem.name}]` : memberNameAsWritten(mem.name);
		this._applyCompletionInsert(
			item,
			writtenName,
			range,
			mem.kind === 'method',
			mem.kind === 'method' && shouldInsertParens(),
		);
		item.filterText = bracketedMember ? writtenName : mem.name;
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
		shouldInsertParens: () => boolean,
	): vscode.CompletionItem {
		const item = new vscode.CompletionItem(id.name, this._identItemKind(id));
		item.detail = id.detail;
		if (id.documentation) {
			item.documentation = new vscode.MarkdownString(id.documentation);
		}
		const callable = id.callable ?? (id.kind === 'runtime' || id.kind === 'procedure');
		this._applyCompletionInsert(
			item,
			id.name,
			range,
			callable,
			callable && shouldInsertParens(),
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
		source: string,
		offset: number,
	): vscode.Range {
		const line = document.lineAt(position.line).text;
		const tokens = completionLineCursorContext(source, offset).significantTokens;
		const last = tokens[tokens.length - 1];
		if (last?.kind === 'bracketedIdentifier' && last.end === offset) {
			const start = position.character - (offset - last.start);
			const close = line.indexOf(']', start + 1);
			return new vscode.Range(position.line, start, position.line, close < 0 ? line.length : close + 1);
		}
		let start = position.character;
		if (line[start - 1] === '$' && /[\p{L}\p{M}0-9_]/u.test(line[start - 2] ?? '')) {
			start -= 1;
		}
		while (start > 0 && /[\p{L}\p{M}0-9_]/u.test(line[start - 1])) {
			start -= 1;
		}
		if (start === position.character && start > 0 && line[start - 1] === '#') {
			start -= 1;
		}
		let end = position.character;
		while (end < line.length && /[\p{L}\p{M}0-9_]/u.test(line[end])) {
			end += 1;
		}
		if (end > start && line[end] === '$') { end += 1; }
		return new vscode.Range(position.line, start, position.line, end);
	}

	private _lineIndent(document: vscode.TextDocument, line: number): string {
		return leadingWhitespace(document.lineAt(line).text);
	}

	private _identItemKind(id: IdentifierCompletion): vscode.CompletionItemKind {
		if (id.callable === false && id.kind === 'runtime') { return vscode.CompletionItemKind.Variable; }
		if (id.callable === true && id.kind === 'global') { return vscode.CompletionItemKind.Function; }
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
