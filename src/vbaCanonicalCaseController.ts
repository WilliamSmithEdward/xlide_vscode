// Canonical-case controller for VBA editing: applies VBE-style canonical
// casing edits (with a bounded retry queue) and owns the touch/idle line
// tracking state machine that decides when a line is recased.
//
// Extracted verbatim from vbaMemberCompletion.ts (audit #27): the apply half
// came from VbaMemberCompletionProvider, the tracking half from the
// registerVbaMemberCompletion closure.

import * as vscode from 'vscode';
import { isVbaDocument } from './xlideFileSystem';
import { procedureHeaderParensEdit } from './vbaSmartEnter';
import {
	type CanonicalCaseContext,
	type CanonicalCaseEdit,
	canonicalCaseBoundaryKind,
	resolveCanonicalCaseEdit,
	resolveCanonicalCaseEdits,
} from './analyzer';
import {
	VbaEditorProjectContextService,
	toIdentifierCompletionContext,
	toMemberCompletionContext,
	toTypeCompletionContext,
} from './vbaEditorProjectContext';

const MAX_PENDING_CANONICAL_CASE_REQUESTS = 16;
const CANONICAL_LINE_IDLE_DELAY_MS = 200;

type CanonicalCaseRequest = {
	document: vscode.TextDocument;
	editorHint?: vscode.TextEditor;
	resolveEdits: (source: string, ctx: CanonicalCaseContext) => CanonicalCaseEdit[];
};

interface CanonicalLineOptions {
	completeProcedureHeader?: boolean;
}

function canonicalCandidateFromEditor(
	editor: vscode.TextEditor | undefined,
): { editor: vscode.TextEditor; position: vscode.Position } | undefined {
	if (!editor || !isVbaDocument(editor.document)) {
		return undefined;
	}
	return { editor, position: editor.selection.active };
}

export class VbaCanonicalCaseController {
	private readonly _pendingCanonicalCaseRequests: CanonicalCaseRequest[] = [];
	private _applyingCanonicalCase = false;
	private _lastCanonicalCandidate = canonicalCandidateFromEditor(vscode.window.activeTextEditor);
	private readonly _canonicalLineTimers = new Map<string, ReturnType<typeof setTimeout>>();
	private readonly _userTouchedCanonicalLines = new Set<string>();
	// Touched lines whose change came while the document matched its file:
	// typing, or a reload. The document turning dirty says typing; a reload
	// never turns it dirty.
	private readonly _unconfirmedCanonicalLines = new Set<string>();

	constructor(
		private readonly _projectContext: VbaEditorProjectContextService,
	) {}

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
			// A pass that runs from a timer holds the editor it started with,
			// which may have closed since: then any editor still showing the
			// document takes the edit, and with none there is nothing to do.
			const visible = vscode.window.visibleTextEditors;
			const editor = editorHint?.document === document && visible.includes(editorHint)
				? editorHint
				: visible.find((candidate) => candidate.document === document);
			if (!editor) {
				return;
			}
			const source = document.getText();
			const edits = resolveEdits(source, this._canonicalCaseContext(document)).filter((edit) => {
				const range = new vscode.Range(
					document.positionAt(edit.start),
					document.positionAt(edit.end),
				);
				return document.getText(range) !== edit.text;
			});
			if (edits.length === 0) {
				return;
			}
			try {
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
			} catch (err) {
				// The editor can still close between the check above and the
				// edit reaching it. That leaves nothing to recase; any other
				// failure is a real one.
				if (vscode.window.visibleTextEditors.includes(editor)) {
					throw err;
				}
			}
		} finally {
			this._applyingCanonicalCase = false;
			const next = this._pendingCanonicalCaseRequests.shift();
			if (next) {
				void this._applyCanonicalCaseEdits(next.document, next.editorHint, next.resolveEdits);
			}
		}
	}

	private _canonicalCaseContext(document: vscode.TextDocument): CanonicalCaseContext {
		const projectCtx = this._projectContext.cachedEditorProjectContext(document) ?? {};
		return {
			member: toMemberCompletionContext(projectCtx),
			identifier: toIdentifierCompletionContext(projectCtx),
			type: toTypeCompletionContext(projectCtx),
		};
	}

	/**
	 * The casing still owed to lines the user typed, as edits for the save
	 * about to write the document. A save that comes before the pause - an
	 * auto-save with a short delay, a quick Ctrl+S - wrote the text uncased,
	 * and the pause then found a document matching its file and left it; now
	 * the save writes the recased text, and no recase is left to make the
	 * document dirty again.
	 */
	pendingEditsForSave(document: vscode.TextDocument): vscode.TextEdit[] {
		const prefix = `${document.uri.toString()}\n`;
		const lines: number[] = [];
		for (const key of [...this._userTouchedCanonicalLines]) {
			if (!key.startsWith(prefix)) {
				continue;
			}
			this._userTouchedCanonicalLines.delete(key);
			const timer = this._canonicalLineTimers.get(key);
			if (timer) {
				clearTimeout(timer);
				this._canonicalLineTimers.delete(key);
			}
			if (!this._unconfirmedCanonicalLines.delete(key)) {
				lines.push(Number(key.slice(prefix.length)));
			}
		}
		const source = document.getText();
		const ctx = this._canonicalCaseContext(document);
		const edits: vscode.TextEdit[] = [];
		for (const lineNumber of lines) {
			if (lineNumber >= document.lineCount) {
				continue;
			}
			const line = document.lineAt(lineNumber);
			const span = { start: document.offsetAt(line.range.start), end: document.offsetAt(line.range.end) };
			for (const edit of resolveCanonicalCaseEdits(source, span, ctx)) {
				const range = new vscode.Range(document.positionAt(edit.start), document.positionAt(edit.end));
				if (document.getText(range) !== edit.text) {
					edits.push(vscode.TextEdit.replace(range, edit.text));
				}
			}
		}
		return edits;
	}

	private _enqueueCanonicalCaseRequest(request: CanonicalCaseRequest): void {
		this._pendingCanonicalCaseRequests.push(request);
		const overflow = this._pendingCanonicalCaseRequests.length - MAX_PENDING_CANONICAL_CASE_REQUESTS;
		if (overflow > 0) {
			this._pendingCanonicalCaseRequests.splice(0, overflow);
		}
	}

	// -----------------------------------------------------------------------
	// Touch/idle line tracking
	// -----------------------------------------------------------------------

	handleTextDocumentChange(event: vscode.TextDocumentChangeEvent): void {
		if (!isVbaDocument(event.document)) {
			return;
		}
		// A reload - an agent's write to an open module, a restore from git, a
		// revert - arrives as a content change too, and leaves the document
		// matching its file, which typing never does. VS Code reports the
		// document still clean at a keystroke's change and turns it dirty just
		// after, in an event of its own; so a change made while it is clean is
		// held until that confirms it as typing.
		if (event.document.isDirty) {
			this._confirmTypedLines(event.document);
		} else if (event.contentChanges.length === 0) {
			// Saved or reverted: it matches its file again, and nothing typed
			// before is left to recase - a save took its casing with it.
			this._forgetTouchedLines(event.document);
		}
		const editorHint = this._editorHintFor(event.document);
		const touchedLines = new Set<number>();
		const immediateLines = new Set<number>();
		const reloadPossible = !event.document.isDirty;
		for (const change of event.contentChanges) {
			const lineNumber = Math.min(change.range.start.line, Math.max(0, event.document.lineCount - 1));
			touchedLines.add(lineNumber);
			const key = this._canonicalLineKey(event.document, lineNumber);
			this._userTouchedCanonicalLines.add(key);
			if (reloadPossible) {
				this._unconfirmedCanonicalLines.add(key);
			}
			if (!change.range.isEmpty || reloadPossible) {
				continue;
			}
			// Token-boundary characters (space, '(', '=', ...) no longer
			// resolve casing synchronously inside the change event; the
			// idle line pass below covers every touched line.
			if (canonicalCaseBoundaryKind(change.text) === 'line') {
				immediateLines.add(change.range.start.line);
				this._scheduleCanonicalLine(
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
			this._scheduleCanonicalLineIdle(event.document, lineNumber, editorHint);
		}
	}

	handleSelectionChange(event: vscode.TextEditorSelectionChangeEvent): void {
		const previous = this._lastCanonicalCandidate;
		if (previous && previous.editor !== event.textEditor) {
			this._applyCanonicalLine(
				previous.editor.document,
				previous.position.line,
				previous.editor,
				{ completeProcedureHeader: true },
			);
		} else if (previous?.editor === event.textEditor) {
			const nextPosition = event.textEditor.selection.active;
			if (previous.position.line !== nextPosition.line) {
				this._applyCanonicalLine(
					previous.editor.document,
					previous.position.line,
					previous.editor,
					{ completeProcedureHeader: true },
				);
			} else {
				this._applyCanonicalPosition(
					previous.editor.document,
					previous.position,
					previous.editor,
				);
			}
		}
		this._lastCanonicalCandidate = canonicalCandidateFromEditor(event.textEditor);
	}

	handleActiveEditorChange(editor: vscode.TextEditor | undefined): void {
		this._flushCanonicalLine();
		this._lastCanonicalCandidate = canonicalCandidateFromEditor(editor);
	}

	handleWindowStateChange(state: vscode.WindowState): void {
		if (!state.focused) {
			this._flushCanonicalLine();
		}
	}

	handleDocumentClose(document: vscode.TextDocument): void {
		const prefix = `${document.uri.toString()}\n`;
		for (const [key, timer] of this._canonicalLineTimers) {
			if (!key.startsWith(prefix)) {
				continue;
			}
			clearTimeout(timer);
			this._canonicalLineTimers.delete(key);
		}
		this._forgetTouchedLines(document);
	}

	private _forgetTouchedLines(document: vscode.TextDocument): void {
		const prefix = `${document.uri.toString()}\n`;
		for (const key of [...this._userTouchedCanonicalLines]) {
			if (key.startsWith(prefix)) {
				this._userTouchedCanonicalLines.delete(key);
				this._unconfirmedCanonicalLines.delete(key);
			}
		}
	}

	/** The document turned dirty: the changes it had while clean were typing. */
	private _confirmTypedLines(document: vscode.TextDocument): void {
		const prefix = `${document.uri.toString()}\n`;
		for (const key of [...this._unconfirmedCanonicalLines]) {
			if (key.startsWith(prefix)) {
				this._unconfirmedCanonicalLines.delete(key);
			}
		}
	}

	/**
	 * Whether a reload is all that touched the line: a change while the
	 * document matched its file that never turned it dirty. The line is
	 * forgotten, so recasing it cannot make the document dirty with edits
	 * nobody made.
	 */
	private _onlyReloaded(document: vscode.TextDocument, lineNumber: number): boolean {
		const key = this._canonicalLineKey(document, lineNumber);
		if (!this._unconfirmedCanonicalLines.delete(key)) {
			return false;
		}
		this._userTouchedCanonicalLines.delete(key);
		return true;
	}

	private _editorHintFor(document: vscode.TextDocument): vscode.TextEditor | undefined {
		const active = vscode.window.activeTextEditor;
		return active?.document === document ? active : undefined;
	}

	private _canonicalLineKey(document: vscode.TextDocument, lineNumber: number): string {
		return `${document.uri.toString()}\n${lineNumber}`;
	}

	private _canonicalLineWasTouched(document: vscode.TextDocument, lineNumber: number): boolean {
		return this._userTouchedCanonicalLines.has(this._canonicalLineKey(document, lineNumber));
	}

	private _applyCanonicalLine(
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
		options: CanonicalLineOptions = {},
	): void {
		if (!this._canonicalLineWasTouched(document, lineNumber) || this._onlyReloaded(document, lineNumber)) {
			return;
		}
		void this.applyCanonicalCaseForLine(document, lineNumber, editorHint, options);
	}

	private _applyCanonicalPosition(
		document: vscode.TextDocument,
		position: vscode.Position,
		editorHint?: vscode.TextEditor,
	): void {
		if (!this._canonicalLineWasTouched(document, position.line) || this._onlyReloaded(document, position.line)) {
			return;
		}
		void this.applyCanonicalCase(document, position, editorHint);
	}

	private _scheduleCanonicalLine(
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
		options: CanonicalLineOptions = {},
		delayMs = 0,
	): void {
		const key = this._canonicalLineKey(document, lineNumber);
		const existing = this._canonicalLineTimers.get(key);
		if (existing) {
			clearTimeout(existing);
		}
		const timer = setTimeout(() => {
			this._canonicalLineTimers.delete(key);
			if (!isVbaDocument(document)) {
				return;
			}
			this._applyCanonicalLine(document, lineNumber, editorHint, options);
		}, delayMs);
		this._canonicalLineTimers.set(key, timer);
	}

	private _scheduleCanonicalLineIdle(
		document: vscode.TextDocument,
		lineNumber: number,
		editorHint?: vscode.TextEditor,
	): void {
		this._scheduleCanonicalLine(document, lineNumber, editorHint, {}, CANONICAL_LINE_IDLE_DELAY_MS);
	}

	private _flushCanonicalLine(): void {
		const candidate = this._lastCanonicalCandidate;
		if (!candidate || !isVbaDocument(candidate.editor.document)) {
			return;
		}
		this._applyCanonicalLine(
			candidate.editor.document,
			candidate.position.line,
			candidate.editor,
			{ completeProcedureHeader: true },
		);
	}
}
