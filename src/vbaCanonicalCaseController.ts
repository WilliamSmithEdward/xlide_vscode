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

	// -----------------------------------------------------------------------
	// Touch/idle line tracking
	// -----------------------------------------------------------------------

	handleTextDocumentChange(event: vscode.TextDocumentChangeEvent): void {
		if (!isVbaDocument(event.document)) {
			return;
		}
		const editorHint = this._editorHintFor(event.document);
		const touchedLines = new Set<number>();
		const immediateLines = new Set<number>();
		for (const change of event.contentChanges) {
			const lineNumber = Math.min(change.range.start.line, Math.max(0, event.document.lineCount - 1));
			touchedLines.add(lineNumber);
			this._userTouchedCanonicalLines.add(this._canonicalLineKey(event.document, lineNumber));
			if (!change.range.isEmpty) {
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
		for (const key of [...this._userTouchedCanonicalLines]) {
			if (key.startsWith(prefix)) {
				this._userTouchedCanonicalLines.delete(key);
			}
		}
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
		if (!this._canonicalLineWasTouched(document, lineNumber)) {
			return;
		}
		void this.applyCanonicalCaseForLine(document, lineNumber, editorHint, options);
	}

	private _applyCanonicalPosition(
		document: vscode.TextDocument,
		position: vscode.Position,
		editorHint?: vscode.TextEditor,
	): void {
		if (!this._canonicalLineWasTouched(document, position.line)) {
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
