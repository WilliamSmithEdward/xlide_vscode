// The `.` XLIDE puts on a With body line - from Smart Enter, its With-member
// continuation and the With snippet - is a convenience for typing a member. A
// lone `.` is no statement (Excel rejects it: oracle
// bare_leading_member_access_inside_with_compile), so when the line is left
// with nothing typed after the dot, the dot goes (issue #1620). Only a dot
// XLIDE inserted is removed; one the user typed stays.
//
// A dot is followed as edits move its line. It is armed once the caret has
// been on its line, however the caret got there, and from then on it goes
// when the user's own navigation or Enter leaves the line. Until then it goes
// only when the caret leaves its whole block, so neither Smart Enter's own
// caret placement nor filling in the snippet's object removes it.

import * as vscode from 'vscode';
import { isVbaDocument } from './xlideFileSystem';

/** A dot XLIDE inserted, and the lines of the With block it was inserted in. */
export interface WithAutoDot {
	/** The dot's line, and that line's text when XLIDE wrote it. */
	line: number;
	text: string;
	/** The block's With line and End With line. */
	first: number;
	last: number;
	/** Whether the caret has been on the dot's line since it was inserted. */
	armed: boolean;
}

/** The part of a change that moves lines. */
export interface LineChange {
	startLine: number;
	endLine: number;
	text: string;
}

/**
 * The dot after a change: lines moved by an edit above it, or undefined when
 * the edit replaced the dot's line along with others.
 */
export function shiftWithAutoDot(dot: WithAutoDot, change: LineChange): WithAutoDot | undefined {
	const delta = (change.text.match(/\r\n|\r|\n/g)?.length ?? 0) - (change.endLine - change.startLine);
	const shifted = (line: number): number => (change.endLine < line ? line + delta : line);
	if (change.startLine <= dot.line && change.endLine >= dot.line && change.startLine !== change.endLine) {
		return undefined;
	}
	const touchesBlock = change.startLine <= dot.last && change.endLine >= dot.first;
	return {
		...dot,
		line: shifted(dot.line),
		first: shifted(dot.first),
		last: touchesBlock && change.startLine >= dot.first ? Math.max(dot.first, dot.last + delta) : shifted(dot.last),
	};
}

/** Whether the dot should go now: its line is as XLIDE left it, and the caret has left. */
export function withAutoDotUnused(dot: WithAutoDot, lineText: string | undefined, caretLine: number): boolean {
	if (lineText !== dot.text || caretLine === dot.line) { return false; }
	return dot.armed || caretLine < dot.first || caretLine > dot.last;
}

/**
 * The dot a With snippet's text carries, placed at `startLine`, or undefined
 * for any other insertion. The snippet is one With line, the body, and End With.
 */
export function withSnippetAutoDot(text: string, startLine: number): WithAutoDot | undefined {
	const lines = text.split(/\r\n|\r|\n/);
	if (lines.length < 3 || !/^[ \t]*With\b/i.test(lines[0]) || !/^[ \t]*End[ \t]+With[ \t]*$/i.test(lines[lines.length - 1])) {
		return undefined;
	}
	const index = lines.findIndex((line, at) => at > 0 && at < lines.length - 1 && /^[ \t]*\.$/.test(line));
	if (index < 0) { return undefined; }
	return { line: startLine + index, text: lines[index], first: startLine, last: startLine + lines.length - 1, armed: false };
}

const dots = new Map<vscode.TextDocument, WithAutoDot[]>();

/** Notes a dot XLIDE wrote on `line`, whose text is now `text`, in the block from `first` to `last`. */
export function trackWithAutoDot(document: vscode.TextDocument, line: number, text: string, first: number, last: number): void {
	if (!text.endsWith('.')) { return; }
	const list = (dots.get(document) ?? []).filter((dot) => dot.line !== line);
	// The edit that wrote the dot may already have carried the caret onto it.
	const editor = vscode.window.activeTextEditor;
	const armed = editor?.document === document && editor.selection.active.line === line;
	list.push({ line, text, first, last, armed });
	dots.set(document, list);
}

/** Arms the dots on the caret's line. */
function arm(editor: vscode.TextEditor): void {
	const caretLine = editor.selection.active.line;
	for (const dot of dots.get(editor.document) ?? []) {
		if (dot.line === caretLine) { dot.armed = true; }
	}
}

/** Removes every tracked dot of `editor`'s document that is no longer used. */
async function clearUnused(editor: vscode.TextEditor): Promise<void> {
	const document = editor.document;
	const list = dots.get(document);
	if (!list || document.isClosed) { return; }
	const caretLine = editor.selection.active.line;
	const stale: WithAutoDot[] = [];
	for (const dot of list) {
		if (caretLine === dot.line) { dot.armed = true; }
		const lineText = dot.line < document.lineCount ? document.lineAt(dot.line).text : undefined;
		if (lineText !== dot.text) { stale.push(dot); continue; }
		if (withAutoDotUnused(dot, lineText, caretLine)) { stale.push(dot); }
	}
	const unused = stale.filter((dot) => dot.line < document.lineCount && document.lineAt(dot.line).text === dot.text);
	dots.set(document, list.filter((dot) => !stale.includes(dot)));
	if (unused.length === 0) { return; }
	await editor.edit((edit) => {
		for (const dot of unused) {
			edit.delete(new vscode.Range(dot.line, dot.text.length - 1, dot.line, dot.text.length));
		}
	});
}

/** Follows the dots through edits and removes each one once its line is left unused. */
export function registerWithAutoDotCleanup(context: vscode.ExtensionContext): void {
	const later = (editor: vscode.TextEditor): void => {
		setTimeout(() => {
			if (vscode.window.activeTextEditor === editor) { void clearUnused(editor); }
		}, 0);
	};
	context.subscriptions.push(
		vscode.workspace.onDidChangeTextDocument((event) => {
			const document = event.document;
			if (!isVbaDocument(document)) { return; }
			let list = dots.get(document) ?? [];
			let lineBreak = false;
			for (const change of event.contentChanges) {
				const lineChange = { startLine: change.range.start.line, endLine: change.range.end.line, text: change.text };
				list = list.flatMap((dot) => shiftWithAutoDot(dot, lineChange) ?? []);
				lineBreak ||= /\r|\n/.test(change.text);
				if (event.contentChanges.length === 1 && event.reason === undefined) {
					const snippet = withSnippetAutoDot(change.text, change.range.start.line);
					if (snippet) { list = [...list.filter((dot) => dot.line !== snippet.line), snippet]; }
				}
			}
			if (list.length === 0) { dots.delete(document); return; }
			dots.set(document, list);
			// Enter on the dot's line moves the caret off it with no navigation.
			const editor = vscode.window.activeTextEditor;
			if (lineBreak && editor?.document === document) { later(editor); }
		}),
		vscode.window.onDidChangeTextEditorSelection((event) => {
			if (!dots.has(event.textEditor.document)) { return; }
			arm(event.textEditor);
			// A selection an edit or an extension set has no kind; only the
			// user's own navigation leaves the line.
			if (event.kind !== undefined) { later(event.textEditor); }
		}),
		vscode.workspace.onDidCloseTextDocument((document) => { dots.delete(document); }),
	);
}
