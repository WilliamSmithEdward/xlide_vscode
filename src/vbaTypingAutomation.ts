// VBA typing automation: onDidChangeTextDocument editors (not language
// providers) for VBA-IDE-style smart Enter auto-block insertion, With-member
// line continuation, and For/Next loop-iterator name sync.
//
// Extracted verbatim from vbaLanguageProviders.ts (audit #21).

import * as vscode from 'vscode';
import { isVbaDocument } from './xlideFileSystem';
import { lexerStrippedLine, lexerStrippedLines } from './analyzer/lexer/strippedLines';
import {
    commentContinuationForLine,
    detectSmartBlockOpener,
    isSmartBlockClosedAhead,
    procedureHeaderParensEdit,
    resolveLoopIteratorSyncEdit,
    loopIteratorSyncMayApply,
    smartBlockInsertion,
    withMemberContinuationText,
} from './vbaSmartEnter';
import {
    xlideEditorBlockLayoutFromConfig,
    xlideEditorContinueCommentOnNewlineFromConfig,
    xlideEditorMirrorCommentSpacingFromConfig,
} from './globalSettings';

/** The one change of an edit to a VBA document, or undefined for any other event. */
function soleVbaChange(e: vscode.TextDocumentChangeEvent): vscode.TextDocumentContentChangeEvent | undefined {
    if (e.reason !== undefined || e.document.isClosed || !isVbaDocument(e.document) || e.contentChanges.length !== 1) { return undefined; }
    return e.contentChanges[0];
}

/**
 * VBA-IDE-style smart Enter: typing a block opener and pressing Enter
 * auto-inserts the matching closer below, leaving the cursor on the indented
 * body line. `With` also seeds the body line with `.` so member completion can
 * start immediately.
 */
export function registerVbaAutoBlock(context: vscode.ExtensionContext): void {
    const applying = new WeakSet<vscode.TextDocument>();

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        const doc = e.document;
        if (applying.has(doc)) { return; }
        const change = soleVbaChange(e);
        if (!change) { return; }
        // React only to a plain Enter (newline plus optional auto-indent),
        // never to pastes or multi-character insertions.
        if (!/^\r?\n[ \t]*$/.test(change.text)) { return; }
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const openerLineIndex = change.range.start.line;
        const openerLine = doc.lineAt(openerLineIndex).text;
        const headerParensEdit = procedureHeaderParensEdit(openerLine);
        const normalizedOpenerLine = headerParensEdit
            ? `${openerLine.slice(0, headerParensEdit.startCol)}${headerParensEdit.newText}${openerLine.slice(headerParensEdit.endCol)}`
            : openerLine;
        const opener = detectSmartBlockOpener(lexerStrippedLine(normalizedOpenerLine));
        if (!opener) {
            // Hold the re-entrancy guard across the continuation edits too, so a
            // second change event (the edit itself, or a fast follow-up keystroke)
            // cannot start a concurrent continuation on the same line.
            applying.add(doc);
            try {
                if (await maybeContinueCommentLine(doc, openerLineIndex)) { return; }
                await maybeContinueWithMemberLine(doc, openerLineIndex);
            } finally {
                applying.delete(doc);
            }
            return;
        }

        const bodyLineIndex = openerLineIndex + 1;
        if (bodyLineIndex >= doc.lineCount) { return; }

        const strippedLines = lexerStrippedLines(doc.getText());
        strippedLines[openerLineIndex] = lexerStrippedLine(normalizedOpenerLine);
        const closedAhead = isSmartBlockClosedAhead(strippedLines, openerLineIndex, opener);

        const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
        const bodyLine = doc.lineAt(bodyLineIndex).text;
        if (!/^[ \t]*$/.test(bodyLine)) { return; }
        const smartBlock = smartBlockInsertion(normalizedOpenerLine, bodyLine, opener, {
            eol,
            insertCloser: !closedAhead,
            layout: xlideEditorBlockLayoutFromConfig(vscode.workspace.getConfiguration('xlide')).value,
        });
        const bodyRange = new vscode.Range(
            new vscode.Position(bodyLineIndex, 0),
            new vscode.Position(bodyLineIndex, bodyLine.length),
        );

        applying.add(doc);
        try {
            const applied = await editor.edit(
                (eb) => {
                    if (headerParensEdit) {
                        eb.insert(
                            new vscode.Position(openerLineIndex, headerParensEdit.startCol),
                            headerParensEdit.newText,
                        );
                    }
                    eb.replace(
                        bodyRange,
                        smartBlock.replacementText,
                    );
                },
                { undoStopBefore: false, undoStopAfter: true },
            );
            if (!applied) { return; }
        } finally {
            applying.delete(doc);
        }

        // Keep the caret on the indented body line, above the inserted End. The
        // delayed pass wins same-Enter listener races such as canonical casing.
        const placeCaret = (): void => {
            if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
                return;
            }
            const caretLineIndex = bodyLineIndex + smartBlock.bodyLineOffset;
            if (caretLineIndex >= doc.lineCount || doc.lineAt(caretLineIndex).text !== smartBlock.bodyText) {
                return;
            }
            const caret = new vscode.Position(
                caretLineIndex,
                smartBlock.bodyText.length,
            );
            editor.selection = new vscode.Selection(caret, caret);
        };
        placeCaret();
        scheduleCaretRetry(editor, placeCaret);
        suggestAfterAutoDot(editor, smartBlock.bodyText);
    });

    context.subscriptions.push(sub);
}

/** Retry placement only while this edit still owns the document and selection. */
function scheduleCaretRetry(editor: vscode.TextEditor, placeCaret: () => void): void {
    const document = editor.document;
    const version = document.version;
    const selection = editor.selection;
    setTimeout(() => {
        if (vscode.window.activeTextEditor !== editor) { return; }
        const current = editor.selection;
        if (document.isClosed || document.version !== version || editor.document !== document ||
            current.active.line !== selection.active.line || current.active.character !== selection.active.character ||
            current.anchor.line !== selection.anchor.line || current.anchor.character !== selection.anchor.character) {
            return;
        }
        placeCaret();
    }, 0);
}

/**
 * Opens the completion list after an auto-inserted leading `.`.
 *
 * Typing a dot triggers the suggest widget because `.` is a registered trigger
 * character, but a dot the editor inserts is not typed, so dropping into a
 * `With` body left the caret after a dot with no list - and backspacing over it
 * and retyping the same character was the only way to see one.
 */
function suggestAfterAutoDot(editor: vscode.TextEditor, expectedLine: string): void {
    if (!expectedLine.endsWith('.')) {
        return;
    }
    const document = editor.document;
    const version = document.version;
    const line = editor.selection.active.line;
    // After the caret settles: the delayed placement pass wins same-Enter
    // listener races, and asking earlier would target the pre-edit position.
    setTimeout(() => {
        if (vscode.window.activeTextEditor !== editor || editor.document !== document ||
            document.isClosed || document.version !== version) {
            return;
        }
        const caret = editor.selection.active;
        if (!editor.selection.isEmpty || caret.line !== line || caret.character !== expectedLine.length ||
            document.lineAt(caret.line).text !== expectedLine) {
            return;
        }
        void vscode.commands.executeCommand('editor.action.triggerSuggest');
    }, 0);
}

/**
 * Continues a whole-line VBA comment on Enter: the new line starts with the same
 * apostrophe run and (per the mirror-spacing setting) the same trailing spaces.
 * Gated by the editor.continueCommentOnNewline setting; returns true when it
 * applied so the caller skips other continuations.
 */
async function maybeContinueCommentLine(
    doc: vscode.TextDocument,
    previousLineIndex: number,
): Promise<boolean> {
    const config = vscode.workspace.getConfiguration('xlide');
    if (!xlideEditorContinueCommentOnNewlineFromConfig(config).value) {
        return false;
    }
    const bodyLineIndex = previousLineIndex + 1;
    if (bodyLineIndex >= doc.lineCount) { return false; }

    const bodyLine = doc.lineAt(bodyLineIndex).text;
    if (!/^[ \t]*$/.test(bodyLine)) { return false; }

    const mirrorSpacing = xlideEditorMirrorCommentSpacingFromConfig(config).value;
    const lineText = commentContinuationForLine(doc.lineAt(previousLineIndex).text, mirrorSpacing);
    if (lineText === undefined) { return false; }

    return (await replaceBodyLine(doc, bodyLineIndex, bodyLine, lineText)) !== undefined;
}

async function maybeContinueWithMemberLine(
    doc: vscode.TextDocument,
    previousLineIndex: number,
): Promise<void> {
    const bodyLineIndex = previousLineIndex + 1;
    if (bodyLineIndex >= doc.lineCount) { return; }

    const bodyLine = doc.lineAt(bodyLineIndex).text;
    if (!/^[ \t]*$/.test(bodyLine)) { return; }

    // Ordinary Enter cannot continue a With member. Inspect the one line
    // before materializing and scanning the whole module for an open With.
    if (!/^[ \t]*\./.test(lexerStrippedLine(doc.lineAt(previousLineIndex).text))) { return; }
    const lineText = withMemberContinuationText(doc.getText(), previousLineIndex);
    if (!lineText) { return; }

    const editor = await replaceBodyLine(doc, bodyLineIndex, bodyLine, lineText);
    if (editor) {
        suggestAfterAutoDot(editor, lineText);
    }
}

/**
 * Replaces the blank body line with `lineText` and parks the caret at its end,
 * now and again once the editor settles. The editor, when the edit applied.
 */
async function replaceBodyLine(
    doc: vscode.TextDocument,
    bodyLineIndex: number,
    bodyLine: string,
    lineText: string,
): Promise<vscode.TextEditor | undefined> {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document !== doc) { return undefined; }

    const bodyRange = new vscode.Range(
        new vscode.Position(bodyLineIndex, 0),
        new vscode.Position(bodyLineIndex, bodyLine.length),
    );
    const applied = await editor.edit(
        (eb) => eb.replace(bodyRange, lineText),
        { undoStopBefore: false, undoStopAfter: true },
    );
    if (!applied) { return undefined; }

    const placeCaret = (): void => {
        if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
            return;
        }
        if (bodyLineIndex >= doc.lineCount || doc.lineAt(bodyLineIndex).text !== lineText) {
            return;
        }
        const caret = new vscode.Position(bodyLineIndex, lineText.length);
        editor.selection = new vscode.Selection(caret, caret);
    };
    placeCaret();
    scheduleCaretRetry(editor, placeCaret);
    return editor;
}

/**
 * Keeps simple loop iterator names paired across `For`/`For Each` and `Next`.
 * This intentionally lives outside snippets so hand-written loops get the same
 * behavior as completed loops.
 */
export function registerVbaLoopIteratorSync(context: vscode.ExtensionContext): void {
    const applying = new WeakSet<vscode.TextDocument>();

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        const doc = e.document;
        if (applying.has(doc)) { return; }
        const change = soleVbaChange(e);
        if (!change) { return; }
        if (/[\r\n]/.test(change.text)) { return; }

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const lineIndex = Math.min(change.range.start.line, doc.lineCount - 1);
        const lineText = doc.lineAt(lineIndex).text;
        const character = Math.min(lineText.length, change.range.start.character + change.text.length);
        if (!loopIteratorSyncMayApply(lineText, character)) { return; }
        const offset = doc.offsetAt(new vscode.Position(lineIndex, character));
        const syncEdit = resolveLoopIteratorSyncEdit(doc.getText(), offset);
        if (!syncEdit) { return; }

        applying.add(doc);
        try {
            await editor.edit(
                (eb) => eb.replace(
                    new vscode.Range(
                        doc.positionAt(syncEdit.span.start),
                        doc.positionAt(syncEdit.span.end),
                    ),
                    syncEdit.newText,
                ),
                { undoStopBefore: false, undoStopAfter: false },
            );
        } finally {
            applying.delete(doc);
        }
    });

    context.subscriptions.push(sub);
}
