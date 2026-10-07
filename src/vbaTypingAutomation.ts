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

        const caretOwnership = enterCaretOwnership(editor,
            new vscode.Position(bodyLineIndex, bodyLine.length),
            new vscode.Position(bodyLineIndex + smartBlock.bodyLineOffset, smartBlock.bodyText.length));
        applying.add(doc);
        let applied = false;
        try {
            applied = await editor.edit(
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
            if (!applied) { caretOwnership.dispose(); }
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
            suggestAfterAutoDot(editor, smartBlock.bodyText);
        };
        caretOwnership.schedule(placeCaret);
    });

    context.subscriptions.push(sub);
}

/** Tracks native navigation until the editor's pending edit has settled. */
function enterCaretOwnership(editor: vscode.TextEditor, before: vscode.Position, after: vscode.Position) {
    const document = editor.document;
    const version = document.version;
    const initial = editor.selection;
    let navigated = false;
    const subscription = vscode.window.onDidChangeTextEditorSelection(event => {
        // Text-edit adjustments have no selection kind. Explicit keyboard,
        // mouse or command navigation takes ownership away from this edit.
        if (event.textEditor === editor && event.kind !== undefined) { navigated = true; }
    });
    const at = (selection: vscode.Selection, position: vscode.Position): boolean =>
        selection.active.line === position.line && selection.active.character === position.character &&
        selection.anchor.line === position.line && selection.anchor.character === position.character;
    return {
        dispose: () => subscription.dispose(),
        schedule(placeCaret: () => void): void {
            // Let queued native selection notifications arrive before deciding
            // to place the caret. Keep observing until this deferred pass ends.
            setTimeout(() => {
                try {
                    const current = editor.selection;
                    const unchanged = current.active.line === initial.active.line && current.active.character === initial.active.character &&
                        current.anchor.line === initial.anchor.line && current.anchor.character === initial.anchor.character;
                    if (navigated || vscode.window.activeTextEditor !== editor || editor.document !== document || document.isClosed ||
                        document.version !== version + 1 || (!unchanged && !at(current, before) && !at(current, after))) { return; }
                    placeCaret();
                } finally { subscription.dispose(); }
            }, 0);
        },
    };
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

    await replaceBodyLine(doc, bodyLineIndex, bodyLine, lineText);
}

/**
 * Replaces the blank body line with `lineText`, then places its caret once the
 * editor settles if navigation has not taken ownership. Returns the edited editor.
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
    const caretOwnership = enterCaretOwnership(editor,
        new vscode.Position(bodyLineIndex, bodyLine.length), new vscode.Position(bodyLineIndex, lineText.length));
    let applied = false;
    try {
        applied = await editor.edit(
            (eb) => eb.replace(bodyRange, lineText),
            { undoStopBefore: false, undoStopAfter: true },
        );
    } finally { if (!applied) { caretOwnership.dispose(); } }
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
        suggestAfterAutoDot(editor, lineText);
    };
    caretOwnership.schedule(placeCaret);
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
