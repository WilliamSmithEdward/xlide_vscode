// VBA typing automation: onDidChangeTextDocument editors (not language
// providers) for VBA-IDE-style smart Enter auto-block insertion, With-member
// line continuation, and For/Next loop-iterator name sync.
//
// Extracted verbatim from vbaLanguageProviders.ts (audit #21).

import * as vscode from 'vscode';
import { isVbaDocument } from './xlideFileSystem';
import { lexerStrippedLine, lexerStrippedLines } from './analyzer/lexer/strippedLines';
import {
    commentContinuationText,
    detectSmartBlockOpener,
    isSmartBlockClosedAhead,
    procedureHeaderParensEdit,
    resolveLoopIteratorSyncEdit,
    smartBlockInsertion,
    withMemberContinuationText,
} from './vbaSmartEnter';
import {
    xlideEditorBlockLayoutFromConfig,
    xlideEditorContinueCommentOnNewlineFromConfig,
    xlideEditorMirrorCommentSpacingFromConfig,
} from './globalSettings';

/**
 * VBA-IDE-style smart Enter: typing a block opener and pressing Enter
 * auto-inserts the matching closer below, leaving the cursor on the indented
 * body line. `With` also seeds the body line with `.` so member completion can
 * start immediately.
 */
export function registerVbaAutoBlock(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        // React only to a plain Enter (newline plus optional auto-indent),
        // never to pastes or multi-character insertions.
        if (!/^\r?\n[ \t]*$/.test(change.text)) { return; }

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
            applying = true;
            try {
                if (await maybeContinueCommentLine(doc, openerLineIndex)) { return; }
                await maybeContinueWithMemberLine(doc, openerLineIndex);
            } finally {
                applying = false;
            }
            return;
        }

        const bodyLineIndex = openerLineIndex + 1;
        if (bodyLineIndex >= doc.lineCount) { return; }

        const strippedLines = lexerStrippedLines(doc.getText());
        strippedLines[openerLineIndex] = lexerStrippedLine(normalizedOpenerLine);
        const closedAhead = isSmartBlockClosedAhead(strippedLines, openerLineIndex, opener);

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

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

        applying = true;
        try {
            await editor.edit(
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
        } finally {
            applying = false;
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
        setTimeout(placeCaret, 0);
    });

    context.subscriptions.push(sub);
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
    const lineText = commentContinuationText(doc.getText(), previousLineIndex, mirrorSpacing);
    if (lineText === undefined) { return false; }

    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document !== doc) { return false; }

    const bodyRange = new vscode.Range(
        new vscode.Position(bodyLineIndex, 0),
        new vscode.Position(bodyLineIndex, bodyLine.length),
    );
    const applied = await editor.edit(
        (eb) => eb.replace(bodyRange, lineText),
        { undoStopBefore: false, undoStopAfter: true },
    );
    if (!applied) { return false; }

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
    setTimeout(placeCaret, 0);
    return true;
}

async function maybeContinueWithMemberLine(
    doc: vscode.TextDocument,
    previousLineIndex: number,
): Promise<void> {
    const bodyLineIndex = previousLineIndex + 1;
    if (bodyLineIndex >= doc.lineCount) { return; }

    const bodyLine = doc.lineAt(bodyLineIndex).text;
    if (!/^[ \t]*$/.test(bodyLine)) { return; }

    const lineText = withMemberContinuationText(doc.getText(), previousLineIndex);
    if (!lineText) { return; }

    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document !== doc) { return; }

    const bodyRange = new vscode.Range(
        new vscode.Position(bodyLineIndex, 0),
        new vscode.Position(bodyLineIndex, bodyLine.length),
    );
    const applied = await editor.edit(
        (eb) => eb.replace(bodyRange, lineText),
        { undoStopBefore: false, undoStopAfter: true },
    );
    if (!applied) { return; }

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
    setTimeout(placeCaret, 0);
}

/**
 * Keeps simple loop iterator names paired across `For`/`For Each` and `Next`.
 * This intentionally lives outside snippets so hand-written loops get the same
 * behavior as completed loops.
 */
export function registerVbaLoopIteratorSync(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        if (/[\r\n]/.test(change.text)) { return; }

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const lineIndex = Math.min(change.range.start.line, doc.lineCount - 1);
        const lineLength = doc.lineAt(lineIndex).text.length;
        const character = Math.min(lineLength, change.range.start.character + change.text.length);
        const offset = doc.offsetAt(new vscode.Position(lineIndex, character));
        const syncEdit = resolveLoopIteratorSyncEdit(doc.getText(), offset);
        if (!syncEdit) { return; }

        applying = true;
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
            applying = false;
        }
    });

    context.subscriptions.push(sub);
}
