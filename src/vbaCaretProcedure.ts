import * as vscode from 'vscode';
import { moduleLocationOfDocument } from './vbaDocumentLocation';
import {
    vbaProcedureAtLine,
    vbaProcedureLabel,
    vbaProcedureRanges,
    type VbaProcedureRange,
} from './vbaProcedureAtLine';

/** Where the caret is, in the terms the tree and the status bar both use. */
export interface VbaCaretPosition {
    projectPath: string;
    moduleName: string;
    /** The module IS a file (a VB6 project), rather than a virtual document. */
    native: boolean;
    /** Undefined in the declarations section, above the first procedure. */
    procedure: VbaProcedureRange | undefined;
    /** `Sub Post`, or `(Declarations)` - what both surfaces show. */
    label: string;
}

/**
 * Follows the caret from one VBA procedure to the next, so the status bar and
 * the explorer agree on where you are without each keeping its own answer.
 *
 * Fires only when the module or the procedure actually changes: the caret
 * moves on every keystroke and arrow key, and most of those moves stay inside
 * the procedure they started in.
 */
export class VbaCaretProcedureTracker implements vscode.Disposable {
    private readonly _emitter = new vscode.EventEmitter<VbaCaretPosition | undefined>();
    readonly onDidChange = this._emitter.event;
    private readonly _disposables: vscode.Disposable[] = [];
    private _current: VbaCaretPosition | undefined;
    /**
     * The active document's procedure ranges, keyed by document and version.
     * Rescanning a large module on every caret move would be the whole cost of
     * this; the version makes an edit, and only an edit, pay for it.
     */
    private _ranges: { key: string; ranges: VbaProcedureRange[] } | undefined;

    constructor() {
        this._disposables.push(
            this._emitter,
            vscode.window.onDidChangeActiveTextEditor(() => this._update()),
            vscode.window.onDidChangeTextEditorSelection((e) => {
                if (e.textEditor === vscode.window.activeTextEditor) {
                    this._update();
                }
            }),
        );
        this._update();
    }

    /** Where the caret is now, for a surface that has just been created. */
    get current(): VbaCaretPosition | undefined {
        return this._current;
    }

    private _update(): void {
        const next = this._read();
        if (samePosition(this._current, next)) {
            return;
        }
        this._current = next;
        this._emitter.fire(next);
    }

    private _read(): VbaCaretPosition | undefined {
        const editor = vscode.window.activeTextEditor;
        if (!editor) {
            return undefined;
        }
        const location = moduleLocationOfDocument(editor.document);
        if (!location) {
            return undefined;
        }
        const key = `${editor.document.uri.toString()}@${editor.document.version}`;
        if (this._ranges?.key !== key) {
            this._ranges = { key, ranges: vbaProcedureRanges(editor.document.getText()) };
        }
        const procedure = vbaProcedureAtLine(this._ranges.ranges, editor.selection.active.line);
        return {
            projectPath: location.projectPath,
            moduleName: location.moduleName,
            native: location.native,
            procedure,
            label: vbaProcedureLabel(procedure),
        };
    }

    dispose(): void {
        for (const d of this._disposables) { d.dispose(); }
    }
}

function samePosition(left: VbaCaretPosition | undefined, right: VbaCaretPosition | undefined): boolean {
    if (!left || !right) {
        return left === right;
    }
    return left.projectPath === right.projectPath
        && left.moduleName === right.moduleName
        && left.label === right.label;
}
