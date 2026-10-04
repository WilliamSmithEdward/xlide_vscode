import * as vscode from 'vscode';
import { moduleLocationOfDocument } from './vbaDocumentLocation';
import {
    vbaProcedureLabel,
    vbaProcedureRanges,
    vbaProcedureLineStructure,
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
     * Procedure ranges for each document's current version. Tab switches reuse
     * them; an edit rescans only that document. Weak keys do not keep closed
     * documents alive, and a reopened document starts with a fresh entry.
     */
    private readonly _ranges = new WeakMap<vscode.TextDocument, { version: number; ranges: VbaProcedureRange[]; lineStructures: string[] }>();

    constructor() {
        this._disposables.push(
            this._emitter,
            vscode.workspace.onDidChangeTextDocument((event) => this._acceptNonStructuralEdit(event)),
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
        const document = editor.document;
        let cached = this._ranges.get(document);
        if (cached?.version !== document.version) {
            const source = document.getText();
            cached = {
                version: document.version,
                ranges: vbaProcedureRanges(source),
                lineStructures: source.split(/\r\n|\r|\n/).map(vbaProcedureLineStructure),
            };
            this._ranges.set(document, cached);
        }
        const procedure = orderedProcedureAtLine(cached.ranges, editor.selection.active.line);
        return {
            projectPath: location.projectPath,
            moduleName: location.moduleName,
            native: location.native,
            procedure,
            label: vbaProcedureLabel(procedure),
        };
    }

    /** Reuse ranges when an edit cannot change headers or their leading lines. */
    private _acceptNonStructuralEdit(event: vscode.TextDocumentChangeEvent): void {
        if (event.contentChanges.length === 0) { return; }
        const document = event.document;
        const cached = this._ranges.get(document);
        if (!cached || cached.version !== document.version - 1 ||
            cached.lineStructures.length !== document.lineCount) { return; }
        for (const change of event.contentChanges) {
            const line = change.range.start.line;
            if (line !== change.range.end.line || /[\r\n]/.test(change.text) ||
                cached.lineStructures[line] !== vbaProcedureLineStructure(document.lineAt(line).text)) {
                return;
            }
        }
        cached.version = document.version;
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

/** Tracker-owned scanner ranges are ordered and disjoint; callers supplying
 * arbitrary ranges to vbaProcedureAtLine retain its existing first-match semantics.
 */
function orderedProcedureAtLine(ranges: readonly VbaProcedureRange[], line: number): VbaProcedureRange | undefined {
    const initial = ranges[0];
    if (!initial || line <= initial.lastLine) {
        return initial && line >= initial.firstLine ? initial : undefined;
    }
    let first = 1, end = ranges.length;
    while (first < end) {
        const middle = first + Math.floor((end - first) / 2);
        if (ranges[middle].firstLine <= line) {
            first = middle + 1;
        } else {
            end = middle;
        }
    }
    const range = ranges[first - 1];
    return range && line <= range.lastLine ? range : undefined;
}
