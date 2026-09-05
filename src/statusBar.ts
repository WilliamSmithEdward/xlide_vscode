import * as vscode from 'vscode';
import * as path from 'path';
import type { VbaCaretPosition, VbaCaretProcedureTracker } from './vbaCaretProcedure';

/**
 * One status bar item: shows "XLIDE: <project> | <module> | <procedure>" when
 * the focused editor is a xlide-vba:// document, the way the VBE names where
 * the caret is. Click to refresh the XLIDE explorer.
 */
export class XlideStatusBar implements vscode.Disposable {
    private readonly _activeItem: vscode.StatusBarItem;
    private readonly _disposables: vscode.Disposable[] = [];

    constructor(caret: VbaCaretProcedureTracker) {
        this._activeItem = vscode.window.createStatusBarItem(
            vscode.StatusBarAlignment.Left,
            100,
        );
        this._activeItem.command = 'xlide.refreshExplorer';
        this._activeItem.tooltip = 'XLIDE: refresh sidebar';

        this._disposables.push(
            this._activeItem,
            caret.onDidChange((position) => this._render(position)),
        );

        this._render(caret.current);
    }

    private _render(position: VbaCaretPosition | undefined): void {
        // A VB6 module is a file the editor already names in its own tab; the
        // item speaks for the modules that have no file of their own.
        if (!position || position.native) {
            this._activeItem.hide();
            return;
        }
        this._activeItem.text = '$(file-code) '
            + `${path.basename(position.projectPath)} | ${position.moduleName} | ${position.label}`;
        this._activeItem.tooltip = position.projectPath;
        this._activeItem.show();
    }

    dispose(): void {
        for (const d of this._disposables) { d.dispose(); }
    }
}
