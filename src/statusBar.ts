import * as vscode from 'vscode';
import * as path from 'path';
import { decodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';

/**
 * One status bar item: shows "XLIDE: <project> | <module>" when the focused
 * editor is a xlide-vba:// document. Click to refresh the XLIDE explorer.
 */
export class XlideStatusBar implements vscode.Disposable {
    private readonly _activeItem: vscode.StatusBarItem;
    private readonly _disposables: vscode.Disposable[] = [];

    constructor() {
        this._activeItem = vscode.window.createStatusBarItem(
            vscode.StatusBarAlignment.Left,
            100,
        );
        this._activeItem.command = 'xlide.refreshExplorer';
        this._activeItem.tooltip = 'XLIDE: refresh sidebar';

        this._disposables.push(
            this._activeItem,
            vscode.window.onDidChangeActiveTextEditor(() => this._refreshActive()),
        );

        this._refreshActive();
    }

    private _refreshActive(): void {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
            this._activeItem.hide();
            return;
        }
        try {
            const { projectPath, moduleName } = decodeModuleUri(editor.document.uri);
            this._activeItem.text = `$(file-code) ${path.basename(projectPath)} | ${moduleName}`;
            this._activeItem.tooltip = projectPath;
            this._activeItem.show();
        } catch {
            this._activeItem.hide();
        }
    }

    dispose(): void {
        for (const d of this._disposables) { d.dispose(); }
    }
}
