import * as vscode from 'vscode';
import * as path from 'path';
import { decodeModuleUri, XLIDE_SCHEME, XLIDE_LIVESHARE_AUTHORITY } from './xlideFileSystem';
import { LiveShareIntegration } from './liveShare';

/**
 * Two status bar items:
 *   - Active module: shows "XLIDE: <workbook> | <module>" when the focused
 *     editor is a xlide-vba:// document. Click to reveal in the XLIDE sidebar.
 *   - Live Share: shows "XLIDE (Live Share): N" while connected as a guest.
 */
export class XlideStatusBar implements vscode.Disposable {
    private readonly _activeItem: vscode.StatusBarItem;
    private readonly _liveShareItem: vscode.StatusBarItem;
    private readonly _disposables: vscode.Disposable[] = [];

    constructor(private readonly _liveShare: LiveShareIntegration) {
        this._activeItem = vscode.window.createStatusBarItem(
            vscode.StatusBarAlignment.Left,
            100,
        );
        this._activeItem.command = 'xlide.refreshExplorer';
        this._activeItem.tooltip = 'XLIDE: refresh sidebar';

        this._liveShareItem = vscode.window.createStatusBarItem(
            vscode.StatusBarAlignment.Left,
            99,
        );
        this._liveShareItem.command = 'xlide.refreshExplorer';

        this._disposables.push(
            this._activeItem,
            this._liveShareItem,
            vscode.window.onDidChangeActiveTextEditor(() => this._refreshActive()),
            this._liveShare.onDidChange(() => this._refreshLiveShare()),
        );

        this._refreshActive();
        this._refreshLiveShare();
    }

    private _refreshActive(): void {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
            this._activeItem.hide();
            return;
        }
        try {
            const uri = editor.document.uri;
            if (uri.authority === XLIDE_LIVESHARE_AUTHORITY) {
                // Remote module: we don't have the workbook name on the guest
                // (only an opaque id), so show a generic indicator.
                this._activeItem.text = `$(file-code) XLIDE (Live Share)`;
                this._activeItem.tooltip = uri.toString();
            } else {
                const { xlsmPath, moduleName } = decodeModuleUri(uri);
                this._activeItem.text = `$(file-code) ${path.basename(xlsmPath)} | ${moduleName}`;
                this._activeItem.tooltip = xlsmPath;
            }
            this._activeItem.show();
        } catch {
            this._activeItem.hide();
        }
    }

    private _refreshLiveShare(): void {
        if (this._liveShare.isGuest) {
            this._liveShareItem.text = '$(remote) XLIDE: Live Share';
            this._liveShareItem.tooltip = 'XLIDE is connected to a Live Share host. Click to refresh.';
            this._liveShareItem.show();
        } else if (this._liveShare.isHost) {
            this._liveShareItem.text = '$(broadcast) XLIDE: Sharing';
            this._liveShareItem.tooltip = 'XLIDE is sharing this workspace over Live Share.';
            this._liveShareItem.show();
        } else {
            this._liveShareItem.hide();
        }
    }

    dispose(): void {
        for (const d of this._disposables) { d.dispose(); }
    }
}
