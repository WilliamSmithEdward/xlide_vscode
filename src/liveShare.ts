import * as vscode from 'vscode';
import * as path from 'path';
import type * as vsls from 'vsls/vscode';
import { PythonBridge } from './pythonBridge';

const SERVICE_NAME = 'WilliamSmithE.xlide';

// RPC method names (kept identical between host and guest).
const RPC_LIST_WORKBOOKS = 'listWorkbooks';
const RPC_LIST_MODULES = 'listModules';
const RPC_READ_MODULE = 'readModule';
const RPC_WRITE_MODULE = 'writeModule';
const RPC_LIST_SUBS = 'listSubs';

const NOTIFY_WORKBOOKS_CHANGED = 'workbooksChanged';
const NOTIFY_FILE_CHANGED = 'fileChanged';

export interface RemoteWorkbookInfo {
    /** Opaque, session-stable id. */
    id: string;
    /** Display name (basename). */
    name: string;
    /** Workspace-relative folder for display, or empty string. */
    relativeFolder: string;
}

export interface RemoteModuleInfo {
    name: string;
    type: string;
}

export interface RemoteSubInfo {
    name: string;
    kind: string;
    line: number;
}

/**
 * Encodes a remote (Live Share) module reference as an xlide-vba:// URI with
 * the special authority "liveshare". The local provider routes these to the
 * Live Share proxy instead of the local Python bridge.
 *
 *   xlide-vba://liveshare/<workbookId>/<moduleName>.bas
 */
export function encodeRemoteModuleUri(workbookId: string, moduleName: string): vscode.Uri {
    return vscode.Uri.parse(
        `xlide-vba://liveshare/${encodeURIComponent(workbookId)}/${encodeURIComponent(moduleName)}.bas`,
        true,
    );
}

export function decodeRemoteModuleUri(uri: vscode.Uri): { workbookId: string; moduleName: string } {
    if (uri.authority !== 'liveshare') {
        throw new Error(`Not a remote xlide-vba URI: ${uri.toString()}`);
    }
    const match = uri.path.match(/^\/([^/]+)\/([^/]+)\.bas$/i);
    if (!match) {
        throw new Error(`Cannot decode remote xlide-vba URI: ${uri.toString()}`);
    }
    return {
        workbookId: decodeURIComponent(match[1]),
        moduleName: decodeURIComponent(match[2]),
    };
}

/**
 * Live Share integration for XLIDE.
 *
 * Capabilities exposed over the shared service (host -> guest):
 *  - listWorkbooks
 *  - listModules
 *  - listSubs
 *  - readModule
 *  - writeModule
 *
 * Macro execution, exports, and configuration changes are intentionally NOT
 * exposed: those remain host-only commands.
 */
export class LiveShareIntegration implements vscode.Disposable {
    private _api: vsls.LiveShare | null = null;
    private _hostService: vsls.SharedService | null = null;
    private _guestProxy: vsls.SharedServiceProxy | null = null;

    /** Host-side: workbookId -> absolute file path. */
    private _hostWorkbooks = new Map<string, string>();
    /** Host-side: file path -> workbookId. */
    private _hostPathToId = new Map<string, string>();

    private readonly _onDidChange = new vscode.EventEmitter<void>();
    /** Fires when role or workbook list changes — UI should refresh. */
    readonly onDidChange = this._onDidChange.event;

    private _disposables: vscode.Disposable[] = [];

    constructor(
        private readonly _bridge: PythonBridge,
        private readonly _out: vscode.OutputChannel,
    ) {}

    async start(): Promise<void> {
        this._out.appendLine('[LiveShare] start() invoked');
        let mod: typeof vsls;
        try {
            mod = await import('vsls/vscode');
        } catch (err) {
            this._out.appendLine(`[LiveShare] API module not available: ${err}`);
            return;
        }
        try {
            this._api = await mod.getApi('WilliamSmithE.xlide');
        } catch (err) {
            this._out.appendLine(`[LiveShare] getApi failed: ${err}`);
            return;
        }
        if (!this._api) {
            this._out.appendLine('[LiveShare] extension not installed; integration disabled.');
            return;
        }
        this._out.appendLine('[LiveShare] API obtained');
        const api = this._api;
        this._disposables.push(api.onDidChangeSession((e) => this._onSessionChange(e)));
        // Handle the case where activation happens after a session is already in progress
        this._out.appendLine(`[LiveShare] initial session role=${api.session?.role}, id=${api.session?.id ?? '<none>'}`);
        await this._onSessionChange({ session: api.session });

        // Refresh remote workbook list when the host's xlsm files change.
        const watcher = vscode.workspace.createFileSystemWatcher('**/*.{xlsm,xlsb,xlam}');
        const refresh = () => void this._refreshHostWorkbooks();
        watcher.onDidCreate(refresh);
        watcher.onDidDelete(refresh);
        this._disposables.push(watcher);
    }

    private async _onSessionChange(e: vsls.SessionChangeEvent): Promise<void> {
        const role = e.session.role;
        const roleName = role === 1 ? 'Host' : role === 2 ? 'Guest' : `None(${role})`;
        this._out.appendLine(`[LiveShare] session change -> role=${roleName}`);
        // Tear down any previous state
        if (this._hostService) {
            try { await this._api?.unshareService(SERVICE_NAME); } catch { /* ignore */ }
            this._hostService = null;
        }
        this._guestProxy = null;

        if (!this._api) { return; }

        if (role === 1 /* Host */) {
            await this._initHost();
        } else if (role === 2 /* Guest */) {
            await this._initGuest();
        }
        this._onDidChange.fire();
    }

    // ------------------------------------------------------------------
    // Host side
    // ------------------------------------------------------------------

    private async _initHost(): Promise<void> {
        if (!this._api) { return; }
        this._out.appendLine(`[LiveShare] host: sharing service '${SERVICE_NAME}'...`);
        let svc: vsls.SharedService | null;
        try {
            svc = await this._api.shareService(SERVICE_NAME);
        } catch (err) {
            this._out.appendLine(`[LiveShare] host: shareService failed: ${err}`);
            return;
        }
        if (!svc) {
            this._out.appendLine('[LiveShare] host: shareService returned null.');
            this._out.appendLine('[LiveShare] host: This is a Live Share restriction on third-party extensions.');
            this._out.appendLine('[LiveShare] host: Fix: set "liveshare.featureSet": "insiders" in VS Code settings on host AND guest, then reload both windows.');
            void this._promptFeatureSetFix();
            return;
        }
        this._hostService = svc;
        svc.onRequest(RPC_LIST_WORKBOOKS, () => this._handleListWorkbooks());
        svc.onRequest(RPC_LIST_MODULES, (args) => this._handleListModules(args));
        svc.onRequest(RPC_LIST_SUBS, (args) => this._handleListSubs(args));
        svc.onRequest(RPC_READ_MODULE, (args) => this._handleReadModule(args));
        svc.onRequest(RPC_WRITE_MODULE, (args) => this._handleWriteModule(args));
        await this._refreshHostWorkbooks();
        this._out.appendLine(`[LiveShare] host: ready, sharing ${this._hostWorkbooks.size} workbook(s)`);
    }

    private async _refreshHostWorkbooks(): Promise<void> {
        const uris = await vscode.workspace.findFiles(
            '**/*.{xlsm,xlsb,xlam}',
            '{**/node_modules/**,**/.venv/**,**/venv/**}',
        );
        this._hostWorkbooks.clear();
        this._hostPathToId.clear();
        for (const uri of uris) {
            const fsPath = uri.fsPath;
            const id = this._idForPath(fsPath);
            this._hostWorkbooks.set(id, fsPath);
            this._hostPathToId.set(fsPath, id);
        }
        if (this._hostService) {
            try {
                this._hostService.notify(NOTIFY_WORKBOOKS_CHANGED, {});
            } catch (err) {
                this._out.appendLine(`Live Share notify failed: ${err}`);
            }
        }
    }

    private _idForPath(fsPath: string): string {
        // Stable, opaque-ish, URL-safe. Avoids leaking full drive paths.
        return Buffer.from(fsPath, 'utf8').toString('base64url');
    }

    private _resolveHostPath(workbookId: string): string {
        const p = this._hostWorkbooks.get(workbookId);
        if (!p) {
            throw new Error(`XLIDE Live Share: unknown workbook id ${workbookId}`);
        }
        return p;
    }

    private _handleListWorkbooks(): RemoteWorkbookInfo[] {
        const wsRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? '';
        const list: RemoteWorkbookInfo[] = [];
        for (const [id, fsPath] of this._hostWorkbooks) {
            list.push({
                id,
                name: path.basename(fsPath),
                relativeFolder: wsRoot ? path.relative(wsRoot, path.dirname(fsPath)) : '',
            });
        }
        list.sort((a, b) => a.name.localeCompare(b.name));
        return list;
    }

    private async _handleListModules(args: unknown[]): Promise<RemoteModuleInfo[]> {
        const id = String((args[0] as { workbookId?: string })?.workbookId ?? '');
        const filePath = this._resolveHostPath(id);
        return await this._bridge.call<RemoteModuleInfo[]>('listModules', { path: filePath });
    }

    private async _handleListSubs(args: unknown[]): Promise<RemoteSubInfo[]> {
        const a = args[0] as { workbookId?: string; module?: string };
        const id = String(a?.workbookId ?? '');
        const module = String(a?.module ?? '');
        const filePath = this._resolveHostPath(id);
        return await this._bridge.call<RemoteSubInfo[]>('listSubs', {
            path: filePath,
            module,
        });
    }

    private async _handleReadModule(args: unknown[]): Promise<{ source: string }> {
        const a = args[0] as { workbookId?: string; module?: string };
        const id = String(a?.workbookId ?? '');
        const module = String(a?.module ?? '');
        const filePath = this._resolveHostPath(id);
        return await this._bridge.call<{ source: string }>('readModule', {
            path: filePath,
            module,
        });
    }

    private async _handleWriteModule(args: unknown[]): Promise<void> {
        const a = args[0] as { workbookId?: string; module?: string; source?: string };
        const id = String(a?.workbookId ?? '');
        const module = String(a?.module ?? '');
        const source = String(a?.source ?? '');
        const filePath = this._resolveHostPath(id);
        await this._bridge.call('writeModule', { path: filePath, module, source });
        // Tell guests so their open editors pick up changes from other peers
        if (this._hostService) {
            try {
                this._hostService.notify(NOTIFY_FILE_CHANGED, { workbookId: id, module });
            } catch { /* ignore */ }
        }
    }

    // ------------------------------------------------------------------
    // Guest side
    // ------------------------------------------------------------------

    private async _initGuest(): Promise<void> {
        if (!this._api) { return; }
        this._out.appendLine(`[LiveShare] guest: connecting to service '${SERVICE_NAME}'...`);
        // Guests also need the "insiders" feature set for the shared service proxy to work.
        const featureSet = vscode.workspace.getConfiguration('liveshare').get<string>('featureSet');
        if (featureSet !== 'insiders') {
            this._out.appendLine(`[LiveShare] guest: liveshare.featureSet='${featureSet}', expected 'insiders'.`);
            void this._promptFeatureSetFix();
        }
        let proxy: vsls.SharedServiceProxy | null;
        try {
            proxy = await this._api.getSharedService(SERVICE_NAME);
        } catch (err) {
            this._out.appendLine(`[LiveShare] guest: getSharedService failed: ${err}`);
            return;
        }
        if (!proxy) {
            this._out.appendLine('[LiveShare] guest: XLIDE service unavailable (host extension not installed/activated?).');
            return;
        }
        this._guestProxy = proxy;
        this._out.appendLine(`[LiveShare] guest: proxy acquired, isServiceAvailable=${proxy.isServiceAvailable}`);
        proxy.onDidChangeIsServiceAvailable((available) => {
            this._out.appendLine(`[LiveShare] guest: service availability changed -> ${available}`);
            this._onDidChange.fire();
        });
        proxy.onNotify(NOTIFY_WORKBOOKS_CHANGED, () => {
            this._out.appendLine('[LiveShare] guest: host notified workbooks changed');
            this._onDidChange.fire();
        });
        proxy.onNotify(NOTIFY_FILE_CHANGED, (args: object) => {
            const a = args as { workbookId?: string; module?: string };
            if (a.workbookId && a.module) {
                this._onRemoteFileChanged(a.workbookId, a.module);
            }
        });
    }

    // ------------------------------------------------------------------
    // Public API consumed by FileSystem provider and Explorer
    // ------------------------------------------------------------------

    get isGuest(): boolean {
        return this._guestProxy !== null && this._guestProxy.isServiceAvailable;
    }

    /**
     * True whenever the current Live Share session role is Guest, regardless
     * of whether the shared service proxy has connected yet. Used to suppress
     * local workbook discovery on the guest side (the workspace is virtual).
     */
    get isInGuestSession(): boolean {
        return this._api?.session?.role === 2 /* Guest */;
    }

    get isHost(): boolean {
        return this._hostService !== null;
    }

    /** Set by XlideFileSystemProvider to receive remote-change notifications. */
    onRemoteFileChanged: (workbookId: string, moduleName: string) => void = () => { /* default no-op */ };

    private _onRemoteFileChanged(workbookId: string, moduleName: string): void {
        try {
            this.onRemoteFileChanged(workbookId, moduleName);
        } catch (err) {
            this._out.appendLine(`onRemoteFileChanged handler error: ${err}`);
        }
    }

    async guestListWorkbooks(): Promise<RemoteWorkbookInfo[]> {
        const p = this._requireProxy();
        this._out.appendLine('[LiveShare] guest: requesting workbook list from host...');
        const list = await p.request(RPC_LIST_WORKBOOKS, []) as RemoteWorkbookInfo[];
        this._out.appendLine(`[LiveShare] guest: received ${list.length} workbook(s) from host`);
        return list;
    }

    async guestListModules(workbookId: string): Promise<RemoteModuleInfo[]> {
        const p = this._requireProxy();
        return await p.request(RPC_LIST_MODULES, [{ workbookId }]) as RemoteModuleInfo[];
    }

    async guestListSubs(workbookId: string, module: string): Promise<RemoteSubInfo[]> {
        const p = this._requireProxy();
        return await p.request(RPC_LIST_SUBS, [{ workbookId, module }]) as RemoteSubInfo[];
    }

    async guestReadModule(workbookId: string, module: string): Promise<string> {
        const p = this._requireProxy();
        const res = await p.request(RPC_READ_MODULE, [{ workbookId, module }]) as { source: string };
        return res.source;
    }

    async guestWriteModule(workbookId: string, module: string, source: string): Promise<void> {
        const p = this._requireProxy();
        await p.request(RPC_WRITE_MODULE, [{ workbookId, module, source }]);
    }

    private _requireProxy(): vsls.SharedServiceProxy {
        if (!this._guestProxy || !this._guestProxy.isServiceAvailable) {
            throw new Error('XLIDE: Live Share host service is not available.');
        }
        return this._guestProxy;
    }

    private async _promptFeatureSetFix(): Promise<void> {
        const cfg = vscode.workspace.getConfiguration('liveshare');
        const current = cfg.get<string>('featureSet');
        if (current === 'insiders') { return; }
        const ENABLE = 'Enable and Reload';
        const choice = await vscode.window.showWarningMessage(
            'XLIDE Live Share integration requires Live Share\'s "insiders" feature set on the host (a Live Share gating for third-party extensions). Enable it now?',
            ENABLE,
            'Dismiss',
        );
        if (choice !== ENABLE) { return; }
        try {
            await cfg.update('featureSet', 'insiders', vscode.ConfigurationTarget.Global);
            await vscode.commands.executeCommand('workbench.action.reloadWindow');
        } catch (err) {
            this._out.appendLine(`[LiveShare] failed to update liveshare.featureSet: ${err}`);
        }
    }

    dispose(): void {
        if (this._hostService && this._api) {
            void this._api.unshareService(SERVICE_NAME).catch(() => { /* ignore */ });
        }
        for (const d of this._disposables) { d.dispose(); }
        this._disposables = [];
        this._onDidChange.dispose();
    }
}
