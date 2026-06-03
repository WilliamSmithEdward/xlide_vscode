import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import {
    isWorkbookSettingsError,
    readWorkbookSettings,
    settingsPathForWorkbook,
} from './workbookSettings';
import { registerXlideCommand } from './xlideCommandRegistration';
import { decodeModuleUri, sameWorkbookPath, XLIDE_SCHEME } from './xlideFileSystem';
import {
    buildXlideSidebarModel,
    type XlideSidebarActiveWorkbook,
    type XlideSidebarCommand,
    type XlideSidebarNode,
    type XlideSidebarSetupStatus,
    type XlideSidebarWorkbookChoice,
} from './xlideSidebarModel';

interface XlideSidebarOptions {
    setupStatus?: () => XlideSidebarSetupStatus;
    workspaceState?: vscode.Memento;
}

interface XlideSidebarRegistration {
    disposables: vscode.Disposable[];
    refresh(): void;
}

class XlideSidebarProvider implements vscode.WebviewViewProvider {
    private _view: vscode.WebviewView | undefined;
    private _refreshVersion = 0;
    private _selectedWorkbookPath: string | undefined;

    constructor(private readonly _options: XlideSidebarOptions = {}) {
        this._selectedWorkbookPath = _options.workspaceState?.get<string>(SELECTED_WORKBOOK_KEY);
    }

    refresh(): void {
        void this._render();
    }

    resolveWebviewView(view: vscode.WebviewView): void {
        this._view = view;
        view.webview.options = { enableScripts: true };
        view.webview.onDidReceiveMessage((message: unknown) => {
            void this._handleMessage(message);
        });
        this.refresh();
    }

    private async _model(): Promise<XlideSidebarNode[]> {
        const workbooks = await workbookFiles();
        const selectedWorkbookPath = await this._validSelectedWorkbookPath(workbooks);
        const activeWorkbook = await activeWorkbookContext(workbooks, selectedWorkbookPath);
        return buildXlideSidebarModel({
            workbookChoices: workbookChoices(workbooks),
            activeWorkbook,
            setupStatus: this._options.setupStatus?.(),
        });
    }

    private async _render(): Promise<void> {
        if (!this._view) {
            return;
        }
        const version = ++this._refreshVersion;
        const model = await this._model();
        if (version !== this._refreshVersion || !this._view) {
            return;
        }
        this._view.webview.html = renderXlideSidebarHtml(model);
    }

    private async _handleMessage(message: unknown): Promise<void> {
        if (!message || typeof message !== 'object') {
            return;
        }
        const payload = message as { type?: unknown; command?: unknown; arguments?: unknown; filePath?: unknown };
        if (payload.type === 'selectWorkbook') {
            await this._selectWorkbook(typeof payload.filePath === 'string' ? payload.filePath : undefined);
            return;
        }
        if (payload.type !== 'runCommand' || typeof payload.command !== 'string') {
            return;
        }
        const args = Array.isArray(payload.arguments) ? payload.arguments : [];
        await vscode.commands.executeCommand(payload.command, ...args);
    }

    private async _selectWorkbook(filePath: string | undefined): Promise<void> {
        const workbooks = await workbookFiles();
        const workbook = filePath
            ? findWorkbook(workbooks, filePath)
            : undefined;
        if (filePath && !workbook) {
            vscode.window.showWarningMessage('XLIDE: That workbook is no longer available in this workspace.');
        }
        await this._setSelectedWorkbookPath(workbook?.fsPath);
        this.refresh();
    }

    private async _validSelectedWorkbookPath(workbooks: readonly vscode.Uri[]): Promise<string | undefined> {
        if (!this._selectedWorkbookPath) {
            return undefined;
        }
        const workbook = findWorkbook(workbooks, this._selectedWorkbookPath);
        if (workbook) {
            return workbook.fsPath;
        }
        await this._setSelectedWorkbookPath(undefined);
        return undefined;
    }

    private async _setSelectedWorkbookPath(filePath: string | undefined): Promise<void> {
        this._selectedWorkbookPath = filePath;
        await this._options.workspaceState?.update(SELECTED_WORKBOOK_KEY, filePath);
    }
}

const SELECTED_WORKBOOK_KEY = 'xlide.sidebar.selectedWorkbookPath';
const XLIDE_SPONSOR_URL = 'https://github.com/sponsors/WilliamSmithEdward';
const XLIDE_PAYPAL_DONATE_URL =
    'https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0' +
    '&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F' +
    '&currency_code=USD';
const XLIDE_CASH_APP_DONATE_URL = 'https://cash.app/$williamesmithjcil';

function registerXlideSidebar(options: XlideSidebarOptions = {}): XlideSidebarRegistration {
    const provider = new XlideSidebarProvider(options);
    const view = vscode.window.registerWebviewViewProvider('xlide.sidebar', provider, {
        webviewOptions: { retainContextWhenHidden: true },
    });

    const disposables = [
        view,
        vscode.workspace.onDidChangeConfiguration((event) => {
            if (event.affectsConfiguration('xlide')) {
                provider.refresh();
            }
        }),
        vscode.workspace.onDidChangeWorkspaceFolders(() => provider.refresh()),
        vscode.window.onDidChangeActiveTextEditor(() => provider.refresh()),
        vscode.commands.registerCommand('xlide.openWorkbookSettings', async (settingsPath?: string) => {
            if (!settingsPath) {
                vscode.window.showWarningMessage('XLIDE: No workbook settings file is available.');
                return;
            }
            try {
                if (!fs.existsSync(settingsPath)) {
                    await fs.promises.writeFile(settingsPath, '{}\n', { encoding: 'utf8', flag: 'wx' });
                }
                const document = await vscode.workspace.openTextDocument(vscode.Uri.file(settingsPath));
                await vscode.window.showTextDocument(document, { preview: false });
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                vscode.window.showErrorMessage(`XLIDE: Could not open workbook settings: ${message}`);
            }
        }),
        registerExternalLinkCommand('xlide.openSponsorLink', XLIDE_SPONSOR_URL),
        registerExternalLinkCommand('xlide.openPayPalDonateLink', XLIDE_PAYPAL_DONATE_URL),
        registerExternalLinkCommand('xlide.openCashAppDonateLink', XLIDE_CASH_APP_DONATE_URL),
        (() => {
            let timer: ReturnType<typeof setTimeout> | undefined;
            const refresh = () => {
                if (timer !== undefined) {
                    clearTimeout(timer);
                }
                timer = setTimeout(() => {
                    timer = undefined;
                    provider.refresh();
                }, 200);
            };
            const watcher = vscode.workspace.createFileSystemWatcher('**/*.{xlsm,xlsb,xlam}');
            watcher.onDidCreate(refresh);
            watcher.onDidDelete(refresh);
            return watcher;
        })(),
        (() => {
            let timer: ReturnType<typeof setTimeout> | undefined;
            const refresh = () => {
                if (timer !== undefined) {
                    clearTimeout(timer);
                }
                timer = setTimeout(() => {
                    timer = undefined;
                    provider.refresh();
                }, 200);
            };
            const watcher = vscode.workspace.createFileSystemWatcher('**/*.xlide_settings.json');
            watcher.onDidCreate(refresh);
            watcher.onDidChange(refresh);
            watcher.onDidDelete(refresh);
            return watcher;
        })(),
    ];

    return {
        disposables,
        refresh: () => provider.refresh(),
    };
}

function registerExternalLinkCommand(command: string, url: string): vscode.Disposable {
    return registerXlideCommand(command, () => vscode.env.openExternal(vscode.Uri.parse(url)));
}

async function workbookFileCount(): Promise<number> {
    return (await workbookFiles()).length;
}

async function workbookFiles(): Promise<vscode.Uri[]> {
    const uris = await vscode.workspace.findFiles(
        '**/*.{xlsm,xlsb,xlam}',
        '{**/node_modules/**,**/.venv/**,**/venv/**}',
    );
    return uris
        .filter((uri) => uri.scheme === 'file' && !path.basename(uri.fsPath).startsWith('~$'))
        .sort((left, right) => left.fsPath.localeCompare(right.fsPath));
}

async function activeWorkbookContext(
    workbooks: readonly vscode.Uri[],
    selectedWorkbookPath?: string,
): Promise<XlideSidebarActiveWorkbook | undefined> {
    if (selectedWorkbookPath && findWorkbook(workbooks, selectedWorkbookPath)) {
        return sidebarWorkbookForPath(selectedWorkbookPath, 'sidebarSelection');
    }
    const activeFromEditor = activeWorkbookPathFromEditor();
    if (activeFromEditor) {
        return sidebarWorkbookForPath(activeFromEditor, 'activeEditor');
    }
    if (workbooks.length === 1) {
        return sidebarWorkbookForPath(workbooks[0].fsPath, 'singleWorkbook');
    }
    return undefined;
}

function workbookChoices(workbooks: readonly vscode.Uri[]): XlideSidebarWorkbookChoice[] {
    return workbooks.map((uri) => ({
        label: vscode.workspace.asRelativePath(uri, false),
        filePath: uri.fsPath,
        description: uri.fsPath,
    }));
}

function findWorkbook(workbooks: readonly vscode.Uri[], filePath: string): vscode.Uri | undefined {
    return workbooks.find((uri) => sameWorkbookPath(uri.fsPath, filePath));
}

function activeWorkbookPathFromEditor(): string | undefined {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
        return undefined;
    }
    const uri = editor.document.uri;
    if (uri.scheme !== XLIDE_SCHEME || uri.authority) {
        return undefined;
    }
    return decodeModuleUri(uri).xlsmPath;
}

async function sidebarWorkbookForPath(
    workbookPath: string,
    selectionSource: XlideSidebarActiveWorkbook['selectionSource'],
): Promise<XlideSidebarActiveWorkbook> {
    const settingsPath = settingsPathForWorkbook(workbookPath);
    const base = {
        label: path.basename(workbookPath),
        filePath: workbookPath,
        settingsPath,
        selectionSource,
    };
    try {
        const exists = fs.existsSync(settingsPath);
        await readWorkbookSettings(workbookPath);
        return {
            ...base,
            settingsState: exists ? 'valid' : 'missing',
        };
    } catch (err) {
        return {
            ...base,
            settingsState: 'invalid',
            settingsMessage: isWorkbookSettingsError(err)
                ? err.message
                : `Unable to read workbook settings: ${err instanceof Error ? err.message : String(err)}`,
        };
    }
}

function renderXlideSidebarHtml(sections: readonly XlideSidebarNode[]): string {
    const nonce = nonceString();
    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE</title>
    <style nonce="${nonce}">
        :root {
            color-scheme: light dark;
        }
        * {
            box-sizing: border-box;
        }
        body {
            margin: 0;
            padding: 12px;
            color: var(--vscode-foreground);
            background: var(--vscode-sideBar-background);
            font-family: var(--vscode-font-family);
            font-size: var(--vscode-font-size);
            line-height: 1.35;
        }
        .shell {
            display: flex;
            flex-direction: column;
            gap: 12px;
            min-width: 0;
        }
        .section {
            border: 1px solid var(--vscode-sideBarSectionHeader-border, var(--vscode-panel-border));
            border-radius: 6px;
            background: var(--vscode-sideBar-background);
            overflow: hidden;
        }
        .sectionHeader {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
            padding: 8px 10px;
            font-weight: 700;
            color: var(--vscode-sideBarTitle-foreground, var(--vscode-foreground));
            background: var(--vscode-sideBarSectionHeader-background);
            border-bottom: 1px solid var(--vscode-sideBarSectionHeader-border, var(--vscode-panel-border));
        }
        .sectionBody {
            display: flex;
            flex-direction: column;
        }
        .row {
            display: grid;
            grid-template-columns: 14px minmax(0, 1fr) auto;
            gap: 8px;
            align-items: center;
            padding: 9px 10px;
            min-width: 0;
        }
        .noDotRow {
            grid-template-columns: minmax(0, 1fr) auto;
        }
        .buttonOnlyRow {
            grid-template-columns: 1fr;
        }
        .buttonOnlyRow button {
            width: 100%;
        }
        .row + .row {
            border-top: 1px solid var(--vscode-panel-border);
        }
        .dot {
            width: 9px;
            height: 9px;
            border-radius: 50%;
            background: var(--vscode-descriptionForeground);
            box-shadow: 0 0 0 2px color-mix(in srgb, currentColor 10%, transparent);
        }
        .pass {
            background: var(--vscode-testing-iconPassed);
        }
        .warn {
            background: var(--vscode-testing-iconQueued, var(--vscode-editorWarning-foreground));
        }
        .fail {
            background: var(--vscode-testing-iconFailed);
        }
        .unknown {
            background: var(--vscode-descriptionForeground);
        }
        .rowText {
            min-width: 0;
        }
        .label {
            font-weight: 600;
            color: var(--vscode-foreground);
            overflow-wrap: anywhere;
            white-space: normal;
        }
        .description {
            margin-top: 1px;
            color: var(--vscode-descriptionForeground);
            overflow-wrap: anywhere;
            white-space: normal;
        }
        button {
            border: 1px solid var(--vscode-button-border, transparent);
            border-radius: 4px;
            padding: 4px 8px;
            color: var(--vscode-button-foreground);
            background: var(--vscode-button-background);
            font: inherit;
            cursor: pointer;
            min-height: 24px;
        }
        button:hover {
            background: var(--vscode-button-hoverBackground);
        }
        button:disabled {
            opacity: 0.55;
            cursor: default;
        }
        button:disabled:hover {
            background: var(--vscode-button-secondaryBackground);
        }
        button.secondary {
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
        }
        button.secondary:hover {
            background: var(--vscode-button-secondaryHoverBackground);
        }
        .actionGrid {
            display: grid;
            grid-template-columns: 1fr;
            gap: 8px;
            padding: 10px;
        }
        .actionCard {
            width: 100%;
            text-align: left;
            padding: 8px 9px;
            border-color: var(--vscode-button-secondaryBackground);
        }
        .actionCard .label {
            color: inherit;
        }
        .actionCard .description {
            color: color-mix(in srgb, currentColor 74%, transparent);
        }
        .selectRow {
            grid-template-columns: 14px minmax(0, 1fr);
        }
        .selectRow.noDotRow {
            grid-template-columns: minmax(0, 1fr);
        }
        select {
            width: 100%;
            min-width: 0;
            margin-top: 7px;
            border: 1px solid var(--vscode-dropdown-border);
            border-radius: 4px;
            padding: 4px 7px;
            color: var(--vscode-dropdown-foreground);
            background: var(--vscode-dropdown-background);
            font: inherit;
        }
        select:focus {
            outline: 1px solid var(--vscode-focusBorder);
            outline-offset: 1px;
        }
        .empty {
            padding: 10px;
            color: var(--vscode-descriptionForeground);
        }
    </style>
</head>
<body>
    <main class="shell">
        ${sections.map((section) => renderSection(section)).join('')}
    </main>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        let ctrlMode = false;
        function setCtrlMode(next) {
            if (ctrlMode === next) {
                return;
            }
            ctrlMode = next;
            document.querySelectorAll('[data-ctrl-command-label]').forEach((button) => {
                button.textContent = ctrlMode ? button.dataset.ctrlCommandLabel : button.dataset.commandLabel;
                button.title = ctrlMode ? button.dataset.ctrlCommandTitle : button.dataset.commandTitle;
            });
        }
        document.addEventListener('keydown', (event) => {
            if (event.key === 'Control') {
                setCtrlMode(true);
            }
        });
        document.addEventListener('keyup', (event) => {
            if (event.key === 'Control') {
                setCtrlMode(false);
            }
        });
        document.addEventListener('mousemove', (event) => setCtrlMode(event.ctrlKey === true));
        window.addEventListener('blur', () => setCtrlMode(false));
        document.addEventListener('contextmenu', (event) => {
            event.preventDefault();
        });
        document.addEventListener('click', (event) => {
            const button = event.target.closest('[data-command]');
            if (!button) {
                return;
            }
            const payload = JSON.parse(button.dataset.command);
            const useCtrlCommand = event.ctrlKey && payload.ctrlCommand;
            vscode.postMessage({
                type: 'runCommand',
                command: useCtrlCommand ? payload.ctrlCommand : payload.command,
                arguments: useCtrlCommand ? payload.ctrlArguments || [] : payload.arguments || []
            });
        });
        document.addEventListener('change', (event) => {
            const select = event.target.closest('select[data-select-id]');
            if (!select) {
                return;
            }
            if (select.dataset.selectId === 'project.targetWorkbook') {
                vscode.postMessage({
                    type: 'selectWorkbook',
                    filePath: select.value || undefined
                });
            }
        });
    </script>
</body>
</html>`;
}

function renderSection(section: XlideSidebarNode): string {
    const children = section.children ?? [];
    const isActionSection = section.id === 'workbookActions' ||
        section.id === 'settings' ||
        section.id === 'support' ||
        section.id === 'donate';
    return `<section class="section" aria-label="${escapeAttr(section.label)}">
        <div class="sectionHeader">${escapeHtml(section.label)}</div>
        <div class="${isActionSection ? 'actionGrid' : 'sectionBody'}">
            ${children.length > 0
        ? children.map((node) => isActionSection && node.kind === 'action'
            ? renderActionNode(node)
            : renderSidebarNode(node, section.id)).join('')
        : '<div class="empty">No items</div>'}
        </div>
    </section>`;
}

function renderSidebarNode(node: XlideSidebarNode, sectionId: string): string {
    if (node.kind === 'select') {
        return renderSelectNode(node, sectionId);
    }
    if (node.kind === 'action') {
        return renderButtonOnlyRow(node);
    }
    return renderRowNode(node, sectionId);
}

function renderActionNode(node: XlideSidebarNode): string {
    if (!node.command && !node.disabled) {
        return renderRowNode(node, '');
    }
    const command = node.command && !node.disabled
        ? ` data-command="${commandAttr(node.command)}"`
        : ' disabled';
    return `<button class="actionCard secondary"${command} title="${escapeAttr(node.tooltip ?? node.label)}">
        <div class="label">${escapeHtml(node.label)}</div>
        ${node.description ? `<div class="description">${escapeHtml(node.description)}</div>` : ''}
    </button>`;
}

function renderButtonOnlyRow(node: XlideSidebarNode): string {
    if (!node.command && !node.disabled) {
        return renderRowNode(node, '');
    }
    const command = node.command && !node.disabled
        ? ` data-command="${commandAttr(node.command)}"`
        : ' disabled';
    return `<div class="row noDotRow buttonOnlyRow" title="${escapeAttr(node.tooltip ?? node.label)}">
        <button class="secondary"${command}>${escapeHtml(node.label)}</button>
    </div>`;
}

function renderRowNode(node: XlideSidebarNode, sectionId: string): string {
    const status = node.status ?? 'unknown';
    const showDot = sectionId === 'setup' && node.kind === 'status';
    const rowClass = showDot ? 'row' : 'row noDotRow';
    const commandTitle = node.command?.tooltip ?? node.command?.title ?? node.label;
    const command = node.command
        ? `<button class="secondary"${node.disabled ? ' disabled' : ` data-command="${commandAttr(node.command)}"`}${commandButtonStateAttrs(node.command, commandTitle)}>${escapeHtml(node.command.title)}</button>`
        : '';
    return `<div class="${rowClass}" title="${escapeAttr(node.tooltip ?? node.label)}">
        ${showDot ? `<span class="dot ${escapeAttr(status)}" aria-hidden="true"></span>` : ''}
        <div class="rowText">
            <div class="label">${escapeHtml(node.label)}</div>
            ${node.description ? `<div class="description">${escapeHtml(node.description)}</div>` : ''}
        </div>
        ${command}
    </div>`;
}

function commandButtonStateAttrs(command: XlideSidebarCommand, title: string): string {
    const base = ` title="${escapeAttr(title)}" data-command-label="${escapeAttr(command.title)}" data-command-title="${escapeAttr(title)}"`;
    if (!command.ctrlCommand || !command.ctrlTitle) {
        return base;
    }
    const ctrlTitle = command.ctrlTooltip ?? command.ctrlTitle;
    return `${base} data-ctrl-command-label="${escapeAttr(command.ctrlTitle)}" data-ctrl-command-title="${escapeAttr(ctrlTitle)}"`;
}

function renderSelectNode(node: XlideSidebarNode, sectionId: string): string {
    const status = node.status ?? 'unknown';
    const options = node.options ?? [];
    const showDot = sectionId === 'setup';
    const rowClass = showDot ? 'row selectRow' : 'row selectRow noDotRow';
    return `<div class="${rowClass}" title="${escapeAttr(node.tooltip ?? node.label)}">
        ${showDot ? `<span class="dot ${escapeAttr(status)}" aria-hidden="true"></span>` : ''}
        <div class="rowText">
            <div class="label">${escapeHtml(node.label)}</div>
            ${node.description ? `<div class="description">${escapeHtml(node.description)}</div>` : ''}
            <select data-select-id="${escapeAttr(node.id)}" aria-label="${escapeAttr(node.label)}">
                ${options.map((option) => renderSelectOption(option, node.value ?? '')).join('')}
            </select>
        </div>
    </div>`;
}

function renderSelectOption(option: { label: string; value: string }, selectedValue: string): string {
    const selected = option.value === selectedValue ? ' selected' : '';
    return `<option value="${escapeAttr(option.value)}"${selected}>${escapeHtml(option.label)}</option>`;
}

function commandAttr(command: XlideSidebarCommand): string {
    return escapeAttr(JSON.stringify(command));
}

function escapeHtml(value: string): string {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function escapeAttr(value: string): string {
    return escapeHtml(value).replace(/"/g, '&quot;');
}

function nonceString(): string {
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    for (let i = 0; i < 24; i++) {
        result += alphabet[Math.floor(Math.random() * alphabet.length)];
    }
    return result;
}

export {
    XlideSidebarProvider,
    type XlideSidebarRegistration,
    registerXlideSidebar,
    workbookFileCount,
    workbookFiles,
};
