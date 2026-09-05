import * as fs from 'fs';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { MACRO_CONTAINER_GLOB } from './macroContainerUi';
import * as path from 'path';
import * as vscode from 'vscode';
import {
    isProjectSettingsError,
    readProjectSettings,
    settingsPathForProject,
} from './projectSettings';
import { registerXlideCommand } from './xlideCommandRegistration';
import { activeLocalVbaEditor, decodeModuleUri, sameProjectPath, XLIDE_SCHEME } from './xlideFileSystem';
import {
    buildXlideSidebarModel,
    isSponsorUrl,
    type XlideSidebarActiveProject,
    type XlideSidebarCommand,
    type XlideSidebarNode,
    type XlideSidebarProjectChoice,
} from './xlideSidebarModel';
import { measurePerformance, startPerformanceTrace } from './performanceTrace';
import { escapeAttr, escapeHtml, randomNonce } from './webview/html';
import { webviewHeadHtml } from './webview/page';
import { WEBVIEW_BODY_CSS, xlideAccentPaletteCss } from './webview/styles';
import { errorMessage } from './util/errors';
import { fileExists } from './util/fs';
import { debounce } from './util/debounce';

interface XlideSidebarOptions {
    workspaceState?: vscode.Memento;
    /** Fired whenever the sidebar view is (re)shown, e.g. to lazy-start the backend. */
    onDidBecomeVisible?: () => void;
}

interface XlideSidebarRegistration {
    disposables: vscode.Disposable[];
    refresh(): void;
}

class XlideSidebarProvider implements vscode.WebviewViewProvider {
    private _view: vscode.WebviewView | undefined;
    private _refreshVersion = 0;
    private _selectedProjectPath: string | undefined;
    private _lastSelectionSource: XlideSidebarActiveProject['selectionSource'] | undefined;
    private _lastRenderedModelJson: string | undefined;
    private _projectFilesPromise: Promise<vscode.Uri[]> | undefined;

    constructor(private readonly _options: XlideSidebarOptions = {}) {
        this._selectedProjectPath = _options.workspaceState?.get<string>(SELECTED_WORKBOOK_KEY);
    }

    refresh(): void {
        void this._render();
    }

    /** Drops the cached workspace project scan; the next render re-globs. */
    invalidateProjectFiles(): void {
        this._projectFilesPromise = undefined;
    }

    /**
     * An editor change only affects the model when an XLIDE editor became
     * active or the displayed project was derived from the active editor.
     */
    shouldRefreshForActiveEditorChange(editor: vscode.TextEditor | undefined): boolean {
        return editor?.document.uri.scheme === XLIDE_SCHEME ||
            this._lastSelectionSource === 'activeEditor';
    }

    resolveWebviewView(view: vscode.WebviewView): void {
        this._view = view;
        this._lastRenderedModelJson = undefined;
        view.webview.options = { enableScripts: true };
        view.webview.onDidReceiveMessage((message: unknown) => {
            void this._handleMessage(message);
        });
        this._options.onDidBecomeVisible?.();
        this.refresh();
    }

    private _projectFiles(): Promise<vscode.Uri[]> {
        this._projectFilesPromise ??= projectFiles();
        return this._projectFilesPromise;
    }

    private async _model(): Promise<XlideSidebarNode[]> {
        const projects = await this._projectFiles();
        const selectedProjectPath = await this._validSelectedProjectPath(projects);
        const activeProject = await activeProjectContext(projects, selectedProjectPath);
        this._lastSelectionSource = activeProject?.selectionSource;
        return buildXlideSidebarModel({
            projectChoices: projectChoices(projects),
            activeProject,
        });
    }

    private async _render(): Promise<void> {
        const trace = startPerformanceTrace('sidebar.render');
        if (!this._view) {
            trace.end('ok', 'hidden');
            return;
        }
        const version = ++this._refreshVersion;
        try {
            const model = await this._model();
            if (version !== this._refreshVersion || !this._view) {
                trace.end('superseded');
                return;
            }
            const modelJson = JSON.stringify(model);
            if (modelJson === this._lastRenderedModelJson) {
                trace.end('ok', 'unchanged');
                return;
            }
            this._lastRenderedModelJson = modelJson;
            this._view.webview.html = renderXlideSidebarHtml(model);
            trace.end('ok', `${model.length} nodes`);
        } catch (err) {
            trace.end('failed');
            throw err;
        }
    }

    private async _handleMessage(message: unknown): Promise<void> {
        if (!message || typeof message !== 'object') {
            return;
        }
        const payload = message as {
            type?: unknown;
            command?: unknown;
            arguments?: unknown;
            filePath?: unknown;
            url?: unknown;
        };
        if (payload.type === 'selectProject') {
            await this._selectProject(typeof payload.filePath === 'string' ? payload.filePath : undefined);
            return;
        }
        // The sponsor rows open or copy an address. The webview names it, but
        // only an address from the model's own list is honored (xlideSidebarModel.ts).
        if (payload.type === 'openSponsorUrl' || payload.type === 'copySponsorUrl') {
            if (!isSponsorUrl(payload.url)) {
                return;
            }
            if (payload.type === 'openSponsorUrl') {
                await vscode.env.openExternal(vscode.Uri.parse(payload.url));
            } else {
                await vscode.env.clipboard.writeText(payload.url);
            }
            return;
        }
        if (payload.type !== 'runCommand' || typeof payload.command !== 'string') {
            return;
        }
        // Trust-boundary hardening: the webview is a distinct security context, so
        // never forward an arbitrary command id from a postMessage payload. Only
        // XLIDE's own commands plus the settings command the sidebar model emits
        // (xlideSidebarModel.ts) are allowed.
        if (!payload.command.startsWith('xlide.') && payload.command !== 'workbench.action.openSettings') {
            return;
        }
        const args = Array.isArray(payload.arguments) ? payload.arguments : [];
        await vscode.commands.executeCommand(payload.command, ...args);
    }

    private async _selectProject(filePath: string | undefined): Promise<void> {
        const projects = await this._projectFiles();
        const project = filePath
            ? findProject(projects, filePath)
            : undefined;
        if (filePath && !project) {
            vscode.window.showWarningMessage('XLIDE: That file is no longer available in this workspace.');
        }
        await this._setSelectedProjectPath(project?.fsPath);
        this.refresh();
    }

    private async _validSelectedProjectPath(projects: readonly vscode.Uri[]): Promise<string | undefined> {
        if (!this._selectedProjectPath) {
            return undefined;
        }
        const project = findProject(projects, this._selectedProjectPath);
        if (project) {
            return project.fsPath;
        }
        await this._setSelectedProjectPath(undefined);
        return undefined;
    }

    private async _setSelectedProjectPath(filePath: string | undefined): Promise<void> {
        this._selectedProjectPath = filePath;
        await this._options.workspaceState?.update(SELECTED_WORKBOOK_KEY, filePath);
    }
}

const SELECTED_WORKBOOK_KEY = 'xlide.sidebar.selectedProjectPath';

function registerXlideSidebar(options: XlideSidebarOptions = {}): XlideSidebarRegistration {
    const provider = new XlideSidebarProvider(options);
    // No retainContextWhenHidden: webview.html always carries the latest model,
    // so a re-shown sidebar rebuilds for free.
    const view = vscode.window.registerWebviewViewProvider('xlide.sidebar', provider);
    const scheduleRefresh = debounce(() => provider.refresh(), 200);
    const projectFilesChanged = () => {
        provider.invalidateProjectFiles();
        scheduleRefresh();
    };

    const disposables = [
        view,
        scheduleRefresh,
        vscode.workspace.onDidChangeConfiguration((event) => {
            if (event.affectsConfiguration('xlide')) {
                scheduleRefresh();
            }
        }),
        vscode.workspace.onDidChangeWorkspaceFolders(projectFilesChanged),
        vscode.window.onDidChangeActiveTextEditor((editor) => {
            if (provider.shouldRefreshForActiveEditorChange(editor)) {
                scheduleRefresh();
            }
        }),
        registerXlideCommand('xlide.openProjectSettings', async (settingsPath?: string) => {
            if (!settingsPath) {
                vscode.window.showWarningMessage('XLIDE: No settings file is available for this file.');
                return;
            }
            try {
                if (!(await fileExists(settingsPath))) {
                    try {
                        await fs.promises.writeFile(settingsPath, '{}\n', { encoding: 'utf8', flag: 'wx' });
                    } catch (err) {
                        if (!isFileAlreadyExistsError(err)) {
                            throw err;
                        }
                    }
                }
                const document = await vscode.workspace.openTextDocument(vscode.Uri.file(settingsPath));
                await vscode.window.showTextDocument(document, { preview: false });
            } catch (err) {
                const message = errorMessage(err);
                vscode.window.showErrorMessage(`XLIDE: Could not open file settings: ${message}`);
            }
        }),
        (() => {
            // The same glob discovery uses, so a project the tree lists is one the
            // sidebar notices arriving and leaving.
            const watcher = vscode.workspace.createFileSystemWatcher(MACRO_CONTAINER_GLOB);
            watcher.onDidCreate(projectFilesChanged);
            watcher.onDidDelete(projectFilesChanged);
            return watcher;
        })(),
        (() => {
            const watcher = vscode.workspace.createFileSystemWatcher('**/*.xlide_settings.json');
            watcher.onDidCreate(scheduleRefresh);
            watcher.onDidChange(scheduleRefresh);
            watcher.onDidDelete(scheduleRefresh);
            return watcher;
        })(),
    ];

    return {
        disposables,
        refresh: () => provider.refresh(),
    };
}

async function projectFiles(): Promise<vscode.Uri[]> {
    return measurePerformance('sidebar.projectFiles', undefined, () => findMacroContainerFiles());
}

async function activeProjectContext(
    projects: readonly vscode.Uri[],
    selectedProjectPath?: string,
): Promise<XlideSidebarActiveProject | undefined> {
    if (selectedProjectPath && findProject(projects, selectedProjectPath)) {
        return sidebarProjectForPath(selectedProjectPath, 'sidebarSelection');
    }
    const activeFromEditor = activeProjectPathFromEditor();
    if (activeFromEditor) {
        return sidebarProjectForPath(activeFromEditor, 'activeEditor');
    }
    if (projects.length === 1) {
        return sidebarProjectForPath(projects[0].fsPath, 'singleProject');
    }
    return undefined;
}

function projectChoices(projects: readonly vscode.Uri[]): XlideSidebarProjectChoice[] {
    return projects.map((uri) => ({
        label: vscode.workspace.asRelativePath(uri, false),
        filePath: uri.fsPath,
        description: uri.fsPath,
    }));
}

function findProject(projects: readonly vscode.Uri[], filePath: string): vscode.Uri | undefined {
    return projects.find((uri) => sameProjectPath(uri.fsPath, filePath));
}

function activeProjectPathFromEditor(): string | undefined {
    const editor = activeLocalVbaEditor();
    return editor ? decodeModuleUri(editor.document.uri).projectPath : undefined;
}

async function sidebarProjectForPath(
    projectPath: string,
    selectionSource: XlideSidebarActiveProject['selectionSource'],
): Promise<XlideSidebarActiveProject> {
    const settingsPath = settingsPathForProject(projectPath);
    const base = {
        label: path.basename(projectPath),
        filePath: projectPath,
        settingsPath,
        selectionSource,
    };
    try {
        const exists = await fileExists(settingsPath);
        await readProjectSettings(projectPath);
        return {
            ...base,
            settingsState: exists ? 'valid' : 'missing',
        };
    } catch (err) {
        return {
            ...base,
            settingsState: 'invalid',
            settingsMessage: isProjectSettingsError(err)
                ? err.message
                : `Unable to read project settings: ${errorMessage(err)}`,
        };
    }
}


function isFileAlreadyExistsError(err: unknown): boolean {
    return typeof err === 'object' &&
        err !== null &&
        'code' in err &&
        (err as { code?: unknown }).code === 'EEXIST';
}

function renderXlideSidebarHtml(sections: readonly XlideSidebarNode[]): string {
    const nonce = randomNonce();
    return `<!DOCTYPE html>
<html lang="en">
<head>
    ${webviewHeadHtml(nonce, 'XLIDE')}
    <style nonce="${nonce}">
        :root {
            color-scheme: light dark;
            ${xlideAccentPaletteCss({
                surface: 'var(--vscode-sideBar-background)',
                accentBorder: 'color-mix(in srgb, var(--xlide-accent-blue) 72%, var(--vscode-dropdown-border))',
            })}
        }
        * {
            box-sizing: border-box;
        }
        ${WEBVIEW_BODY_CSS}
        body {
            padding: 12px;
            background: var(--vscode-sideBar-background);
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
        .section.hasCustomSelect {
            overflow: visible;
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
        .customSelect {
            position: relative;
            margin-top: 7px;
        }
        .selectButton {
            width: 100%;
            min-height: 30px;
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            gap: 8px;
            align-items: center;
            border-color: var(--xlide-accent-border);
            color: var(--vscode-dropdown-foreground);
            background: var(--vscode-dropdown-background);
            text-align: left;
        }
        .selectButton:hover,
        .selectButton[aria-expanded="true"] {
            background: color-mix(in srgb, var(--xlide-accent-blue) 18%, var(--vscode-dropdown-background));
        }
        .selectButton:focus {
            outline: 1px solid var(--xlide-accent-border);
            outline-offset: 1px;
        }
        .selectButtonLabel {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .selectChevron {
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
        }
        .selectMenu {
            position: absolute;
            z-index: 20;
            inset-inline: 0;
            top: calc(100% + 3px);
            max-height: 180px;
            overflow: auto;
            border: 1px solid var(--xlide-accent-border);
            border-radius: 4px;
            padding: 2px;
            background: var(--vscode-dropdown-background);
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.32);
        }
        .selectMenu[hidden] {
            display: none;
        }
        .selectOption {
            width: 100%;
            min-height: 28px;
            display: block;
            border: 0;
            border-radius: 3px;
            padding: 5px 7px;
            color: var(--vscode-dropdown-foreground);
            background: transparent;
            text-align: left;
        }
        .selectOption:hover,
        .selectOption:focus {
            outline: none;
            background: var(--vscode-list-hoverBackground);
        }
        .selectOption[aria-selected="true"] {
            color: var(--vscode-button-foreground);
            background: var(--xlide-accent-background);
        }
        .selectOption[aria-selected="true"]:hover,
        .selectOption[aria-selected="true"]:focus {
            background: var(--xlide-accent-hover-background);
        }
        .empty {
            padding: 10px;
            color: var(--vscode-descriptionForeground);
        }
        /* Subtle on purpose: a small quiet control under the workflow, not a card. */
        .sponsorFooter {
            display: flex;
            justify-content: center;
            padding: 2px 0 4px;
        }
        .sponsorToggle {
            border: 0;
            background: transparent;
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
            padding: 4px 10px;
            min-height: 0;
        }
        .sponsorToggle:hover {
            color: var(--vscode-foreground);
            background: var(--vscode-toolbar-hoverBackground, rgba(128, 128, 128, 0.2));
        }
        .sponsorBackdrop {
            position: fixed;
            inset: 0;
            z-index: 50;
            display: flex;
            align-items: flex-start;
            justify-content: center;
            padding: 12vh 12px 12px;
            background: rgba(0, 0, 0, 0.35);
        }
        .sponsorBackdrop[hidden] {
            display: none;
        }
        .sponsorCard {
            width: 460px;
            max-width: 100%;
            background: var(--vscode-editorWidget-background, var(--vscode-sideBar-background));
            border: 1px solid var(--vscode-editorWidget-border, var(--vscode-panel-border));
            border-radius: 6px;
            box-shadow: 0 8px 28px rgba(0, 0, 0, 0.45);
            padding-bottom: 4px;
        }
        .sponsorHead {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 12px 14px 8px;
        }
        .sponsorTitle {
            font-size: 15px;
            font-weight: 600;
        }
        .sponsorClose {
            width: 24px;
            height: 24px;
            min-height: 0;
            padding: 0;
            border: 0;
            background: transparent;
            color: inherit;
            font-size: 16px;
            line-height: 1;
            opacity: 0.75;
        }
        .sponsorClose:hover {
            opacity: 1;
            background: var(--vscode-toolbar-hoverBackground, rgba(128, 128, 128, 0.2));
        }
        .sponsorNote {
            margin: 0;
            padding: 0 14px 12px;
            color: var(--vscode-descriptionForeground);
            line-height: 1.5;
        }
        .sponsorNote.thanks {
            font-size: 12px;
        }
        .sponsorList {
            display: flex;
            flex-direction: column;
            gap: 6px;
            padding: 0 14px 12px;
        }
        .sponsorRow {
            display: flex;
            align-items: stretch;
            gap: 6px;
            min-width: 0;
        }
        /* The whole row is the target, not the label, and the arrow on the
           right says where pressing it goes. */
        .sponsorOpen {
            flex: 1 1 auto;
            display: flex;
            align-items: center;
            gap: 10px;
            min-width: 0;
            padding: 8px 10px;
            text-align: left;
        }
        .sponsorIcon {
            flex: 0 0 auto;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 16px;
            height: 16px;
            font-size: 13px;
            line-height: 1;
        }
        .sponsorIcon svg {
            width: 16px;
            height: 16px;
            fill: currentColor;
        }
        .sponsorWords {
            display: flex;
            flex-direction: column;
            min-width: 0;
            flex: 1 1 auto;
        }
        .sponsorDetail {
            font-size: 11px;
            opacity: 0.7;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .sponsorAway {
            flex: 0 0 auto;
            opacity: 0.55;
        }
        .sponsorAway svg {
            width: 12px;
            height: 12px;
            fill: none;
            stroke: currentColor;
            stroke-width: 1.6;
        }
        .sponsorCopy {
            flex: 0 0 auto;
            padding: 3px 10px;
            font-size: 12px;
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
        function closeSelects(except) {
            document.querySelectorAll('[data-select-menu]').forEach((menu) => {
                if (menu === except) {
                    return;
                }
                menu.hidden = true;
                const button = document.querySelector('[data-select-toggle][aria-controls="' + menu.id + '"]');
                button?.setAttribute('aria-expanded', 'false');
            });
        }
        function optionButtons(menu) {
            return Array.from(menu.querySelectorAll('[data-select-option]'));
        }
        function focusOption(menu, direction) {
            const options = optionButtons(menu);
            if (options.length === 0) {
                return;
            }
            const currentIndex = Math.max(0, options.indexOf(document.activeElement));
            const nextIndex = direction === 'previous'
                ? (currentIndex + options.length - 1) % options.length
                : (currentIndex + 1) % options.length;
            options[nextIndex].focus();
        }
        function selectOption(option) {
            if (option.dataset.selectId === 'project.targetProject') {
                vscode.postMessage({
                    type: 'selectProject',
                    filePath: option.dataset.selectValue || undefined
                });
            }
            closeSelects();
        }
        const sponsorBackdrop = document.getElementById('sponsor-backdrop');
        let sponsorReturnFocus = null;
        function sponsorRing() {
            return Array.from(sponsorBackdrop.querySelectorAll('button')).filter((one) => !one.disabled);
        }
        function openSponsor() {
            if (!sponsorBackdrop) {
                return;
            }
            sponsorReturnFocus = document.activeElement;
            sponsorBackdrop.hidden = false;
            sponsorBackdrop.querySelector('[data-sponsor-open]')?.focus();
        }
        function closeSponsor() {
            if (!sponsorBackdrop || sponsorBackdrop.hidden) {
                return;
            }
            sponsorBackdrop.hidden = true;
            sponsorReturnFocus?.focus?.();
        }
        // Mousedown, not click: a drag that starts on the card and releases
        // over the backdrop is a missed text selection, not a request to close.
        sponsorBackdrop?.addEventListener('mousedown', (event) => {
            if (event.target === sponsorBackdrop) {
                closeSponsor();
            }
        });
        document.addEventListener('keydown', (event) => {
            if (!sponsorBackdrop || sponsorBackdrop.hidden) {
                return;
            }
            if (event.key === 'Escape') {
                event.preventDefault();
                event.stopPropagation();
                closeSponsor();
                return;
            }
            // The trap aria-modal claims: Tab cycles inside the card.
            if (event.key === 'Tab') {
                const ring = sponsorRing();
                const first = ring[0];
                const last = ring[ring.length - 1];
                if (!first) {
                    event.preventDefault();
                    return;
                }
                const active = document.activeElement;
                const inside = sponsorBackdrop.contains(active);
                if (event.shiftKey && (!inside || active === first)) {
                    event.preventDefault();
                    last.focus();
                } else if (!event.shiftKey && (!inside || active === last)) {
                    event.preventDefault();
                    first.focus();
                }
            }
        }, true);
        document.addEventListener('click', (event) => {
            const sponsorToggle = event.target.closest?.('[data-sponsor-toggle]');
            if (sponsorToggle) {
                openSponsor();
                return;
            }
            if (event.target.closest?.('[data-sponsor-close]')) {
                closeSponsor();
                return;
            }
            const option = event.target.closest?.('[data-select-option]');
            if (option) {
                selectOption(option);
                return;
            }
            const toggle = event.target.closest?.('[data-select-toggle]');
            if (toggle) {
                const menu = document.getElementById(toggle.getAttribute('aria-controls'));
                if (!menu) {
                    return;
                }
                const open = menu.hidden;
                closeSelects(open ? menu : undefined);
                menu.hidden = !open;
                toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
                if (open) {
                    const selected = menu.querySelector('[aria-selected="true"]');
                    (selected ?? menu.querySelector('[data-select-option]'))?.focus();
                }
                return;
            }
            const sponsorOpen = event.target.closest?.('[data-sponsor-open]');
            if (sponsorOpen) {
                vscode.postMessage({ type: 'openSponsorUrl', url: sponsorOpen.dataset.sponsorOpen });
                return;
            }
            const sponsorCopy = event.target.closest?.('[data-sponsor-copy]');
            if (sponsorCopy) {
                vscode.postMessage({ type: 'copySponsorUrl', url: sponsorCopy.dataset.sponsorCopy });
                sponsorCopy.textContent = 'Copied';
                window.setTimeout(() => { sponsorCopy.textContent = 'Copy'; }, 1200);
                return;
            }
            const button = event.target.closest('[data-command]');
            if (!button) {
                closeSelects();
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
        document.addEventListener('keydown', (event) => {
            const toggle = event.target.closest?.('[data-select-toggle]');
            if (toggle && (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowDown')) {
                event.preventDefault();
                const menu = document.getElementById(toggle.getAttribute('aria-controls'));
                if (!menu) {
                    return;
                }
                closeSelects(menu);
                menu.hidden = false;
                toggle.setAttribute('aria-expanded', 'true');
                (menu.querySelector('[aria-selected="true"]') ?? menu.querySelector('[data-select-option]'))?.focus();
                return;
            }
            const option = event.target.closest?.('[data-select-option]');
            if (!option) {
                if (event.key === 'Escape') {
                    closeSelects();
                }
                return;
            }
            const menu = option.closest('[data-select-menu]');
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                selectOption(option);
            } else if (event.key === 'ArrowDown') {
                event.preventDefault();
                focusOption(menu, 'next');
            } else if (event.key === 'ArrowUp') {
                event.preventDefault();
                focusOption(menu, 'previous');
            } else if (event.key === 'Escape') {
                event.preventDefault();
                closeSelects();
                const button = document.querySelector('[data-select-toggle][aria-controls="' + menu.id + '"]');
                button?.focus();
            }
        });
    </script>
</body>
</html>`;
}

function renderSection(section: XlideSidebarNode): string {
    if (section.id === 'sponsor') {
        return renderSponsorSection(section);
    }
    const children = section.children ?? [];
    const isActionSection = section.id === 'projectActions' ||
        section.id === 'settings' ||
        section.id === 'support';
    const sectionClass = children.some((child) => child.kind === 'select') ? 'section hasCustomSelect' : 'section';
    return `<section class="${sectionClass}" aria-label="${escapeAttr(section.label)}">
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

/**
 * The sponsor section renders as a quiet footer button that opens a modal,
 * the way the VBA editor add-in's heart button does, so the addresses never
 * sit above a workflow action. The modal's content is the section's nodes:
 * the blurb, the three link rows, and the thanks line.
 */
function renderSponsorSection(section: XlideSidebarNode): string {
    const children = section.children ?? [];
    const notes = children.filter((node) => node.kind === 'note');
    const links = children.filter((node) => node.kind === 'link');
    return `<div class="sponsorFooter">
        <button class="sponsorToggle" type="button" data-sponsor-toggle aria-haspopup="dialog" aria-controls="sponsor-backdrop" title="${escapeAttr(section.label)}">${HEART} Support</button>
    </div>
    <div class="sponsorBackdrop" id="sponsor-backdrop" hidden>
        <div class="sponsorCard" role="dialog" aria-modal="true" aria-labelledby="sponsor-title">
            <div class="sponsorHead">
                <div class="sponsorTitle" id="sponsor-title">${escapeHtml(section.label)}</div>
                <button class="sponsorClose" type="button" data-sponsor-close aria-label="Close" title="Close (Esc)">&times;</button>
            </div>
            ${notes[0] ? renderSponsorNote(notes[0]) : ''}
            <div class="sponsorList">${links.map((node) => renderSponsorRow(node)).join('')}</div>
            ${notes[1] ? renderSponsorNote(notes[1]) : ''}
        </div>
    </div>`;
}

/** The red heart, as the add-in's toolbar button draws it. */
const HEART = '\u2764\uFE0F';

function renderSponsorNote(node: XlideSidebarNode): string {
    const cls = node.id === 'sponsor.thanks' ? 'sponsorNote thanks' : 'sponsorNote';
    return `<p class="${cls}">${escapeHtml(node.label)}</p>`;
}

function renderSponsorRow(node: XlideSidebarNode): string {
    const url = node.url ?? '';
    return `<div class="sponsorRow">
        <button class="sponsorOpen secondary" type="button" data-sponsor-open="${escapeAttr(url)}" title="${escapeAttr(url)}">
            <span class="sponsorIcon" aria-hidden="true">${sponsorIconHtml(node.icon ?? '')}</span>
            <span class="sponsorWords">
                <span class="label">${escapeHtml(node.label)}</span>
                <span class="sponsorDetail">${escapeHtml(node.description ?? '')}</span>
            </span>
            <span class="sponsorAway" aria-hidden="true">${EXTERNAL_LINK_SVG}</span>
        </button>
        <button class="sponsorCopy secondary" type="button" data-sponsor-copy="${escapeAttr(url)}" title="Copy the address">Copy</button>
    </div>`;
}

/**
 * The webview's CSP allows no fonts or images, so the marks are inline SVG
 * paths. An icon name outside the set is printed as text: the emoji case.
 */
function sponsorIconHtml(icon: string): string {
    switch (icon) {
        case 'github':
            return GITHUB_SVG;
        case 'credit-card':
            return CREDIT_CARD_SVG;
        default:
            return escapeHtml(icon);
    }
}

const GITHUB_SVG = '<svg viewBox="0 0 16 16"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8z"/></svg>';

const CREDIT_CARD_SVG = '<svg viewBox="0 0 16 16"><path d="M1 3.5A1.5 1.5 0 0 1 2.5 2h11A1.5 1.5 0 0 1 15 3.5v9a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 1 12.5v-9zM2.5 3a.5.5 0 0 0-.5.5V5h12V3.5a.5.5 0 0 0-.5-.5h-11zM14 7H2v5.5a.5.5 0 0 0 .5.5h11a.5.5 0 0 0 .5-.5V7zM3 9h4v1.5H3V9z"/></svg>';

const EXTERNAL_LINK_SVG = '<svg viewBox="0 0 16 16"><path d="M9 2h5v5M14 2 7 9M12 9v4.5a.5.5 0 0 1-.5.5h-9a.5.5 0 0 1-.5-.5v-9a.5.5 0 0 1 .5-.5H7"/></svg>';

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
    const showDot = false;
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
    const showDot = false;
    const rowClass = showDot ? 'row selectRow' : 'row selectRow noDotRow';
    const selectedValue = node.value ?? '';
    const selectedOption = options.find((option) => option.value === selectedValue);
    const selectedLabel = selectedOption?.label ?? node.description ?? node.label;
    const menuId = `select-menu-${slugId(node.id)}`;
    return `<div class="${rowClass}" title="${escapeAttr(node.tooltip ?? node.label)}">
        ${showDot ? `<span class="dot ${escapeAttr(status)}" aria-hidden="true"></span>` : ''}
        <div class="rowText">
            <div class="label">${escapeHtml(node.label)}</div>
            ${node.description ? `<div class="description">${escapeHtml(node.description)}</div>` : ''}
            <div class="customSelect">
                <button
                    class="selectButton"
                    type="button"
                    data-select-toggle
                    aria-haspopup="listbox"
                    aria-expanded="false"
                    aria-controls="${escapeAttr(menuId)}"
                    title="${escapeAttr(selectedOption?.value || selectedLabel)}"
                >
                    <span class="selectButtonLabel">${escapeHtml(selectedLabel)}</span>
                    <span class="selectChevron" aria-hidden="true">&#9662;</span>
                </button>
                <div
                    class="selectMenu"
                    id="${escapeAttr(menuId)}"
                    role="listbox"
                    aria-label="${escapeAttr(node.label)}"
                    data-select-menu
                    hidden
                >
                    ${options.map((option) => renderSelectOption(node.id, option, selectedValue)).join('')}
                </div>
            </div>
        </div>
    </div>`;
}

function renderSelectOption(
    selectId: string,
    option: { label: string; value: string },
    selectedValue: string,
): string {
    const selected = option.value === selectedValue;
    return `<button
        class="selectOption"
        type="button"
        role="option"
        data-select-option
        data-select-id="${escapeAttr(selectId)}"
        data-select-value="${escapeAttr(option.value)}"
        aria-selected="${selected ? 'true' : 'false'}"
        title="${escapeAttr(option.value || option.label)}"
    >${escapeHtml(option.label)}</button>`;
}

function commandAttr(command: XlideSidebarCommand): string {
    return escapeAttr(JSON.stringify(command));
}

function slugId(value: string): string {
    return value.replace(/[^a-z0-9_-]/gi, '-');
}

export {
    XlideSidebarProvider,
    type XlideSidebarRegistration,
    registerXlideSidebar,
    renderXlideSidebarHtml,
    projectFiles,
};
