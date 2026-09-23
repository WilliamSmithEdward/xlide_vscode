import { workspaceFiles } from './util/workspaceFiles';
import { workspaceUriFor } from './util/workspaceUris';
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
import { sameProjectPath, XLIDE_SCHEME } from './xlideFileSystem';
import { resolveProjectTarget, SELECTED_PROJECT_STATE_KEY, storeSelectedProject } from './projectTarget';
import { AGENT_INSTRUCTIONS, AGENT_INSTRUCTIONS_STEPS } from './agentInstructions';
import {
    buildXlideSidebarModel,
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
        // The dialog shows the same text, but what is copied is the host's own.
        if (payload.type === 'copyAgentInstructions') {
            let copied = true;
            try {
                await vscode.env.clipboard.writeText(AGENT_INSTRUCTIONS);
            } catch {
                copied = false;
            }
            await this._view?.webview.postMessage({ type: copied ? 'agentInstructionsCopied' : 'agentInstructionsCopyFailed' });
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
        await storeSelectedProject(this._options.workspaceState, filePath);
    }
}

// The key itself lives with the ladder that reads it, so a keybinding
// resolving the target sees the same pick this sidebar wrote.
const SELECTED_WORKBOOK_KEY = SELECTED_PROJECT_STATE_KEY;

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
                    // Another window may have created it in between; the
                    // loser of that race writes nothing and opens what won.
                    await workspaceFiles.createTextIfAbsent(settingsPath, '{}\n');
                }
                const document = await vscode.workspace.openTextDocument(workspaceUriFor(settingsPath));
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

/**
 * The file the sidebar shows and its buttons act on. Resolved through the
 * shared ladder so a keybinding, which has no sidebar and no row to go on,
 * always lands on the same file as the Open button the user can see.
 */
async function activeProjectContext(
    projects: readonly vscode.Uri[],
    selectedProjectPath?: string,
): Promise<XlideSidebarActiveProject | undefined> {
    const target = await resolveProjectTarget({ projects, selectedProjectPath });
    return target ? sidebarProjectForPath(target.filePath, target.source) : undefined;
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
        /* A sidebar is narrow, so the card takes its full width behind a slim
           margin rather than floating as a fixed-width dialog would. */
        .dialogBackdrop {
            position: fixed;
            inset: 0;
            z-index: 50;
            display: flex;
            align-items: flex-start;
            justify-content: center;
            padding: 20px 8px 8px;
            background: rgba(0, 0, 0, 0.35);
        }
        .dialogBackdrop[hidden] {
            display: none;
        }
        .dialogCard {
            width: 100%;
            max-height: 100%;
            overflow: auto;
            background: var(--vscode-editorWidget-background, var(--vscode-sideBar-background));
            border: 1px solid var(--vscode-editorWidget-border, var(--vscode-panel-border));
            border-radius: 6px;
            box-shadow: 0 8px 28px rgba(0, 0, 0, 0.45);
            padding-bottom: 4px;
        }
        .dialogHead {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px 12px 6px;
        }
        .dialogTitle {
            font-size: 15px;
            font-weight: 600;
        }
        .dialogClose {
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
        .dialogClose:hover {
            opacity: 1;
            background: var(--vscode-toolbar-hoverBackground, rgba(128, 128, 128, 0.2));
        }
        /* The agent instructions: steps for the person, then the text for the
           agent in a box that can be selected and scrolled but not edited. */
        .agentCard {
            display: flex;
            flex-direction: column;
            gap: 10px;
            padding-bottom: 12px;
        }
        .agentSteps {
            margin: 0;
            padding: 0 12px 0 30px;
            line-height: 1.5;
        }
        .agentSteps li + li {
            margin-top: 4px;
        }
        .agentSteps code {
            font-family: var(--vscode-editor-font-family, monospace);
            font-size: 0.95em;
        }
        .agentTextLabel {
            padding: 0 12px;
            font-weight: 600;
        }
        .agentText {
            display: block;
            width: calc(100% - 24px);
            margin: -6px 12px 0;
            min-height: 180px;
            height: 45vh;
            resize: vertical;
            padding: 6px 8px;
            border: 1px solid var(--vscode-input-border, var(--vscode-panel-border));
            border-radius: 4px;
            color: var(--vscode-input-foreground);
            background: var(--vscode-input-background);
            font-family: var(--vscode-editor-font-family, monospace);
            font-size: 12px;
            line-height: 1.45;
        }
        .agentText:focus {
            outline: 1px solid var(--vscode-focusBorder);
            outline-offset: -1px;
        }
        .agentActions {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 0 12px;
        }
        .agentActions button {
            min-width: 72px;
        }
        .agentStatus {
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
        // The agent instructions dialog takes the focus until it closes.
        let openDialog = null;
        let dialogReturnFocus = null;
        function dialogRing() {
            return Array.from(openDialog.querySelectorAll('button, textarea')).filter((one) => !one.disabled);
        }
        function showDialog(id) {
            const dialog = document.getElementById(id);
            if (!dialog) {
                return;
            }
            closeDialog();
            dialogReturnFocus = document.activeElement;
            openDialog = dialog;
            dialog.hidden = false;
            (dialog.querySelector('[data-dialog-focus]') ?? dialogRing()[0])?.focus();
        }
        function closeDialog() {
            if (!openDialog) {
                return;
            }
            openDialog.hidden = true;
            openDialog = null;
            dialogReturnFocus?.focus?.();
        }
        // Mousedown, not click: a drag that starts on the card and releases
        // over the backdrop is a missed text selection, not a request to close.
        document.querySelectorAll('[data-dialog]').forEach((dialog) => {
            dialog.addEventListener('mousedown', (event) => {
                if (event.target === dialog) {
                    closeDialog();
                }
            });
        });
        document.addEventListener('keydown', (event) => {
            if (!openDialog) {
                return;
            }
            if (event.key === 'Escape') {
                event.preventDefault();
                event.stopPropagation();
                closeDialog();
                return;
            }
            // The trap aria-modal claims: Tab cycles inside the card.
            if (event.key === 'Tab') {
                const ring = dialogRing();
                const first = ring[0];
                const last = ring[ring.length - 1];
                if (!first) {
                    event.preventDefault();
                    return;
                }
                const active = document.activeElement;
                const inside = openDialog.contains(active);
                if (event.shiftKey && (!inside || active === first)) {
                    event.preventDefault();
                    last.focus();
                } else if (!event.shiftKey && (!inside || active === last)) {
                    event.preventDefault();
                    first.focus();
                }
            }
        }, true);
        // The host copies its own text and says whether the clipboard took it.
        const agentCopy = document.querySelector('[data-agent-copy]');
        const agentStatus = document.getElementById('agent-instructions-status');
        let agentCopyTimer = undefined;
        window.addEventListener('message', (event) => {
            const type = event.data?.type;
            if (!agentCopy || (type !== 'agentInstructionsCopied' && type !== 'agentInstructionsCopyFailed')) {
                return;
            }
            const copied = type === 'agentInstructionsCopied';
            agentCopy.textContent = copied ? 'Copied' : 'Copy';
            agentStatus.textContent = copied
                ? 'Copied to the clipboard.'
                : 'Could not copy. Select the text, then press Ctrl+C.';
            window.clearTimeout(agentCopyTimer);
            agentCopyTimer = window.setTimeout(() => {
                agentCopy.textContent = 'Copy';
                agentStatus.textContent = '';
            }, copied ? 2000 : 6000);
        });
        document.addEventListener('click', (event) => {
            const dialogOpen = event.target.closest?.('[data-dialog-open]');
            if (dialogOpen) {
                showDialog(dialogOpen.dataset.dialogOpen);
                return;
            }
            if (event.target.closest?.('[data-dialog-close]')) {
                closeDialog();
                return;
            }
            if (event.target.closest?.('[data-agent-copy]')) {
                vscode.postMessage({ type: 'copyAgentInstructions' });
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
    const children = section.children ?? [];
    const isActionSection = section.id === 'agenticAi' ||
        section.id === 'projectActions' ||
        section.id === 'settings' ||
        section.id === 'support';
    const sectionClass = children.some((child) => child.kind === 'select') ? 'section hasCustomSelect' : 'section';
    // A dialog sits beside its section, not inside it, so nothing clips it.
    const dialogs = children.some((node) => node.dialog === 'agentInstructions') ? renderAgentInstructionsDialog() : '';
    return `<section class="${sectionClass}" aria-label="${escapeAttr(section.label)}">
        <div class="sectionHeader">${escapeHtml(section.label)}</div>
        <div class="${isActionSection ? 'actionGrid' : 'sectionBody'}">
            ${children.length > 0
        ? children.map((node) => isActionSection && node.kind === 'action'
            ? renderActionNode(node)
            : renderSidebarNode(node)).join('')
        : '<div class="empty">No items</div>'}
        </div>
    </section>${dialogs}`;
}

/**
 * The agent instructions dialog: what to do with them, then the text itself,
 * selectable but not editable, and a Copy button. The host copies its own
 * copy of the text; the webview only asks.
 */
function renderAgentInstructionsDialog(): string {
    const steps = AGENT_INSTRUCTIONS_STEPS
        .map((step) => `<li>${escapeHtml(step).replace(/`([^`]+)`/g, '<code>$1</code>')}</li>`)
        .join('');
    return `<div class="dialogBackdrop" id="agent-instructions-dialog" data-dialog hidden>
        <div class="dialogCard agentCard" role="dialog" aria-modal="true" aria-labelledby="agent-instructions-title" aria-describedby="agent-instructions-steps">
            <div class="dialogHead">
                <div class="dialogTitle" id="agent-instructions-title">Agent Instructions</div>
                <button class="dialogClose" type="button" data-dialog-close aria-label="Close" title="Close (Esc)">&times;</button>
            </div>
            <ol class="agentSteps" id="agent-instructions-steps">${steps}</ol>
            <label class="agentTextLabel" for="agent-instructions-text">Instructions for your agent</label>
            <textarea class="agentText" id="agent-instructions-text" readonly spellcheck="false">${escapeHtml(AGENT_INSTRUCTIONS)}</textarea>
            <div class="agentActions">
                <button type="button" data-agent-copy data-dialog-focus>Copy</button>
                <button class="secondary" type="button" data-dialog-close>Close</button>
                <span class="agentStatus" id="agent-instructions-status" role="status"></span>
            </div>
        </div>
    </div>`;
}

function renderSidebarNode(node: XlideSidebarNode): string {
    if (node.kind === 'select') {
        return renderSelectNode(node);
    }
    if (node.kind === 'action') {
        return renderButtonOnlyRow(node);
    }
    return renderRowNode(node);
}

function renderActionNode(node: XlideSidebarNode): string {
    if (node.dialog) {
        return `<button class="actionCard secondary" type="button" data-dialog-open="${escapeAttr(dialogId(node.dialog))}" aria-haspopup="dialog" title="${escapeAttr(node.tooltip ?? node.label)}">
        <div class="label">${escapeHtml(node.label)}</div>
    </button>`;
    }
    if (!node.command && !node.disabled) {
        return renderRowNode(node);
    }
    const command = node.command && !node.disabled
        ? ` data-command="${commandAttr(node.command)}"`
        : ' disabled';
    return `<button class="actionCard secondary"${command} title="${escapeAttr(node.tooltip ?? node.label)}">
        <div class="label">${escapeHtml(node.label)}</div>
        ${node.description ? `<div class="description">${escapeHtml(node.description)}</div>` : ''}
    </button>`;
}

function dialogId(dialog: NonNullable<XlideSidebarNode['dialog']>): string {
    switch (dialog) {
        case 'agentInstructions':
            return 'agent-instructions-dialog';
    }
}

function renderButtonOnlyRow(node: XlideSidebarNode): string {
    if (!node.command && !node.disabled) {
        return renderRowNode(node);
    }
    const command = node.command && !node.disabled
        ? ` data-command="${commandAttr(node.command)}"`
        : ' disabled';
    return `<div class="row noDotRow buttonOnlyRow" title="${escapeAttr(node.tooltip ?? node.label)}">
        <button class="secondary"${command}>${escapeHtml(node.label)}</button>
    </div>`;
}

function renderRowNode(node: XlideSidebarNode): string {
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

function renderSelectNode(node: XlideSidebarNode): string {
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
