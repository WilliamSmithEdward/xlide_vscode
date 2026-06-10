import * as vscode from 'vscode';
import * as path from 'path';
import type { ImportMode, ModuleSyncFolderSource, ModuleSyncModeSource, ModuleSyncPlan } from './moduleSyncPlan';
import { settingsPathForWorkbook, type ExportMode } from './workbookSettings';
import { measurePerformance, measurePerformanceSync } from './performanceTrace';
import { randomNonce, scriptJson } from './webview/html';
import { webviewHeadHtml } from './webview/page';
import { DebouncedRefresher, RefreshGate } from './webview/refresh';
import { errorMessage } from './util/errors';

export interface ModuleSyncApplyResult {
    summary: string;
    changed: number;
    skipped: number;
    removed?: number;
    failed: number;
}

export interface ModuleSyncSettings {
    folderPath: string;
    folderPathSource?: ModuleSyncFolderSource;
    exportMode?: ExportMode;
    exportModeSource?: ModuleSyncModeSource;
    importMode?: ImportMode;
    importModeSource?: ModuleSyncModeSource;
    settingsPath?: string;
}

export interface ModuleSyncPreviewOptions {
    onChooseFolder?: (current: ModuleSyncSettings) => Promise<ModuleSyncPlan | undefined>;
    onRefresh?: (settings: ModuleSyncSettings) => Promise<ModuleSyncPlan>;
    onReloadWorkbookSettings?: () => Promise<ModuleSyncPlan | undefined>;
    onSaveSettings?: (settings: ModuleSyncSettings) => Promise<ModuleSyncApplyResult>;
}

export function openModuleSyncPreview(
    context: vscode.ExtensionContext,
    plan: ModuleSyncPlan,
    onApply: (plan: ModuleSyncPlan, selectedIds: readonly string[]) => Promise<ModuleSyncApplyResult>,
    options: ModuleSyncPreviewOptions = {},
): Promise<ModuleSyncApplyResult | undefined> {
    return new Promise((resolve) => {
        const panel = vscode.window.createWebviewPanel(
            'xlide.moduleSyncPreview',
            plan.direction === 'export' ? 'XLIDE Export Preview' : 'XLIDE Import Preview',
            vscode.ViewColumn.One,
            {
                enableScripts: true,
                retainContextWhenHidden: true,
                localResourceRoots: [context.extensionUri],
            },
        );
        let resolved = false;
        let disposed = false;
        let watchedFolderPath: string | undefined;
        let folderWatcherDisposables: vscode.Disposable[] = [];
        const done = (result: ModuleSyncApplyResult | undefined): void => {
            if (!resolved) {
                resolved = true;
                resolve(result);
            }
        };

        let currentPlan = plan;
        const disposeFolderWatcher = (): void => {
            for (const disposable of folderWatcherDisposables) {
                disposable.dispose();
            }
            folderWatcherDisposables = [];
            watchedFolderPath = undefined;
        };
        const postRefreshError = (err: unknown): void => {
            void panel.webview.postMessage({ type: 'error', error: errorMessage(err) });
        };
        const gate = new RefreshGate();
        const folderRefresher = new DebouncedRefresher({
            refresh: () => refreshPlanFromDisk(),
            onError: postRefreshError,
            defaultDelayMs: 300,
            gate,
        });
        const settingsRefresher = new DebouncedRefresher({
            refresh: () => refreshPlanFromWorkbookSettings(),
            onError: postRefreshError,
            defaultDelayMs: 300,
            gate,
        });
        const configureFolderWatcher = (): void => {
            if (!options.onRefresh || disposed || !currentPlan.folderPath || watchedFolderPath === currentPlan.folderPath) {
                return;
            }
            disposeFolderWatcher();
            watchedFolderPath = currentPlan.folderPath;
            const watcher = vscode.workspace.createFileSystemWatcher(
                new vscode.RelativePattern(currentPlan.folderPath, '**/*'),
            );
            const refreshIfRelevant = (uri: vscode.Uri): void => {
                if (isModuleSyncWatchedPath(uri.fsPath)) {
                    folderRefresher.schedule();
                }
            };
            folderWatcherDisposables = [
                watcher.onDidChange(refreshIfRelevant),
                watcher.onDidCreate(refreshIfRelevant),
                watcher.onDidDelete(refreshIfRelevant),
                watcher,
            ];
        };
        const runExclusive = <T>(operation: () => Promise<T>): Promise<T> => gate.runExclusive(operation);
        const updateCurrentPlan = (nextPlan: ModuleSyncPlan): void => {
            currentPlan = nextPlan;
            configureFolderWatcher();
        };
        async function refreshPlanFromDisk(): Promise<void> {
            const refresh = options.onRefresh;
            if (!refresh || disposed) {
                return;
            }
            await measurePerformance('moduleSync.refreshFromDisk', currentPlan.direction, async () => {
                await panel.webview.postMessage({ type: 'refreshing', message: 'Disk changes detected. Refreshing preview...' });
                updateCurrentPlan(await refresh(settingsFromPlan(currentPlan)));
                await panel.webview.postMessage({
                    type: 'plan',
                    plan: currentPlan,
                    message: 'Disk changes detected. Preview refreshed.',
                });
            });
        }
        async function refreshPlanFromWorkbookSettings(): Promise<void> {
            const reload = options.onReloadWorkbookSettings;
            if (!reload || disposed) {
                return;
            }
            await measurePerformance('moduleSync.refreshWorkbookSettings', currentPlan.direction, async () => {
                await panel.webview.postMessage({ type: 'refreshing', message: 'Workbook settings changed. Refreshing preview...' });
                const nextPlan = await reload();
                if (nextPlan) {
                    updateCurrentPlan(nextPlan);
                    await panel.webview.postMessage({
                        type: 'plan',
                        plan: currentPlan,
                        message: 'Workbook settings changed. Preview refreshed.',
                    });
                } else {
                    await panel.webview.postMessage({ type: 'ready' });
                }
            });
        }
        configureFolderWatcher();
        const workbookSettingsPath = settingsPathForWorkbook(currentPlan.workbookPath);
        const workbookSettingsWatcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(
            path.dirname(workbookSettingsPath),
            path.basename(workbookSettingsPath),
        ));
        const workbookSettingsWatcherDisposables = [
            workbookSettingsWatcher.onDidCreate(() => settingsRefresher.schedule()),
            workbookSettingsWatcher.onDidChange(() => settingsRefresher.schedule()),
            workbookSettingsWatcher.onDidDelete(() => settingsRefresher.schedule()),
            workbookSettingsWatcher,
        ];
        panel.webview.html = measurePerformanceSync(
            'moduleSync.renderHtml',
            `${currentPlan.direction}:${currentPlan.items.length} items`,
            () => renderModuleSyncHtml(currentPlan),
        );
        const messageSub = panel.webview.onDidReceiveMessage(async (message: {
            type?: string;
            selectedIds?: string[];
            folderPath?: string;
            exportMode?: ExportMode;
            importMode?: ImportMode;
            itemId?: string;
            side?: 'left' | 'right';
            showHeaders?: boolean;
            autoSaveSettings?: boolean;
            quiet?: boolean;
        }) => {
            if (message.type === 'cancel') {
                done(undefined);
                panel.dispose();
                return;
            }
            if (message.type === 'copy-code') {
                const item = currentPlan.items.find((candidate) => candidate.id === message.itemId);
                if (!item || (message.side !== 'left' && message.side !== 'right')) {
                    await panel.webview.postMessage({ type: 'error', error: 'No code selected to copy.' });
                    return;
                }
                const code = message.side === 'left'
                    ? (message.showHeaders ? item.leftRawCode : item.leftCode)
                    : (message.showHeaders ? item.rightRawCode : item.rightCode);
                await vscode.env.clipboard.writeText(code);
                await panel.webview.postMessage({ type: 'copied', side: message.side });
                return;
            }
            if (message.type === 'choose-folder') {
                const chooseFolder = options.onChooseFolder;
                if (!chooseFolder) {
                    return;
                }
                await panel.webview.postMessage({ type: 'refreshing' });
                try {
                    await measurePerformance('moduleSync.chooseFolder', currentPlan.direction, async () => {
                        await runExclusive(async () => {
                            const nextPlan = await chooseFolder(settingsFromMessage(currentPlan, message));
                            if (nextPlan) {
                                updateCurrentPlan(nextPlan);
                                await panel.webview.postMessage({
                                    type: 'plan',
                                    plan: currentPlan,
                                    message: 'Settings updated. Review the refreshed diff before applying.',
                                    autoSaveSettings: true,
                                });
                            } else {
                                await panel.webview.postMessage({ type: 'ready' });
                            }
                        });
                    });
                } catch (err) {
                    const error = errorMessage(err);
                    await panel.webview.postMessage({ type: 'error', error });
                }
                return;
            }
            if (message.type === 'refresh-settings') {
                const refresh = options.onRefresh;
                if (!refresh) {
                    return;
                }
                await panel.webview.postMessage({ type: 'refreshing' });
                try {
                    await measurePerformance('moduleSync.refreshSettings', currentPlan.direction, async () => {
                        await runExclusive(async () => {
                            updateCurrentPlan(await refresh(settingsFromMessage(currentPlan, message)));
                            await panel.webview.postMessage({
                                type: 'plan',
                                plan: currentPlan,
                                message: 'Settings updated. Review the refreshed diff before applying.',
                                autoSaveSettings: message.autoSaveSettings === true,
                            });
                        });
                    });
                } catch (err) {
                    const error = errorMessage(err);
                    await panel.webview.postMessage({ type: 'error', error });
                }
                return;
            }
            if (message.type === 'save-settings') {
                const saveSettings = options.onSaveSettings;
                if (!saveSettings) {
                    return;
                }
                if (!message.quiet) {
                    await panel.webview.postMessage({ type: 'saving-settings' });
                }
                try {
                    await runExclusive(async () => {
                        const result = await saveSettings(settingsFromMessage(currentPlan, message));
                        await panel.webview.postMessage({ type: 'settings-saved', result, quiet: message.quiet === true });
                    });
                } catch (err) {
                    const error = errorMessage(err);
                    await panel.webview.postMessage({ type: 'error', error });
                }
                return;
            }
            if (message.type !== 'apply') {
                return;
            }
            const selectedIds = message.selectedIds ?? [];
            await panel.webview.postMessage({ type: 'applying' });
            try {
                await measurePerformance('moduleSync.apply', currentPlan.direction, async () => {
                    await runExclusive(async () => {
                        const result = await onApply(currentPlan, selectedIds);
                        await panel.webview.postMessage({ type: 'applied', result });
                        done(result);
                    });
                });
            } catch (err) {
                const error = errorMessage(err);
                await panel.webview.postMessage({ type: 'error', error });
            }
        });
        panel.onDidDispose(() => {
            disposed = true;
            folderRefresher.dispose();
            settingsRefresher.dispose();
            disposeFolderWatcher();
            for (const disposable of workbookSettingsWatcherDisposables) {
                disposable.dispose();
            }
            messageSub.dispose();
            done(undefined);
        });
    });
}

function renderModuleSyncHtml(plan: ModuleSyncPlan): string {
    const nonce = randomNonce();
    const data = scriptJson(plan);
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    ${webviewHeadHtml(nonce, plan.title)}
    <style nonce="${nonce}">
        :root {
            --border: color-mix(in srgb, var(--vscode-editor-foreground) 18%, transparent);
            --muted: var(--vscode-descriptionForeground);
            --row: color-mix(in srgb, var(--vscode-editor-background) 90%, var(--vscode-editor-foreground) 10%);
            --changed: color-mix(in srgb, var(--vscode-gitDecoration-modifiedResourceForeground) 16%, transparent);
            --added: color-mix(in srgb, var(--vscode-gitDecoration-addedResourceForeground) 18%, transparent);
            --removed: color-mix(in srgb, var(--vscode-gitDecoration-deletedResourceForeground) 16%, transparent);
            --warn: color-mix(in srgb, var(--vscode-editorWarning-foreground) 18%, transparent);
        }
        * { box-sizing: border-box; }
        body {
            margin: 0;
            color: var(--vscode-editor-foreground);
            background: var(--vscode-editor-background);
            font: var(--vscode-font-size) var(--vscode-font-family);
            height: 100vh;
            overflow: hidden;
        }
        .shell { display: grid; grid-template-rows: auto 1fr auto; height: 100vh; }
        header {
            padding: 12px 16px;
            border-bottom: 1px solid var(--border);
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 12px;
            align-items: center;
        }
        h1 {
            margin: 0 0 4px;
            font-size: 16px;
            font-weight: 600;
        }
        .sub { color: var(--muted); font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .settings {
            margin-top: 8px;
            display: grid;
            grid-template-columns: minmax(180px, 1fr) auto auto;
            gap: 8px;
            align-items: end;
            max-width: 780px;
        }
        .field {
            display: grid;
            gap: 3px;
            min-width: 0;
        }
        .field label {
            color: var(--muted);
            font-size: 11px;
        }
        .folderValue {
            min-height: 28px;
            padding: 5px 8px;
            border: 1px solid var(--border);
            background: var(--vscode-input-background);
            color: var(--vscode-input-foreground);
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .settingMeta {
            color: var(--muted);
            font-size: 11px;
            min-height: 14px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        select {
            min-height: 28px;
            color: var(--vscode-dropdown-foreground);
            background: var(--vscode-dropdown-background);
            border: 1px solid var(--vscode-dropdown-border);
            padding: 3px 8px;
        }
        .hidden { display: none; }
        .actions { display: flex; gap: 8px; align-items: center; }
        button {
            color: var(--vscode-button-foreground);
            background: var(--vscode-button-background);
            border: 0;
            padding: 5px 10px;
            min-height: 28px;
            border-radius: 2px;
            cursor: pointer;
        }
        button.secondary {
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
        }
        button:disabled { opacity: .55; cursor: default; }
        main {
            --list-width: 34%;
            display: grid;
            grid-template-columns: minmax(280px, var(--list-width)) 6px minmax(320px, 1fr);
            min-height: 0;
        }
        aside { min-height: 0; display: grid; grid-template-rows: auto auto 1fr; }
        .splitter {
            cursor: col-resize;
            background: color-mix(in srgb, var(--border) 45%, transparent);
            border-left: 1px solid var(--border);
            border-right: 1px solid var(--border);
        }
        .splitter:hover,
        .splitter.dragging {
            background: color-mix(in srgb, var(--vscode-focusBorder) 60%, transparent);
        }
        body.resizing {
            cursor: col-resize;
            user-select: none;
        }
        .toolbar {
            display: flex;
            gap: 8px;
            padding: 8px;
            border-bottom: 1px solid var(--border);
        }
        .toolbar button {
            flex: 1;
            border: 1px solid var(--vscode-button-border, transparent);
            font-weight: 600;
        }
        .toolbar button.secondary {
            border: 1px solid var(--border);
        }
        .warnings {
            padding: 8px 12px;
            display: none;
            background: var(--warn);
            border-bottom: 1px solid var(--border);
            color: var(--vscode-editorWarning-foreground);
            line-height: 1.35;
        }
        .warnings.visible { display: block; }
        .list { overflow: auto; }
        .item {
            width: 100%;
            display: grid;
            grid-template-columns: 56px minmax(0, 1fr) 132px;
            gap: 8px;
            align-items: center;
            padding: 8px 10px 8px 3px;
            border-bottom: 1px solid var(--border);
            border-left: 3px solid transparent;
            cursor: pointer;
        }
        .item.status-create { border-left-color: var(--vscode-gitDecoration-addedResourceForeground); }
        .item.status-remove, .item.status-error { border-left-color: var(--vscode-gitDecoration-deletedResourceForeground); }
        .item.status-write { border-left-color: var(--vscode-gitDecoration-modifiedResourceForeground); }
        .item.status-skip { border-left-color: var(--vscode-editorWarning-foreground); }
        .item:hover, .item.active { background: var(--row); }
        .checkHit {
            width: 52px;
            min-height: 44px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
        }
        .checkHit.disabled { cursor: default; }
        .item input {
            margin: 0;
            width: 18px;
            height: 18px;
            cursor: pointer;
        }
        .item input:disabled { cursor: default; }
        .itemText { min-width: 0; }
        .name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .meta { color: var(--muted); font-size: 12px; margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .badge {
            justify-self: end;
            max-width: 132px;
            font-size: 11px;
            padding: 2px 6px;
            border-radius: 999px;
            border: 1px solid var(--border);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .badge.create { color: var(--vscode-gitDecoration-addedResourceForeground); background: var(--added); }
        .badge.remove, .badge.error { color: var(--vscode-gitDecoration-deletedResourceForeground); background: var(--removed); }
        .badge.write { color: var(--vscode-gitDecoration-modifiedResourceForeground); background: var(--changed); }
        .badge.skip { color: var(--vscode-editorWarning-foreground); background: var(--warn); }
        .badge.same { color: var(--muted); }
        section {
            --diff-grid: 56px minmax(0, 1fr) 56px minmax(0, 1fr);
            min-height: 0;
            display: grid;
            grid-template-rows: auto auto 1fr;
        }
        .diff-tools {
            padding: 8px 12px;
            border-bottom: 1px solid var(--border);
            display: flex;
            justify-content: flex-end;
            gap: 8px;
        }
        .diff-head {
            border-bottom: 1px solid var(--border);
            display: grid;
            grid-template-columns: var(--diff-grid);
            min-height: 40px;
        }
        .diff-head-gutter {
            border-right: 1px solid var(--border);
        }
        .diff-title-wrap {
            min-width: 0;
            display: flex;
            gap: 8px;
            align-items: center;
            justify-content: flex-start;
            padding: 8px;
        }
        .diff-title {
            flex: 0 1 auto;
            min-width: 0;
            font-weight: 600;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .copyCode {
            flex: 0 0 auto;
            min-height: 24px;
            padding: 3px 8px;
            font-size: 11px;
        }
        .diff { overflow: auto; font-family: var(--vscode-editor-font-family); font-size: var(--vscode-editor-font-size); }
        .line {
            display: grid;
            grid-template-columns: var(--diff-grid);
            min-height: 20px;
            border-bottom: 1px solid color-mix(in srgb, var(--border) 40%, transparent);
        }
        .ln {
            color: var(--muted);
            text-align: right;
            padding: 2px 8px;
            user-select: none;
            border-right: 1px solid var(--border);
        }
        pre {
            margin: 0;
            padding: 2px 8px;
            white-space: pre-wrap;
            overflow-wrap: anywhere;
            min-width: 0;
        }
        .line.changed pre { background: var(--changed); }
        .line.added pre:not(:empty) { background: var(--added); }
        .line.removed pre:not(:empty) { background: var(--removed); }
        footer {
            min-height: 34px;
            padding: 8px 12px;
            border-top: 1px solid var(--border);
            color: var(--muted);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
        }
        .result { color: var(--vscode-editor-foreground); }
        .quickTooltip {
            position: fixed;
            z-index: 1000;
            max-width: min(420px, calc(100vw - 16px));
            padding: 6px 8px;
            border: 1px solid var(--vscode-editorWidget-border, var(--border));
            border-radius: 3px;
            color: var(--vscode-editorWidget-foreground, var(--vscode-editor-foreground));
            background: var(--vscode-editorWidget-background, var(--vscode-editor-background));
            box-shadow: 0 4px 12px color-mix(in srgb, black 35%, transparent);
            font-size: 12px;
            line-height: 1.35;
            pointer-events: none;
            opacity: 0;
            transform: translateY(2px);
            transition: opacity 70ms ease, transform 70ms ease;
        }
        .quickTooltip.visible {
            opacity: 1;
            transform: translateY(0);
        }
    </style>
</head>
<body>
    <div class="shell">
        <header>
            <div>
                <h1 id="title"></h1>
                <div class="sub" id="subtitle"></div>
                <div class="settings">
                    <div class="field">
                        <label id="folderLabel"></label>
                        <div class="folderValue" id="folderValue"></div>
                        <div class="settingMeta" id="folderSource"></div>
                    </div>
                    <button class="secondary" id="chooseFolder">Change</button>
                    <div class="field" id="modeField">
                        <label for="syncMode" id="modeLabel">Export mode</label>
                        <select id="syncMode">
                            <option value="exportAll">Export All (No Deletes)</option>
                            <option value="trueUp">Export All + Delete Missing</option>
                        </select>
                        <div class="settingMeta" id="modeSource"></div>
                    </div>
                </div>
            </div>
            <div class="actions">
                <button class="secondary" id="cancel">Cancel</button>
                <button id="apply">Apply Selected</button>
            </div>
        </header>
        <main id="layout">
            <aside>
                <div class="toolbar">
                    <button class="secondary" id="selectChanged">Select Pending</button>
                    <button class="secondary" id="clear">Clear</button>
                </div>
                <div class="warnings" id="warnings"></div>
                <div class="list" id="list"></div>
            </aside>
            <div class="splitter" id="splitter" role="separator" aria-orientation="vertical" aria-label="Resize module list"></div>
            <section>
                <div class="diff-tools">
                    <button class="secondary" id="toggleHeaders" aria-pressed="false">Show Headers in Diff</button>
                </div>
                <div class="diff-head">
                    <div class="diff-head-gutter"></div>
                    <div class="diff-title-wrap">
                        <div class="diff-title" id="leftTitle"></div>
                        <button class="secondary copyCode" id="copyLeft">Copy</button>
                    </div>
                    <div class="diff-head-gutter"></div>
                    <div class="diff-title-wrap">
                        <div class="diff-title" id="rightTitle"></div>
                        <button class="secondary copyCode" id="copyRight">Copy</button>
                    </div>
                </div>
                <div class="diff" id="diff"></div>
            </section>
        </main>
        <footer>
            <span id="counts"></span>
            <span class="result" id="result"></span>
        </footer>
    </div>
    <div class="quickTooltip" id="tooltip" role="tooltip" hidden></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        let plan = ${data};
        let selected = new Set();
        let activeId = plan.items[0]?.id;
        let applying = false;
        let applied = false;
        let showHeaders = false;
        let settingsSaving = false;
        let settingsSaveTimer;
        let tooltipTimer;
        let tooltipTarget;
        const tooltipDelayMs = 140;

        const el = id => document.getElementById(id);

        function clamp(value, min, max) {
            return Math.max(min, Math.min(max, value));
        }

        function setListWidth(width) {
            const layout = el('layout');
            const bounds = layout.getBoundingClientRect();
            const max = Math.max(320, bounds.width - 360);
            layout.style.setProperty('--list-width', \`\${clamp(width, 280, max)}px\`);
        }

        function installSplitter() {
            const splitter = el('splitter');
            const layout = el('layout');
            splitter.addEventListener('pointerdown', event => {
                event.preventDefault();
                splitter.setPointerCapture(event.pointerId);
                splitter.classList.add('dragging');
                document.body.classList.add('resizing');
            });
            splitter.addEventListener('pointermove', event => {
                if (!splitter.hasPointerCapture(event.pointerId)) return;
                const bounds = layout.getBoundingClientRect();
                setListWidth(event.clientX - bounds.left);
            });
            function stopDrag(event) {
                if (splitter.hasPointerCapture(event.pointerId)) {
                    splitter.releasePointerCapture(event.pointerId);
                }
                splitter.classList.remove('dragging');
                document.body.classList.remove('resizing');
            }
            splitter.addEventListener('pointerup', stopDrag);
            splitter.addEventListener('pointercancel', stopDrag);
        }

        function selectedFromPlan() {
            return new Set(plan.items.filter(item => item.checked && item.selectable).map(item => item.id));
        }

        function isRelevantItem(item) {
            return item.selectable && item.status !== 'unchanged' && !item.status.startsWith('skipping');
        }

        function statusTone(item) {
            if (item.status === 'will-create') return 'create';
            if (item.status === 'will-remove') return 'remove';
            if (item.status === 'will-write' || item.status === 'will-update') return 'write';
            if (item.status === 'unchanged') return 'same';
            if (item.status.startsWith('skipping')) return 'skip';
            return 'error';
        }

        function copyTooltip(item, side, hasCode) {
            if (hasCode) {
                return side === 'left' ? 'Copy left code to clipboard.' : 'Copy right code to clipboard.';
            }
            const title = side === 'left' ? item.leftTitle : item.rightTitle;
            if (title.startsWith('Repo:')) {
                return item.existsInRepo ? 'No repo file code to copy.' : 'Repo file does not exist yet.';
            }
            if (title.startsWith('Workbook:')) {
                return item.existsInWorkbook ? 'No workbook module code to copy.' : 'Workbook module does not exist yet.';
            }
            return side === 'left' ? 'No left-side code to copy.' : 'No right-side code to copy.';
        }

        function shouldShowWarnings() {
            return plan.warnings.some(warning => !warning.includes('skipping import unless the module already exists in the workbook'));
        }

        function currentSettings() {
            return {
                folderPath: plan.folderPath,
                exportMode: plan.direction === 'export' ? el('syncMode').value : undefined,
                importMode: plan.direction === 'import' ? el('syncMode').value : undefined,
            };
        }

        function option(value, label, description) {
            const item = document.createElement('option');
            item.value = value;
            item.textContent = label;
            if (description) {
                item.title = description;
            }
            return item;
        }

        function setTooltip(targetOrId, text) {
            const target = typeof targetOrId === 'string' ? el(targetOrId) : targetOrId;
            if (!target) return;
            target.dataset.tooltip = text;
            target.removeAttribute('title');
        }

        function clearTooltip(targetOrId) {
            const target = typeof targetOrId === 'string' ? el(targetOrId) : targetOrId;
            if (!target) return;
            delete target.dataset.tooltip;
            target.removeAttribute('title');
        }

        function clearTooltipTimer() {
            if (tooltipTimer) {
                clearTimeout(tooltipTimer);
                tooltipTimer = undefined;
            }
        }

        function hideTooltip() {
            clearTooltipTimer();
            tooltipTarget = undefined;
            const tooltip = el('tooltip');
            tooltip.classList.remove('visible');
            tooltip.hidden = true;
        }

        function showTooltipFor(target) {
            const text = target.dataset.tooltip;
            if (!text) return;
            const tooltip = el('tooltip');
            tooltip.textContent = text;
            tooltip.hidden = false;
            tooltip.classList.remove('visible');
            const rect = target.getBoundingClientRect();
            const pad = 8;
            const tooltipWidth = tooltip.offsetWidth;
            const tooltipHeight = tooltip.offsetHeight;
            const maxLeft = Math.max(pad, window.innerWidth - tooltipWidth - pad);
            const left = clamp(rect.left, pad, maxLeft);
            let top = rect.bottom + 6;
            if (top + tooltipHeight > window.innerHeight - pad) {
                top = Math.max(pad, rect.top - tooltipHeight - 6);
            }
            tooltip.style.left = \`\${left}px\`;
            tooltip.style.top = \`\${top}px\`;
            requestAnimationFrame(() => tooltip.classList.add('visible'));
        }

        function installFastTooltips() {
            document.addEventListener('mouseover', event => {
                const target = event.target.closest?.('[data-tooltip]');
                if (!target) return;
                tooltipTarget = target;
                clearTooltipTimer();
                tooltipTimer = setTimeout(() => {
                    if (tooltipTarget === target) {
                        showTooltipFor(target);
                    }
                }, tooltipDelayMs);
            });
            document.addEventListener('mouseout', event => {
                const target = event.target.closest?.('[data-tooltip]');
                if (!target) return;
                if (event.relatedTarget && target.contains(event.relatedTarget)) return;
                hideTooltip();
            });
            document.addEventListener('focusin', event => {
                const target = event.target.closest?.('[data-tooltip]');
                if (target) {
                    tooltipTarget = target;
                    showTooltipFor(target);
                }
            });
            document.addEventListener('focusout', () => hideTooltip());
            document.addEventListener('pointerdown', event => {
                if (event.target.closest?.('[data-tooltip]')) {
                    hideTooltip();
                }
            });
            document.addEventListener('click', event => {
                if (event.target.closest?.('[data-tooltip]')) {
                    hideTooltip();
                }
            });
            document.addEventListener('scroll', () => hideTooltip(), true);
            window.addEventListener('resize', () => hideTooltip());
        }

        function modeDescription(modeValue) {
            const mode = modeValue || el('syncMode').value;
            if (plan.direction === 'export') {
                return mode === 'trueUp'
                    ? 'Export every workbook module to the selected folder, then delete stale .bas/.cls module files that no longer exist in the workbook.'
                    : 'Export every workbook module to the selected folder. XLIDE will create missing module files and update changed files, but will not delete stale files.';
            }
            return mode === 'trueUpStandardClass'
                ? 'Import/update selected .bas/.cls files, then delete workbook-only standard/class modules missing from the folder. New standard/class modules can be created; existing document modules and UserForm .cls code-behind are updated on name match; document modules and UserForm code-behind are never created or deleted by this mode.'
                : 'Import/update selected .bas/.cls files without deleting workbook modules. New standard/class modules can be created; existing document modules and UserForm .cls code-behind are updated on name match; missing document modules and UserForm code-behind are skipped because XLIDE cannot create them directly.';
        }

        function settingsSourceLabel(source) {
            if (source === 'workbook') return 'Workbook override';
            if (source === 'session') return 'Current session';
            if (source === 'machine') return 'VS Code machine setting';
            if (source === 'unknown') return 'Unknown';
            return 'Built-in default';
        }

        function folderSourceLabel(source) {
            if (source === 'workbook') return 'Workbook sidecar';
            if (source === 'session') return 'Current session';
            return 'Not saved';
        }

        function settingsPathDescription() {
            return plan.settingsPath ? \` Settings file: \${plan.settingsPath}\` : '';
        }

        function updateModeTitle() {
            clearTooltip('syncMode');
            clearTooltip('modeLabel');
            clearTooltip('modeField');
        }

        function renderChrome() {
            el('title').textContent = plan.title;
            el('subtitle').textContent = \`\${plan.workbookPath} <-> \${plan.folderPath}\${plan.exportMode ? '  [' + plan.exportMode + ']' : ''}\`;
            el('folderLabel').textContent = plan.direction === 'export' ? 'Export folder' : 'Import folder';
            el('folderValue').textContent = plan.folderPath;
            el('folderSource').textContent = \`Source: \${folderSourceLabel(plan.folderPathSource)}\`;
            setTooltip('folderValue', plan.direction === 'export'
                ? \`Folder XLIDE will compare against and write selected workbook modules into: \${plan.folderPath}\`
                : \`Folder XLIDE will compare against and import selected module files from: \${plan.folderPath}\`);
            setTooltip('folderSource', \`This folder is workbook-scoped.\${settingsPathDescription()}\`);
            setTooltip('chooseFolder', plan.direction === 'export'
                ? 'Choose the folder to compare with this workbook and receive exported module files.'
                : 'Choose the folder containing module files to compare with and import into this workbook.');
            el('modeField').classList.remove('hidden');
            const mode = el('syncMode');
            mode.innerHTML = '';
            if (plan.direction === 'export') {
                el('modeLabel').textContent = 'Export mode';
                mode.append(option('exportAll', 'Export All (No Deletes)', modeDescription('exportAll')));
                mode.append(option('trueUp', 'Export All + Delete Missing', modeDescription('trueUp')));
                mode.value = plan.exportMode || 'exportAll';
            } else {
                el('modeLabel').textContent = 'Import mode';
                mode.append(option('updateOnly', 'Import/Update (No Deletes)', modeDescription('updateOnly')));
                mode.append(option('trueUpStandardClass', 'Import/Update + Delete Missing', modeDescription('trueUpStandardClass')));
                mode.value = plan.importMode || 'updateOnly';
            }
            const modeSource = plan.direction === 'export' ? plan.exportModeSource : plan.importModeSource;
            el('modeSource').textContent = \`Source: \${settingsSourceLabel(modeSource)}\`;
            setTooltip('modeSource', \`This mode uses a workbook override when present, otherwise the built-in default.\${settingsPathDescription()}\`);
            updateModeTitle();
            el('selectChanged').textContent = 'Select Pending';
            setTooltip('selectChanged', plan.direction === 'import'
                ? 'Select every pending import row that will create, update, or delete a workbook module under the current import mode.'
                : 'Select every pending export row that will create, overwrite, or remove files under the current export mode.');
            setTooltip('clear', 'Clear the current module selection without changing files.');
            setTooltip('apply', plan.direction === 'import'
                ? 'Apply the selected import changes to the workbook.'
                : 'Apply the selected export changes to the folder.');
            setTooltip('cancel', 'Close this preview without applying changes.');
            setTooltip('toggleHeaders', 'Toggle hidden VBA Attribute header lines in the diff preview and copy buttons.');
            if (shouldShowWarnings()) {
                el('warnings').classList.add('visible');
                el('warnings').textContent = plan.warnings.join('\\n');
            } else {
                el('warnings').classList.remove('visible');
                el('warnings').textContent = '';
            }
        }

        function setPlan(nextPlan, message, autoSaveSettings) {
            plan = nextPlan;
            selected = selectedFromPlan();
            activeId = plan.items[0]?.id;
            applying = false;
            applied = false;
            el('apply').textContent = 'Apply Selected';
            el('result').textContent = message || '';
            renderChrome();
            renderList();
            renderDiff();
            if (autoSaveSettings) {
                scheduleSettingsAutosave();
            }
        }

        function renderList() {
            const list = el('list');
            list.innerHTML = '';
            for (const item of plan.items) {
                const tone = statusTone(item);
                const row = document.createElement('div');
                row.className = \`item status-\${tone}\${item.id === activeId ? ' active' : ''}\`;
                row.dataset.id = item.id;
                row.setAttribute('aria-selected', item.id === activeId ? 'true' : 'false');
                const checkHit = document.createElement('div');
                checkHit.className = 'checkHit' + (!item.selectable ? ' disabled' : '');
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.checked = selected.has(item.id);
                checkbox.disabled = !item.selectable;
                checkHit.addEventListener('click', event => {
                    event.stopPropagation();
                    if (!item.selectable) return;
                    activeId = item.id;
                    if (selected.has(item.id)) selected.delete(item.id);
                    else selected.add(item.id);
                    renderList();
                    renderDiff();
                });
                checkHit.append(checkbox);
                const text = document.createElement('div');
                text.className = 'itemText';
                const name = document.createElement('div');
                name.className = 'name';
                name.textContent = item.moduleName;
                const meta = document.createElement('div');
                meta.className = 'meta';
                meta.textContent = [item.relativeName, item.moduleType, item.warning].filter(Boolean).join(' | ');
                text.append(name, meta);
                const badge = document.createElement('span');
                badge.className = \`badge \${tone}\`;
                badge.textContent = item.detail || item.status;
                badge.title = item.warning || item.detail || item.status;
                row.append(checkHit, text, badge);
                row.addEventListener('click', () => {
                    activeId = item.id;
                    renderList();
                    renderDiff();
                });
                list.append(row);
            }
            renderCounts();
        }

        function renderDiff() {
            const item = plan.items.find(candidate => candidate.id === activeId) || plan.items[0];
            const diff = el('diff');
            diff.innerHTML = '';
            if (!item) {
                el('leftTitle').textContent = '';
                el('rightTitle').textContent = '';
                el('copyLeft').disabled = true;
                el('copyRight').disabled = true;
                el('toggleHeaders').disabled = true;
                el('toggleHeaders').textContent = showHeaders ? 'Hide Headers in Diff' : 'Show Headers in Diff';
                el('toggleHeaders').setAttribute('aria-pressed', String(showHeaders));
                const empty = document.createElement('pre');
                empty.textContent = 'No module differences found for the current settings.';
                diff.append(empty);
                return;
            }
            const leftCode = showHeaders ? item.leftRawCode : item.leftCode;
            const rightCode = showHeaders ? item.rightRawCode : item.rightCode;
            const diffLines = showHeaders ? item.diffWithHeaders : item.diff;
            el('leftTitle').textContent = item.leftTitle;
            el('rightTitle').textContent = item.rightTitle;
            el('copyLeft').disabled = !leftCode;
            el('copyRight').disabled = !rightCode;
            setTooltip('copyLeft', copyTooltip(item, 'left', Boolean(leftCode)));
            setTooltip('copyRight', copyTooltip(item, 'right', Boolean(rightCode)));
            el('toggleHeaders').disabled = false;
            el('toggleHeaders').textContent = showHeaders ? 'Hide Headers in Diff' : 'Show Headers in Diff';
            el('toggleHeaders').setAttribute('aria-pressed', String(showHeaders));
            for (const line of diffLines) {
                const row = document.createElement('div');
                row.className = 'line ' + line.kind;
                const leftNo = document.createElement('div');
                leftNo.className = 'ln';
                leftNo.textContent = line.leftNumber || '';
                const left = document.createElement('pre');
                left.className = 'left';
                left.textContent = line.left;
                const rightNo = document.createElement('div');
                rightNo.className = 'ln';
                rightNo.textContent = line.rightNumber || '';
                const right = document.createElement('pre');
                right.className = 'right';
                right.textContent = line.right;
                row.append(leftNo, left, rightNo, right);
                diff.append(row);
            }
        }

        function renderCounts() {
            const selectedItems = plan.items.filter(item => selected.has(item.id));
            const unsupported = selectedItems.filter(item => item.unsupportedDirectCreation).length;
            const settingsStatus = settingsSaving ? ' | auto-saving settings' : '';
            el('counts').textContent = \`\${selectedItems.length} selected\${unsupported ? ' | ' + unsupported + ' will show skipping import warning' : ''}\${settingsStatus}\`;
            el('apply').disabled = applying || applied || selectedItems.length === 0;
            el('chooseFolder').disabled = applying;
            el('syncMode').disabled = applying;
        }

        function clearSettingsAutosave() {
            if (settingsSaveTimer) {
                clearTimeout(settingsSaveTimer);
                settingsSaveTimer = undefined;
            }
        }

        function scheduleSettingsAutosave() {
            clearSettingsAutosave();
            el('result').textContent = 'Settings changed. Auto-saving...';
            settingsSaveTimer = setTimeout(() => {
                settingsSaveTimer = undefined;
                settingsSaving = true;
                renderCounts();
                vscode.postMessage({ type: 'save-settings', quiet: true, ...currentSettings() });
            }, 650);
        }

        function refreshSettings() {
            clearSettingsAutosave();
            applying = true;
            el('result').textContent = 'Refreshing...';
            renderCounts();
            vscode.postMessage({ type: 'refresh-settings', autoSaveSettings: true, ...currentSettings() });
        }

        el('selectChanged').addEventListener('click', () => {
            selected.clear();
            for (const item of plan.items) {
                if (isRelevantItem(item)) selected.add(item.id);
            }
            renderList();
        });
        el('clear').addEventListener('click', () => {
            selected.clear();
            renderList();
        });
        el('copyLeft').addEventListener('click', event => {
            event.stopPropagation();
            vscode.postMessage({ type: 'copy-code', itemId: activeId, side: 'left', showHeaders });
        });
        el('copyRight').addEventListener('click', event => {
            event.stopPropagation();
            vscode.postMessage({ type: 'copy-code', itemId: activeId, side: 'right', showHeaders });
        });
        document.addEventListener('contextmenu', event => {
            event.preventDefault();
        });
        el('toggleHeaders').addEventListener('click', () => {
            showHeaders = !showHeaders;
            renderDiff();
        });
        el('chooseFolder').addEventListener('click', () => {
            clearSettingsAutosave();
            applying = true;
            el('result').textContent = 'Choosing folder...';
            renderCounts();
            vscode.postMessage({ type: 'choose-folder', autoSaveSettings: true, ...currentSettings() });
        });
        el('syncMode').addEventListener('change', () => {
            updateModeTitle();
            refreshSettings();
        });
        el('cancel').addEventListener('click', () => {
            clearSettingsAutosave();
            vscode.postMessage({ type: 'cancel' });
        });
        el('apply').addEventListener('click', () => {
            if (applied) return;
            clearSettingsAutosave();
            applying = true;
            el('result').textContent = 'Applying...';
            renderCounts();
            vscode.postMessage({ type: 'apply', selectedIds: Array.from(selected) });
        });

        window.addEventListener('message', event => {
            const message = event.data;
            if (message.type === 'applying') {
                applying = true;
                el('result').textContent = 'Applying...';
                renderCounts();
            } else if (message.type === 'refreshing') {
                applying = true;
                el('result').textContent = message.message || 'Refreshing...';
                renderCounts();
            } else if (message.type === 'ready') {
                applying = false;
                el('result').textContent = '';
                renderCounts();
            } else if (message.type === 'plan') {
                setPlan(message.plan, message.message || 'Settings updated. Review the refreshed diff before applying.', message.autoSaveSettings === true);
            } else if (message.type === 'saving-settings') {
                settingsSaving = true;
                el('result').textContent = 'Saving settings...';
                renderCounts();
            } else if (message.type === 'settings-saved') {
                settingsSaving = false;
                el('result').textContent = message.quiet ? 'Settings auto-saved.' : message.result.summary;
                renderCounts();
            } else if (message.type === 'copied') {
                el('result').textContent = message.side === 'left' ? 'Left code copied.' : 'Right code copied.';
            } else if (message.type === 'applied') {
                applying = false;
                applied = true;
                el('result').textContent = message.result.summary;
                el('apply').textContent = 'Applied';
                renderCounts();
            } else if (message.type === 'error') {
                applying = false;
                settingsSaving = false;
                el('result').textContent = message.error;
                renderCounts();
            }
        });

        selected = selectedFromPlan();
        installSplitter();
        installFastTooltips();
        renderChrome();
        renderList();
        renderDiff();
    </script>
</body>
</html>`;
}

function settingsFromMessage(
    plan: ModuleSyncPlan,
    message: { folderPath?: string; exportMode?: ExportMode; importMode?: ImportMode },
): ModuleSyncSettings {
    const exportModeSource = message.exportMode !== undefined && message.exportMode !== plan.exportMode
        ? 'session'
        : plan.exportModeSource;
    const importModeSource = message.importMode !== undefined && message.importMode !== plan.importMode
        ? 'session'
        : plan.importModeSource;
    return {
        folderPath: message.folderPath ?? plan.folderPath,
        folderPathSource: plan.folderPathSource,
        exportMode: message.exportMode ?? plan.exportMode,
        exportModeSource,
        importMode: message.importMode ?? plan.importMode,
        importModeSource,
        settingsPath: plan.settingsPath,
    };
}

function settingsFromPlan(plan: ModuleSyncPlan): ModuleSyncSettings {
    return {
        folderPath: plan.folderPath,
        folderPathSource: plan.folderPathSource,
        exportMode: plan.exportMode,
        exportModeSource: plan.exportModeSource,
        importMode: plan.importMode,
        importModeSource: plan.importModeSource,
        settingsPath: plan.settingsPath,
    };
}

function isModuleSyncWatchedPath(filePath: string): boolean {
    return /\.(bas|cls)$/i.test(filePath);
}
