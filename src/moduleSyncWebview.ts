import * as vscode from 'vscode';
import type { ModuleSyncPlan } from './moduleSyncPlan';
import type { ExportMode } from './moduleExport';

export interface ModuleSyncApplyResult {
    summary: string;
    changed: number;
    skipped: number;
    removed?: number;
    failed: number;
}

export interface ModuleSyncSettings {
    folderPath: string;
    exportMode?: ExportMode;
}

export interface ModuleSyncPreviewOptions {
    onChooseFolder?: (current: ModuleSyncSettings) => Promise<ModuleSyncPlan | undefined>;
    onRefresh?: (settings: ModuleSyncSettings) => Promise<ModuleSyncPlan>;
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
        let operationInFlight = false;
        let queuedFolderRefresh = false;
        let folderRefreshTimer: ReturnType<typeof setTimeout> | undefined;
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
        const queueFolderRefresh = (delayMs = 300): void => {
            if (!options.onRefresh || disposed) {
                return;
            }
            if (operationInFlight) {
                queuedFolderRefresh = true;
                return;
            }
            if (folderRefreshTimer) {
                clearTimeout(folderRefreshTimer);
            }
            folderRefreshTimer = setTimeout(() => {
                folderRefreshTimer = undefined;
                void refreshPlanFromDisk();
            }, delayMs);
        };
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
                    queueFolderRefresh();
                }
            };
            folderWatcherDisposables = [
                watcher.onDidChange(refreshIfRelevant),
                watcher.onDidCreate(refreshIfRelevant),
                watcher.onDidDelete(refreshIfRelevant),
                watcher,
            ];
        };
        const runExclusive = async <T>(operation: () => Promise<T>): Promise<T> => {
            operationInFlight = true;
            try {
                return await operation();
            } finally {
                operationInFlight = false;
                if (queuedFolderRefresh && !disposed) {
                    queuedFolderRefresh = false;
                    queueFolderRefresh(100);
                }
            }
        };
        const updateCurrentPlan = (nextPlan: ModuleSyncPlan): void => {
            currentPlan = nextPlan;
            configureFolderWatcher();
        };
        async function refreshPlanFromDisk(): Promise<void> {
            const refresh = options.onRefresh;
            if (!refresh || disposed) {
                return;
            }
            if (operationInFlight) {
                queuedFolderRefresh = true;
                return;
            }
            try {
                await panel.webview.postMessage({ type: 'refreshing', message: 'Disk changes detected. Refreshing preview...' });
                await runExclusive(async () => {
                    updateCurrentPlan(await refresh(settingsFromPlan(currentPlan)));
                    await panel.webview.postMessage({
                        type: 'plan',
                        plan: currentPlan,
                        message: 'Disk changes detected. Preview refreshed.',
                    });
                });
            } catch (err) {
                const error = err instanceof Error ? err.message : String(err);
                await panel.webview.postMessage({ type: 'error', error });
            }
        }
        configureFolderWatcher();
        panel.webview.html = renderModuleSyncHtml(panel.webview, currentPlan);
        const messageSub = panel.webview.onDidReceiveMessage(async (message: {
            type?: string;
            selectedIds?: string[];
            folderPath?: string;
            exportMode?: ExportMode;
        }) => {
            if (message.type === 'cancel') {
                done(undefined);
                panel.dispose();
                return;
            }
            if (message.type === 'choose-folder') {
                const chooseFolder = options.onChooseFolder;
                if (!chooseFolder) {
                    return;
                }
                await panel.webview.postMessage({ type: 'refreshing' });
                try {
                    await runExclusive(async () => {
                        const nextPlan = await chooseFolder(settingsFromMessage(currentPlan, message));
                        if (nextPlan) {
                            updateCurrentPlan(nextPlan);
                            await panel.webview.postMessage({
                                type: 'plan',
                                plan: currentPlan,
                                message: 'Settings updated. Review the refreshed diff before applying.',
                            });
                        } else {
                            await panel.webview.postMessage({ type: 'ready' });
                        }
                    });
                } catch (err) {
                    const error = err instanceof Error ? err.message : String(err);
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
                    await runExclusive(async () => {
                        updateCurrentPlan(await refresh(settingsFromMessage(currentPlan, message)));
                        await panel.webview.postMessage({
                            type: 'plan',
                            plan: currentPlan,
                            message: 'Settings updated. Review the refreshed diff before applying.',
                        });
                    });
                } catch (err) {
                    const error = err instanceof Error ? err.message : String(err);
                    await panel.webview.postMessage({ type: 'error', error });
                }
                return;
            }
            if (message.type === 'save-settings') {
                const saveSettings = options.onSaveSettings;
                if (!saveSettings) {
                    return;
                }
                await panel.webview.postMessage({ type: 'saving-settings' });
                try {
                    await runExclusive(async () => {
                        const result = await saveSettings(settingsFromMessage(currentPlan, message));
                        await panel.webview.postMessage({ type: 'settings-saved', result });
                    });
                } catch (err) {
                    const error = err instanceof Error ? err.message : String(err);
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
                await runExclusive(async () => {
                    const result = await onApply(currentPlan, selectedIds);
                    await panel.webview.postMessage({ type: 'applied', result });
                    done(result);
                });
            } catch (err) {
                const error = err instanceof Error ? err.message : String(err);
                await panel.webview.postMessage({ type: 'error', error });
            }
        });
        panel.onDidDispose(() => {
            disposed = true;
            if (folderRefreshTimer) {
                clearTimeout(folderRefreshTimer);
                folderRefreshTimer = undefined;
            }
            disposeFolderWatcher();
            messageSub.dispose();
            done(undefined);
        });
    });
}

function renderModuleSyncHtml(webview: vscode.Webview, plan: ModuleSyncPlan): string {
    const nonce = getNonce();
    const data = JSON.stringify(plan).replace(/</g, '\\u003c');
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtml(plan.title)}</title>
    <style>
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
        .toolbar button { flex: 1; }
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
            grid-template-columns: 28px minmax(0, 1fr) 132px;
            gap: 8px;
            align-items: center;
            padding: 8px 10px;
            border-bottom: 1px solid var(--border);
            cursor: pointer;
        }
        .item:hover, .item.active { background: var(--row); }
        .item input {
            margin: 0;
            width: 18px;
            height: 18px;
            justify-self: center;
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
        .badge.skip { color: var(--vscode-editorWarning-foreground); background: var(--warn); }
        .badge.same { color: var(--muted); }
        section { min-height: 0; display: grid; grid-template-rows: auto 1fr; }
        .diff-head {
            padding: 10px 12px;
            border-bottom: 1px solid var(--border);
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }
        .diff-title { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .diff { overflow: auto; font-family: var(--vscode-editor-font-family); font-size: var(--vscode-editor-font-size); }
        .line {
            display: grid;
            grid-template-columns: 56px minmax(0, 1fr) 56px minmax(0, 1fr);
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
        .line.added pre.right { background: var(--added); }
        .line.removed pre.left { background: var(--removed); }
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
                    </div>
                    <button class="secondary" id="chooseFolder">Change</button>
                    <div class="field" id="modeField">
                        <label for="exportMode">Export mode</label>
                        <select id="exportMode">
                            <option value="trueUp">True Up</option>
                            <option value="replaceExistingOnly">Replace Existing Only</option>
                        </select>
                    </div>
                </div>
            </div>
            <div class="actions">
                <button class="secondary" id="saveSettings">Save Settings</button>
                <button class="secondary" id="cancel">Cancel</button>
                <button id="apply">Apply Selected</button>
            </div>
        </header>
        <main id="layout">
            <aside>
                <div class="toolbar">
                    <button class="secondary" id="selectChanged">Select Changed</button>
                    <button class="secondary" id="clear">Clear</button>
                </div>
                <div class="warnings" id="warnings"></div>
                <div class="list" id="list"></div>
            </aside>
            <div class="splitter" id="splitter" role="separator" aria-orientation="vertical" aria-label="Resize module list"></div>
            <section>
                <div class="diff-head">
                    <div class="diff-title" id="leftTitle"></div>
                    <div class="diff-title" id="rightTitle"></div>
                </div>
                <div class="diff" id="diff"></div>
            </section>
        </main>
        <footer>
            <span id="counts"></span>
            <span class="result" id="result"></span>
        </footer>
    </div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        let plan = ${data};
        let selected = new Set();
        let activeId = plan.items[0]?.id;
        let applying = false;
        let applied = false;

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

        function currentSettings() {
            return {
                folderPath: plan.folderPath,
                exportMode: plan.direction === 'export' ? el('exportMode').value : undefined,
            };
        }

        function renderChrome() {
            el('title').textContent = plan.title;
            el('subtitle').textContent = \`\${plan.workbookPath} <-> \${plan.folderPath}\${plan.exportMode ? '  [' + plan.exportMode + ']' : ''}\`;
            el('folderLabel').textContent = plan.direction === 'export' ? 'Export folder' : 'Import folder';
            el('folderValue').textContent = plan.folderPath;
            el('folderValue').title = plan.folderPath;
            el('modeField').classList.toggle('hidden', plan.direction !== 'export');
            if (plan.direction === 'export') {
                el('exportMode').value = plan.exportMode || 'trueUp';
            }
            el('selectChanged').textContent = plan.direction === 'import' ? 'Select Importable' : 'Select Changed';
            if (plan.warnings.length) {
                el('warnings').classList.add('visible');
                el('warnings').textContent = plan.warnings.join('\\n');
            } else {
                el('warnings').classList.remove('visible');
                el('warnings').textContent = '';
            }
        }

        function setPlan(nextPlan, message) {
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
        }

        function renderList() {
            const list = el('list');
            list.innerHTML = '';
            for (const item of plan.items) {
                const row = document.createElement('div');
                row.className = 'item' + (item.id === activeId ? ' active' : '');
                row.dataset.id = item.id;
                row.setAttribute('aria-selected', item.id === activeId ? 'true' : 'false');
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.checked = selected.has(item.id);
                checkbox.disabled = !item.selectable;
                checkbox.addEventListener('click', event => {
                    event.stopPropagation();
                    activeId = item.id;
                    if (checkbox.checked) selected.add(item.id);
                    else selected.delete(item.id);
                    renderList();
                    renderDiff();
                });
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
                badge.className = 'badge' + (item.status.startsWith('skipping') ? ' skip' : item.status === 'unchanged' ? ' same' : '');
                badge.textContent = item.detail || item.status;
                row.append(checkbox, text, badge);
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
                const empty = document.createElement('pre');
                empty.textContent = 'No module differences found for the current settings.';
                diff.append(empty);
                return;
            }
            el('leftTitle').textContent = item.leftTitle;
            el('rightTitle').textContent = item.rightTitle;
            for (const line of item.diff) {
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
            el('counts').textContent = \`\${selectedItems.length} selected\${unsupported ? ' | ' + unsupported + ' will show skipping import warning' : ''}\`;
            el('apply').disabled = applying || applied || selectedItems.length === 0;
            el('saveSettings').disabled = applying;
            el('chooseFolder').disabled = applying;
            el('exportMode').disabled = applying;
        }

        function refreshSettings() {
            applying = true;
            el('result').textContent = 'Refreshing...';
            renderCounts();
            vscode.postMessage({ type: 'refresh-settings', ...currentSettings() });
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
        el('chooseFolder').addEventListener('click', () => {
            applying = true;
            el('result').textContent = 'Choosing folder...';
            renderCounts();
            vscode.postMessage({ type: 'choose-folder', ...currentSettings() });
        });
        el('exportMode').addEventListener('change', () => refreshSettings());
        el('saveSettings').addEventListener('click', () => {
            applying = true;
            el('result').textContent = 'Saving settings...';
            renderCounts();
            vscode.postMessage({ type: 'save-settings', ...currentSettings() });
        });
        el('cancel').addEventListener('click', () => vscode.postMessage({ type: 'cancel' }));
        el('apply').addEventListener('click', () => {
            if (applied) return;
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
                setPlan(message.plan, message.message || 'Settings updated. Review the refreshed diff before applying.');
            } else if (message.type === 'saving-settings') {
                applying = true;
                el('result').textContent = 'Saving settings...';
                renderCounts();
            } else if (message.type === 'settings-saved') {
                applying = false;
                el('result').textContent = message.result.summary;
                renderCounts();
            } else if (message.type === 'applied') {
                applying = false;
                applied = true;
                el('result').textContent = message.result.summary;
                el('apply').textContent = 'Applied';
                renderCounts();
            } else if (message.type === 'error') {
                applying = false;
                el('result').textContent = message.error;
                renderCounts();
            }
        });

        selected = selectedFromPlan();
        installSplitter();
        renderChrome();
        renderList();
        renderDiff();
    </script>
</body>
</html>`;
}

function getNonce(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let text = '';
    for (let i = 0; i < 32; i++) {
        text += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return text;
}

function escapeHtml(value: string): string {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function settingsFromMessage(
    plan: ModuleSyncPlan,
    message: { folderPath?: string; exportMode?: ExportMode },
): ModuleSyncSettings {
    return {
        folderPath: message.folderPath ?? plan.folderPath,
        exportMode: message.exportMode ?? plan.exportMode,
    };
}

function settingsFromPlan(plan: ModuleSyncPlan): ModuleSyncSettings {
    return {
        folderPath: plan.folderPath,
        exportMode: plan.exportMode,
    };
}

function isModuleSyncWatchedPath(filePath: string): boolean {
    return /\.(bas|cls|frm)$/i.test(filePath);
}
