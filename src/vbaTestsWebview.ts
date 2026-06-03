import * as path from 'path';
import * as vscode from 'vscode';

export type VbaTestSupportState = 'installed' | 'missing' | 'outdated' | 'blocked' | 'unknown';

export interface VbaTestSupportStatusModel {
    state: VbaTestSupportState;
    title: string;
    description: string;
    actionLabel: string;
    canInstall: boolean;
    canRun: boolean;
}

export type VbaTestRuntimeState = 'installed' | 'missing' | 'blocked' | 'unknown';

export interface VbaTestRuntimeStatusModel {
    state: VbaTestRuntimeState;
    title: string;
    description: string;
    canRun: boolean;
}

export interface VbaTestTagFilterModel {
    name: string;
    testCount: number;
}

export interface VbaTestDiscoveryStatusModel {
    totalTests: number;
    taggedTests: number;
    untaggedTests: number;
    tags: VbaTestTagFilterModel[];
    error?: string;
}

export interface VbaTestsPanelModel {
    filePath: string;
    workbookName: string;
    support: VbaTestSupportStatusModel;
    runtime: VbaTestRuntimeStatusModel;
    discovery: VbaTestDiscoveryStatusModel;
}

export interface VbaTestsRunFilterRequest {
    includeTags: string[];
    excludeTags: string[];
    failFast: boolean;
}

export interface VbaTestsPanelOptions {
    getModel: () => Promise<VbaTestsPanelModel>;
    onInstallSupport?: () => Promise<void>;
    onRunAll?: () => Promise<void>;
    onRunWithFilters?: (request: VbaTestsRunFilterRequest) => Promise<void>;
    onDidChangeWorkbookTree?: vscode.Event<unknown>;
}

interface VbaTestsWebviewMessage {
    type?: string;
    includeTags?: unknown;
    excludeTags?: unknown;
    failFast?: unknown;
}

interface OpenVbaTestsPanelEntry {
    panel: vscode.WebviewPanel;
    options: VbaTestsPanelOptions;
    refresh: () => Promise<void>;
}

const openVbaTestsPanels = new Map<string, OpenVbaTestsPanelEntry>();

export function openVbaTestsPanel(
    context: vscode.ExtensionContext,
    filePath: string,
    options: VbaTestsPanelOptions,
): vscode.WebviewPanel {
    const panelKey = vbaTestsPanelKey(filePath);
    const existing = openVbaTestsPanels.get(panelKey);
    if (existing) {
        existing.options = options;
        existing.panel.reveal(vscode.ViewColumn.Beside);
        void existing.refresh().catch((err) => {
            const error = err instanceof Error ? err.message : String(err);
            void existing.panel.webview.postMessage({ type: 'error', error });
        });
        return existing.panel;
    }

    let disposed = false;
    let refreshTimer: ReturnType<typeof setTimeout> | undefined;
    let refreshVersion = 0;
    const panel = vscode.window.createWebviewPanel(
        'xlideVbaTests',
        `XLIDE Tests: ${path.basename(filePath)}`,
        vscode.ViewColumn.Beside,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
        },
    );
    const entry: OpenVbaTestsPanelEntry = {
        panel,
        options,
        refresh: async () => { /* assigned below */ },
    };

    const renderPanel = async (): Promise<void> => {
        const model = await entry.options.getModel();
        if (disposed) {
            return;
        }
        panel.title = `XLIDE Tests: ${model.workbookName}`;
        panel.webview.html = renderVbaTestsHtml(panel.webview, model);
    };
    entry.refresh = renderPanel;
    openVbaTestsPanels.set(panelKey, entry);

    const refreshPanel = async (requestVersion: number): Promise<void> => {
        if (disposed || requestVersion !== refreshVersion) {
            return;
        }
        await renderPanel();
    };

    const scheduleRefresh = (): void => {
        const requestVersion = ++refreshVersion;
        if (disposed) {
            return;
        }
        if (refreshTimer) {
            clearTimeout(refreshTimer);
        }
        refreshTimer = setTimeout(() => {
            refreshTimer = undefined;
            void refreshPanel(requestVersion).catch((err) => {
                const error = err instanceof Error ? err.message : String(err);
                void panel.webview.postMessage({ type: 'error', error });
            });
        }, 250);
    };

    const runAndRefresh = async (
        operation: (() => Promise<void>) | undefined,
        missingMessage: string,
    ): Promise<void> => {
        if (!operation) {
            await panel.webview.postMessage({ type: 'error', error: missingMessage });
            return;
        }
        await operation();
        await renderPanel();
    };

    void renderPanel().catch((err) => {
        const error = err instanceof Error ? err.message : String(err);
        panel.webview.html = renderVbaTestsErrorHtml(panel.webview, path.basename(filePath), error);
    });

    const messageSub = panel.webview.onDidReceiveMessage(async (message: VbaTestsWebviewMessage) => {
        try {
            if (message.type === 'installSupport') {
                await runAndRefresh(entry.options.onInstallSupport, 'XLIDE test support installation is not available.');
                await panel.webview.postMessage({ type: 'refreshed' });
                return;
            }
            if (message.type === 'runAll') {
                await runAndRefresh(entry.options.onRunAll, 'XLIDE test execution is not available.');
                return;
            }
            if (message.type === 'runWithFilters') {
                const request = runFilterRequestFromMessage(message);
                await runAndRefresh(
                    entry.options.onRunWithFilters
                        ? () => entry.options.onRunWithFilters?.(request) ?? Promise.resolve()
                        : undefined,
                    'XLIDE filtered test execution is not available.',
                );
            }
        } catch (err) {
            const error = err instanceof Error ? err.message : String(err);
            await panel.webview.postMessage({ type: 'error', error });
            await renderPanel().catch(() => { /* keep existing error visible */ });
        }
    });

    const treeSub = entry.options.onDidChangeWorkbookTree?.(() => scheduleRefresh());
    const panelDisposables = [
        messageSub,
        ...(treeSub ? [treeSub] : []),
    ];
    panel.onDidDispose(() => {
        disposed = true;
        openVbaTestsPanels.delete(panelKey);
        if (refreshTimer) {
            clearTimeout(refreshTimer);
            refreshTimer = undefined;
        }
        for (const sub of panelDisposables) {
            sub.dispose();
        }
    });

    context.subscriptions.push(panel);
    return panel;
}

export function renderVbaTestsHtml(
    webviewOrModel: vscode.Webview | VbaTestsPanelModel,
    maybeModel?: VbaTestsPanelModel,
): string {
    const model = maybeModel ?? webviewOrModel as VbaTestsPanelModel;
    const webview = maybeModel ? webviewOrModel as vscode.Webview : undefined;
    const nonce = randomNonce();
    const cspSource = webview?.cspSource ?? 'vscode-resource:';
    const runEnabled = model.support.canRun && model.runtime.canRun;
    const runDisabled = runEnabled ? '' : 'disabled';
    const installDisabled = model.support.canInstall ? '' : 'disabled';
    const installTitle = model.support.canInstall
        ? `${model.support.actionLabel} ${model.workbookName}`
        : model.support.description;
    const runHelp = !model.support.canRun
        ? 'Install the bundled XlideAssert.bas support module before running workbook tests.'
        : !model.runtime.canRun
            ? model.runtime.description
            : '';
    const hasTagFilters = model.discovery.tags.length > 0;
    const filterRunDisabled = runEnabled && hasTagFilters ? '' : 'disabled';
    const filterRunTitle = !runEnabled
        ? runHelp
        : hasTagFilters
            ? 'Run selected tag filters'
            : 'No test tags discovered in this workbook.';
    const tagNamesJson = scriptJson(model.discovery.tags.map((tag) => tag.name));
    const workbookPathJson = scriptJson(model.filePath);
    const canRunJson = JSON.stringify(runEnabled);
    const hasTagsJson = JSON.stringify(hasTagFilters);

    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Tests</title>
    <style>
        body {
            margin: 0;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
            font-size: var(--vscode-font-size);
        }
        .shell {
            max-width: 1040px;
            padding: 22px;
        }
        header {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 16px;
            border-bottom: 1px solid var(--vscode-panel-border);
            padding-bottom: 16px;
        }
        h1 {
            margin: 0 0 4px;
            font-size: 20px;
            font-weight: 700;
        }
        h2 {
            margin: 0;
            font-size: 15px;
            font-weight: 700;
        }
        .subtle {
            color: var(--vscode-descriptionForeground);
            line-height: 1.45;
        }
        button {
            min-height: 32px;
            border: 1px solid var(--vscode-button-border, transparent);
            border-radius: 4px;
            padding: 5px 12px;
            color: var(--vscode-button-foreground);
            background: var(--vscode-button-background);
            font: inherit;
            font-weight: 600;
            cursor: pointer;
        }
        button:hover:not(:disabled) {
            background: var(--vscode-button-hoverBackground);
        }
        button.secondary {
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
        }
        button.secondary:hover:not(:disabled) {
            background: var(--vscode-button-secondaryHoverBackground);
        }
        button:disabled {
            cursor: not-allowed;
            opacity: 0.55;
        }
        main {
            display: grid;
            gap: 16px;
            padding-top: 16px;
        }
        .statusGrid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 16px;
            align-items: stretch;
        }
        section {
            border: 1px solid var(--vscode-panel-border);
            border-radius: 6px;
            background: var(--vscode-sideBar-background);
        }
        .sectionHeader {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            padding: 13px 15px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        .sectionBody {
            padding: 15px;
        }
        .statusRow {
            display: grid;
            grid-template-columns: auto 1fr;
            gap: 12px;
            align-items: start;
        }
        .statusDot {
            width: 11px;
            height: 11px;
            margin-top: 4px;
            border-radius: 50%;
            background: var(--vscode-descriptionForeground);
        }
        .status-installed .statusDot {
            background: var(--vscode-testing-iconPassed, #73c991);
        }
        .status-missing .statusDot,
        .status-outdated .statusDot,
        .status-unknown .statusDot {
            background: var(--vscode-testing-iconQueued, #cca700);
        }
        .status-blocked .statusDot {
            background: var(--vscode-testing-iconErrored, #f14c4c);
        }
        .statusTitle {
            margin-bottom: 3px;
            font-weight: 700;
        }
        .helpText {
            max-width: 760px;
            color: var(--vscode-descriptionForeground);
            line-height: 1.5;
        }
        .runGrid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 10px;
        }
        .runButton {
            width: 100%;
            min-height: 44px;
            text-align: left;
        }
        .filterPanel {
            display: grid;
            gap: 12px;
            margin-top: 14px;
            border-top: 1px solid var(--vscode-panel-border);
            padding-top: 14px;
        }
        .filterHeader,
        .filterColumnHeader,
        .filterActions,
        .checkRow {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .filterHeader {
            justify-content: space-between;
        }
        .filterTitle {
            font-weight: 700;
        }
        .filterSummary {
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
        }
        .filterColumns {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 12px;
        }
        .filterColumn {
            display: grid;
            gap: 8px;
            min-width: 0;
        }
        .filterColumnHeader {
            justify-content: space-between;
            min-height: 30px;
        }
        .filterColumnTitle {
            font-weight: 700;
        }
        .filterTools {
            display: flex;
            gap: 6px;
        }
        button.compact {
            min-height: 28px;
            padding: 3px 8px;
            font-size: 12px;
        }
        .tagList {
            display: grid;
            gap: 6px;
        }
        .tagChoice {
            display: grid;
            grid-template-columns: auto minmax(0, 1fr) auto;
            gap: 9px;
            align-items: center;
            min-height: 32px;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 4px;
            padding: 5px 8px;
            background: var(--vscode-editor-background);
        }
        .tagChoice input,
        .checkRow input {
            accent-color: var(--vscode-checkbox-selectBackground, var(--vscode-button-background));
        }
        .tagName {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-weight: 600;
        }
        .tagCount {
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
        }
        .emptyState {
            border: 1px dashed var(--vscode-panel-border);
            border-radius: 4px;
            padding: 10px;
            color: var(--vscode-descriptionForeground);
        }
        .toast {
            position: fixed;
            right: 18px;
            bottom: 18px;
            max-width: 360px;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 6px;
            padding: 10px 12px;
            color: var(--vscode-notifications-foreground);
            background: var(--vscode-notifications-background);
            box-shadow: 0 8px 22px rgba(0, 0, 0, 0.32);
            opacity: 0;
            transform: translateY(8px);
            transition: opacity 120ms ease, transform 120ms ease;
            pointer-events: none;
        }
        .toast.visible {
            opacity: 1;
            transform: translateY(0);
        }
        @media (max-width: 620px) {
            .shell {
                padding: 16px;
            }
            header,
            .sectionHeader {
                display: grid;
            }
        }
    </style>
</head>
<body>
    <div class="shell">
        <header>
            <div>
                <h1>XLIDE Unit Tests</h1>
                <div class="subtle">${escapeHtml(model.workbookName)}</div>
            </div>
        </header>
        <main>
            <div class="statusGrid">
                <section class="status-${escapeAttr(model.support.state)}">
                    <div class="sectionHeader">
                        <h2>Test Support</h2>
                        <button type="button" data-action="installSupport" title="${escapeAttr(installTitle)}" ${installDisabled}>${escapeHtml(model.support.actionLabel)}</button>
                    </div>
                    <div class="sectionBody">
                        <div class="statusRow">
                            <span class="statusDot"></span>
                            <div>
                                <div class="statusTitle">${escapeHtml(model.support.title)}</div>
                                <div class="helpText">${escapeHtml(model.support.description)}</div>
                            </div>
                        </div>
                    </div>
                </section>
                <section class="status-${escapeAttr(model.runtime.state)}">
                    <div class="sectionHeader">
                        <h2>Runtime</h2>
                    </div>
                    <div class="sectionBody">
                        <div class="statusRow">
                            <span class="statusDot"></span>
                            <div>
                                <div class="statusTitle">${escapeHtml(model.runtime.title)}</div>
                                <div class="helpText">${escapeHtml(model.runtime.description)}</div>
                            </div>
                        </div>
                    </div>
                </section>
            </div>
            <section>
                <div class="sectionHeader">
                    <h2>Run</h2>
                </div>
                <div class="sectionBody">
                    <div class="runGrid">
                        <button class="runButton" type="button" data-action="runAll" ${runDisabled}>Run All Tests</button>
                        <button class="runButton" type="button" data-action="runWithFilters" title="${escapeAttr(filterRunTitle)}" ${filterRunDisabled}>Run With Filters</button>
                    </div>
                    ${runHelp ? `<p class="helpText">${escapeHtml(runHelp)}</p>` : ''}
                    ${renderTagFilters(model)}
                </div>
            </section>
        </main>
    </div>
    <div class="toast" id="toast"></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const toast = document.getElementById('toast');
        const workbookPath = ${workbookPathJson};
        const tagNames = ${tagNamesJson};
        const canRun = ${canRunJson};
        const hasTags = ${hasTagsJson};
        let toastTimer;
        let filterState = initialFilterState();
        let running = false;

        function showToast(message) {
            toast.textContent = message;
            toast.classList.add('visible');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => toast.classList.remove('visible'), 2600);
        }

        function initialFilterState() {
            const saved = vscode.getState?.();
            if (saved?.workbookPath === workbookPath) {
                return {
                    workbookPath,
                    includeTags: reconcileTags(saved.includeTags),
                    excludeTags: reconcileTags(saved.excludeTags),
                    failFast: Boolean(saved.failFast),
                };
            }
            return {
                workbookPath,
                includeTags: [...tagNames],
                excludeTags: [],
                failFast: false,
            };
        }

        function reconcileTags(value) {
            if (!Array.isArray(value)) {
                return [];
            }
            const available = new Set(tagNames);
            return [...new Set(value.map((tag) => String(tag).trim()).filter((tag) => available.has(tag)))];
        }

        function saveFilterState() {
            vscode.setState?.(filterState);
        }

        function setRunning(next) {
            running = next;
            syncFilterUi();
        }

        function syncFilterUi() {
            document.querySelectorAll('input[data-filter-kind]').forEach((input) => {
                const list = input.dataset.filterKind === 'exclude' ? filterState.excludeTags : filterState.includeTags;
                input.checked = list.includes(input.dataset.tag);
            });
            const failFast = document.getElementById('failFast');
            if (failFast) {
                failFast.checked = filterState.failFast;
            }
            const summary = document.getElementById('filterSummary');
            if (summary) {
                const includeCount = filterState.includeTags.length;
                const excludeCount = filterState.excludeTags.length;
                summary.textContent = includeCount + ' include, ' + excludeCount + ' exclude';
            }
            document.querySelectorAll('[data-filter-action], input[data-filter-kind], #failFast').forEach((control) => {
                control.disabled = running;
            });
            const runAll = document.querySelector('button[data-action="runAll"]');
            if (runAll && canRun) {
                runAll.disabled = running;
            }
            const runWithFilters = document.querySelector('button[data-action="runWithFilters"]');
            if (runWithFilters && canRun && hasTags) {
                runWithFilters.disabled = running ||
                    (filterState.includeTags.length === 0 && filterState.excludeTags.length === 0);
            }
        }

        function setFilterTags(kind, values) {
            const selected = [...new Set(values.filter((tag) => tagNames.includes(tag)))];
            if (kind === 'exclude') {
                filterState.excludeTags = selected;
                filterState.includeTags = filterState.includeTags.filter((tag) => !selected.includes(tag));
            } else {
                filterState.includeTags = selected;
                filterState.excludeTags = filterState.excludeTags.filter((tag) => !selected.includes(tag));
            }
            saveFilterState();
            syncFilterUi();
        }

        function toggleFilterTag(kind, tag, checked) {
            const list = kind === 'exclude' ? filterState.excludeTags : filterState.includeTags;
            const next = checked
                ? [...new Set([...list, tag])]
                : list.filter((candidate) => candidate !== tag);
            setFilterTags(kind, next);
        }

        document.addEventListener('change', (event) => {
            const input = event.target.closest?.('input[data-filter-kind]');
            if (input) {
                toggleFilterTag(input.dataset.filterKind, input.dataset.tag, input.checked);
                return;
            }
            if (event.target?.id === 'failFast') {
                filterState.failFast = Boolean(event.target.checked);
                saveFilterState();
            }
        });

        document.addEventListener('click', (event) => {
            const filterButton = event.target.closest?.('[data-filter-action]');
            if (filterButton) {
                const action = filterButton.dataset.filterAction;
                const kind = filterButton.dataset.filterKind;
                if (action === 'selectAll') {
                    setFilterTags(kind, tagNames);
                } else if (action === 'clear') {
                    setFilterTags(kind, []);
                }
                return;
            }
            const button = event.target.closest?.('[data-action]');
            if (!button || button.disabled) {
                return;
            }
            if (button.dataset.action === 'runAll') {
                setRunning(true);
                vscode.postMessage({ type: 'runAll' });
                return;
            }
            if (button.dataset.action === 'runWithFilters') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runWithFilters',
                    includeTags: filterState.includeTags,
                    excludeTags: filterState.excludeTags,
                    failFast: filterState.failFast,
                });
                return;
            }
            vscode.postMessage({ type: button.dataset.action });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'error') {
                setRunning(false);
                showToast(event.data.error || 'XLIDE test action failed');
            } else if (event.data?.type === 'refreshed') {
                setRunning(false);
                showToast('Test support refreshed');
            }
        });

        syncFilterUi();
    </script>
</body>
</html>`;
}

function renderVbaTestsErrorHtml(webview: vscode.Webview, workbookName: string, error: string): string {
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Tests Error</title>
    <style>
        body {
            margin: 0;
            padding: 24px;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
        }
        h1 {
            margin: 0 0 6px;
            font-size: 18px;
        }
        .subtle {
            color: var(--vscode-descriptionForeground);
            margin-bottom: 18px;
        }
        .error {
            border: 1px solid var(--vscode-inputValidation-errorBorder);
            border-radius: 4px;
            padding: 12px;
            white-space: pre-wrap;
        }
    </style>
</head>
<body>
    <main>
        <h1>XLIDE Tests Could Not Load</h1>
        <div class="subtle">${escapeHtml(workbookName)}</div>
        <div class="error">${escapeHtml(error)}</div>
    </main>
</body>
</html>`;
}

function renderTagFilters(model: VbaTestsPanelModel): string {
    const discovery = model.discovery;
    const tagSummary = discovery.error
        ? escapeHtml(discovery.error)
        : `${discovery.tags.length} tags, ${discovery.taggedTests} tagged tests`;
    return /* html */`<div class="filterPanel">
        <div class="filterHeader">
            <div>
                <div class="filterTitle">Tag Filters</div>
                <div class="filterSummary">${tagSummary}</div>
            </div>
            <label class="checkRow" title="Stop after the first failure, unexpected pass, timeout, or host error.">
                <input id="failFast" type="checkbox">
                <span>Fail Fast</span>
            </label>
        </div>
        ${discovery.tags.length > 0 ? /* html */`
            <div class="filterColumns">
                ${renderTagFilterColumn('include', 'Include Tags', discovery.tags, true)}
                ${renderTagFilterColumn('exclude', 'Exclude Tags', discovery.tags, false)}
            </div>
            <div class="filterSummary" id="filterSummary"></div>
        ` : /* html */`
            <div class="emptyState">No test tags discovered</div>
        `}
    </div>`;
}

function renderTagFilterColumn(
    kind: 'include' | 'exclude',
    title: string,
    tags: readonly VbaTestTagFilterModel[],
    checked: boolean,
): string {
    return /* html */`<div class="filterColumn">
        <div class="filterColumnHeader">
            <span class="filterColumnTitle">${escapeHtml(title)}</span>
            <span class="filterTools">
                <button class="secondary compact" type="button" data-filter-action="selectAll" data-filter-kind="${kind}">Select All</button>
                <button class="secondary compact" type="button" data-filter-action="clear" data-filter-kind="${kind}">Clear</button>
            </span>
        </div>
        <div class="tagList">
            ${tags.map((tag) => renderTagFilterChoice(kind, tag, checked)).join('')}
        </div>
    </div>`;
}

function renderTagFilterChoice(
    kind: 'include' | 'exclude',
    tag: VbaTestTagFilterModel,
    checked: boolean,
): string {
    return /* html */`<label class="tagChoice">
        <input type="checkbox" data-filter-kind="${kind}" data-tag="${escapeAttr(tag.name)}" ${checked ? 'checked' : ''}>
        <span class="tagName">${escapeHtml(tag.name)}</span>
        <span class="tagCount">${tag.testCount}</span>
    </label>`;
}

function runFilterRequestFromMessage(message: VbaTestsWebviewMessage): VbaTestsRunFilterRequest {
    return {
        includeTags: stringListFromUnknown(message.includeTags),
        excludeTags: stringListFromUnknown(message.excludeTags),
        failFast: Boolean(message.failFast),
    };
}

function stringListFromUnknown(value: unknown): string[] {
    if (!Array.isArray(value)) {
        return [];
    }
    const seen = new Set<string>();
    const result: string[] = [];
    for (const item of value) {
        const tag = String(item).trim();
        const key = tag.toLowerCase();
        if (!tag || seen.has(key)) {
            continue;
        }
        seen.add(key);
        result.push(tag);
    }
    return result;
}

function randomNonce(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let nonce = '';
    for (let i = 0; i < 32; i++) {
        nonce += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return nonce;
}

function escapeHtml(value: unknown): string {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function escapeAttr(value: unknown): string {
    return escapeHtml(value);
}

function scriptJson(value: unknown): string {
    return JSON.stringify(value).replace(/</g, '\\u003c');
}

function vbaTestsPanelKey(filePath: string): string {
    const normalized = path.normalize(filePath);
    return process.platform === 'win32' ? normalized.toLowerCase() : normalized;
}
