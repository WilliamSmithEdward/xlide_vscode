import * as path from 'path';
import * as vscode from 'vscode';
import { measurePerformance } from './performanceTrace';
import { escapeAttr, escapeHtml, randomNonce, scriptJson } from './webview/html';
import { errorMessage } from './util/errors';

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

export interface VbaTestListItemModel {
    id: string;
    qualifiedName: string;
    moduleName: string;
    procedureName: string;
    line: number;
    tags: string[];
}

export interface VbaTestDiscoveryStatusModel {
    totalTests: number;
    taggedTests: number;
    untaggedTests: number;
    tags: VbaTestTagFilterModel[];
    tests: VbaTestListItemModel[];
    error?: string;
}

export interface VbaTestLastFailedModel {
    count: number;
    tests: Array<{
        id: string;
        qualifiedName: string;
        status: string;
    }>;
}

export interface VbaTestsPanelModel {
    filePath: string;
    workbookName: string;
    support: VbaTestSupportStatusModel;
    runtime: VbaTestRuntimeStatusModel;
    discovery: VbaTestDiscoveryStatusModel;
    lastFailed?: VbaTestLastFailedModel;
}

export interface VbaTestsRunFilterRequest {
    includeTags: string[];
    excludeTags: string[];
    failFast: boolean;
}

export interface VbaTestsRunSelectedRequest {
    testIds: string[];
    failFast: boolean;
}

export interface VbaTestsRunModeRequest {
    failFast: boolean;
}

export interface VbaTestsPanelOptions {
    getModel: () => Promise<VbaTestsPanelModel>;
    onInstallSupport?: () => Promise<void>;
    onRunAll?: () => Promise<void>;
    onRunWithFilters?: (request: VbaTestsRunFilterRequest) => Promise<void>;
    onRunSelected?: (request: VbaTestsRunSelectedRequest) => Promise<void>;
    onRunCurrentModule?: (request: VbaTestsRunModeRequest) => Promise<void>;
    onRunCurrentTest?: (request: VbaTestsRunModeRequest) => Promise<void>;
    onRerunFailed?: () => Promise<void>;
    onDidChangeWorkbookTree?: vscode.Event<unknown>;
}

interface VbaTestsWebviewMessage {
    type?: string;
    includeTags?: unknown;
    excludeTags?: unknown;
    testIds?: unknown;
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
        existing.panel.reveal(vscode.ViewColumn.Active);
        void existing.refresh().catch((err) => {
            const error = errorMessage(err);
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
        vscode.ViewColumn.Active,
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
        await measurePerformance('vbaTests.renderPanel', path.basename(filePath), async () => {
        const model = await entry.options.getModel();
        if (disposed) {
            return;
        }
        panel.title = `XLIDE Tests: ${model.workbookName}`;
        panel.webview.html = renderVbaTestsHtml(model);
        });
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
                const error = errorMessage(err);
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
        const error = errorMessage(err);
        panel.webview.html = renderVbaTestsErrorHtml(path.basename(filePath), error);
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
                return;
            }
            if (message.type === 'runSelected') {
                const request = runSelectedRequestFromMessage(message);
                await runAndRefresh(
                    entry.options.onRunSelected
                        ? () => entry.options.onRunSelected?.(request) ?? Promise.resolve()
                        : undefined,
                    'XLIDE selected test execution is not available.',
                );
                return;
            }
            if (message.type === 'runCurrentModule') {
                const request = runModeRequestFromMessage(message);
                await runAndRefresh(
                    entry.options.onRunCurrentModule
                        ? () => entry.options.onRunCurrentModule?.(request) ?? Promise.resolve()
                        : undefined,
                    'XLIDE current-module test execution is not available.',
                );
                return;
            }
            if (message.type === 'runCurrentTest') {
                const request = runModeRequestFromMessage(message);
                await runAndRefresh(
                    entry.options.onRunCurrentTest
                        ? () => entry.options.onRunCurrentTest?.(request) ?? Promise.resolve()
                        : undefined,
                    'XLIDE current-test execution is not available.',
                );
                return;
            }
            if (message.type === 'rerunFailed') {
                await runAndRefresh(entry.options.onRerunFailed, 'XLIDE rerun failed is not available.');
            }
        } catch (err) {
            const error = errorMessage(err);
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

export function renderVbaTestsHtml(model: VbaTestsPanelModel): string {
    const nonce = randomNonce();
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
    const failedCount = model.lastFailed?.count ?? 0;
    const hasLastFailed = failedCount > 0;
    const hasTests = model.discovery.tests.length > 0;
    const filterRunDisabled = runEnabled && hasTagFilters ? '' : 'disabled';
    const selectedRunDisabled = runEnabled && hasTests ? '' : 'disabled';
    const rerunFailedDisabled = runEnabled && hasLastFailed ? '' : 'disabled';
    const filterRunTitle = !runEnabled
        ? runHelp
        : hasTagFilters
            ? 'Run selected tag filters'
            : 'No test tags discovered in this workbook.';
    const selectedRunTitle = !runEnabled
        ? runHelp
        : hasTests
            ? 'Run the checked tests below'
            : 'No tests discovered in this workbook.';
    const currentScopeTitle = !runEnabled
        ? runHelp
        : `Use the active editor if it belongs to ${model.workbookName}.`;
    const rerunFailedTitle = !runEnabled
        ? runHelp
        : hasLastFailed
            ? `Rerun ${failedCount} failed, timed out, host-error, or unexpected-pass test${failedCount === 1 ? '' : 's'} from the last run.`
            : 'No failed tests from the last run.';
    const tagNamesJson = scriptJson(model.discovery.tags.map((tag) => tag.name));
    const testIdsJson = scriptJson(model.discovery.tests.map((test) => test.id));
    const workbookPathJson = scriptJson(model.filePath);
    const canRunJson = JSON.stringify(runEnabled);
    const hasTagsJson = JSON.stringify(hasTagFilters);
    const hasTestsJson = JSON.stringify(hasTests);
    const hasLastFailedJson = JSON.stringify(hasLastFailed);

    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Tests</title>
    <style nonce="${nonce}">
        :root {
            --xlide-accent-blue: #2d5f94;
            --xlide-accent-blue-hover: #376fa8;
            --xlide-accent-background: color-mix(in srgb, var(--xlide-accent-blue) 82%, var(--vscode-editor-background));
            --xlide-accent-hover-background: color-mix(in srgb, var(--xlide-accent-blue-hover) 84%, var(--vscode-editor-background));
            --xlide-accent-border: color-mix(in srgb, var(--xlide-accent-blue) 78%, var(--vscode-panel-border));
        }
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
            background: var(--xlide-accent-background);
            font: inherit;
            font-weight: 600;
            cursor: pointer;
        }
        button:hover:not(:disabled) {
            background: var(--xlide-accent-hover-background);
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
            border-color: var(--xlide-accent-border);
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
        .testListHeader,
        .checkRow {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .filterHeader {
            justify-content: space-between;
        }
        .testListHeader {
            justify-content: space-between;
            min-height: 30px;
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
        .testList {
            display: grid;
            gap: 6px;
            max-height: 290px;
            overflow: auto;
            padding-right: 2px;
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
        .testChoice {
            display: grid;
            grid-template-columns: auto minmax(180px, 1fr) minmax(120px, 0.55fr);
            gap: 10px;
            align-items: center;
            min-height: 32px;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 4px;
            padding: 7px 8px;
            background: var(--vscode-editor-background);
            min-width: 0;
        }
        .tagChoice input,
        .testChoice input,
        .checkRow input {
            accent-color: var(--vscode-checkbox-selectBackground, var(--xlide-accent-background));
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
        .testName {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-weight: 700;
        }
        .testMeta,
        .testTags {
            overflow: hidden;
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .testTags {
            text-align: right;
        }
        .testTag {
            display: inline-block;
            margin-left: 4px;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 999px;
            padding: 1px 6px;
            color: var(--vscode-descriptionForeground);
            background: color-mix(in srgb, var(--vscode-sideBar-background) 80%, var(--vscode-button-secondaryBackground));
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
            .testChoice {
                grid-template-columns: auto minmax(0, 1fr);
            }
            .testTags {
                grid-column: 2;
                text-align: left;
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
                        <button class="runButton" type="button" data-action="runSelected" title="${escapeAttr(selectedRunTitle)}" ${selectedRunDisabled}>Run Selected</button>
                        <button class="runButton" type="button" data-action="runCurrentModule" title="${escapeAttr(currentScopeTitle)}" ${runDisabled}>Run Current Module</button>
                        <button class="runButton" type="button" data-action="runCurrentTest" title="${escapeAttr(currentScopeTitle)}" ${runDisabled}>Run Current Test</button>
                        <button class="runButton" type="button" data-action="runWithFilters" title="${escapeAttr(filterRunTitle)}" ${filterRunDisabled}>Run With Filters</button>
                        <button class="runButton" type="button" data-action="rerunFailed" title="${escapeAttr(rerunFailedTitle)}" ${rerunFailedDisabled}>Rerun Failed${hasLastFailed ? ` (${failedCount})` : ''}</button>
                    </div>
                    ${runHelp ? `<p class="helpText">${escapeHtml(runHelp)}</p>` : ''}
                    ${renderTestSelection(model)}
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
        const testIds = ${testIdsJson};
        const canRun = ${canRunJson};
        const hasTags = ${hasTagsJson};
        const hasTests = ${hasTestsJson};
        const hasLastFailed = ${hasLastFailedJson};
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
                    selectedTestIds: Array.isArray(saved.selectedTestIds)
                        ? reconcileTestIds(saved.selectedTestIds)
                        : [...testIds],
                    failFast: Boolean(saved.failFast),
                };
            }
            return {
                workbookPath,
                includeTags: [...tagNames],
                excludeTags: [],
                selectedTestIds: [...testIds],
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

        function reconcileTestIds(value) {
            if (!Array.isArray(value)) {
                return [];
            }
            const available = new Set(testIds);
            return [...new Set(value.map((id) => String(id).trim()).filter((id) => available.has(id)))];
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
            document.querySelectorAll('input[data-test-id]').forEach((input) => {
                input.checked = filterState.selectedTestIds.includes(input.dataset.testId);
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
            const selectedSummary = document.getElementById('selectedTestSummary');
            if (selectedSummary) {
                const count = filterState.selectedTestIds.length;
                selectedSummary.textContent = count + ' selected, ' + testIds.length + ' discovered';
            }
            document.querySelectorAll('[data-filter-action], [data-test-action], input[data-filter-kind], input[data-test-id], #failFast').forEach((control) => {
                control.disabled = running;
            });
            const runAll = document.querySelector('button[data-action="runAll"]');
            if (runAll && canRun) {
                runAll.disabled = running;
            }
            const runSelected = document.querySelector('button[data-action="runSelected"]');
            if (runSelected && canRun && hasTests) {
                runSelected.disabled = running || filterState.selectedTestIds.length === 0;
            }
            document.querySelectorAll('button[data-action="runCurrentModule"], button[data-action="runCurrentTest"]').forEach((button) => {
                if (canRun) {
                    button.disabled = running;
                }
            });
            const runWithFilters = document.querySelector('button[data-action="runWithFilters"]');
            if (runWithFilters && canRun && hasTags) {
                runWithFilters.disabled = running ||
                    (filterState.includeTags.length === 0 && filterState.excludeTags.length === 0);
            }
            const rerunFailed = document.querySelector('button[data-action="rerunFailed"]');
            if (rerunFailed && canRun && hasLastFailed) {
                rerunFailed.disabled = running;
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

        function setSelectedTestIds(values) {
            filterState.selectedTestIds = reconcileTestIds(values);
            saveFilterState();
            syncFilterUi();
        }

        function toggleSelectedTest(id, checked) {
            const next = checked
                ? [...new Set([...filterState.selectedTestIds, id])]
                : filterState.selectedTestIds.filter((candidate) => candidate !== id);
            setSelectedTestIds(next);
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
                return;
            }
            const testInput = event.target.closest?.('input[data-test-id]');
            if (testInput) {
                toggleSelectedTest(testInput.dataset.testId, testInput.checked);
            }
        });

        document.addEventListener('click', (event) => {
            const testButton = event.target.closest?.('[data-test-action]');
            if (testButton) {
                if (testButton.dataset.testAction === 'selectAll') {
                    setSelectedTestIds(testIds);
                } else if (testButton.dataset.testAction === 'clear') {
                    setSelectedTestIds([]);
                }
                return;
            }
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
            if (button.dataset.action === 'runSelected') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runSelected',
                    testIds: filterState.selectedTestIds,
                    failFast: filterState.failFast,
                });
                return;
            }
            if (button.dataset.action === 'runCurrentModule') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runCurrentModule',
                    failFast: filterState.failFast,
                });
                return;
            }
            if (button.dataset.action === 'runCurrentTest') {
                setRunning(true);
                vscode.postMessage({
                    type: 'runCurrentTest',
                    failFast: filterState.failFast,
                });
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
            if (button.dataset.action === 'rerunFailed') {
                setRunning(true);
                vscode.postMessage({ type: 'rerunFailed' });
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

function renderVbaTestsErrorHtml(workbookName: string, error: string): string {
    const nonce = randomNonce();
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Tests Error</title>
    <style nonce="${nonce}">
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

function renderTestSelection(model: VbaTestsPanelModel): string {
    const discovery = model.discovery;
    const testSummary = discovery.error
        ? escapeHtml(discovery.error)
        : `${discovery.tests.length} discovered`;
    return /* html */`<div class="filterPanel">
        <div class="testListHeader">
            <div>
                <div class="filterTitle">Tests</div>
                <div class="filterSummary" id="selectedTestSummary">${testSummary}</div>
            </div>
            <span class="filterTools">
                <button class="secondary compact" type="button" data-test-action="selectAll">Select All</button>
                <button class="secondary compact" type="button" data-test-action="clear">Clear</button>
            </span>
        </div>
        ${discovery.tests.length > 0 ? /* html */`
            <div class="testList">
                ${discovery.tests.map(renderTestChoice).join('')}
            </div>
        ` : /* html */`
            <div class="emptyState">No tests discovered</div>
        `}
    </div>`;
}

function renderTestChoice(test: VbaTestListItemModel): string {
    const tags = test.tags.length > 0
        ? test.tags.map((tag) => `<span class="testTag">${escapeHtml(tag)}</span>`).join('')
        : '<span class="subtle">untagged</span>';
    return /* html */`<label class="testChoice" title="${escapeAttr(`${test.qualifiedName} at line ${test.line}`)}">
        <input type="checkbox" data-test-id="${escapeAttr(test.id)}" checked>
        <span>
            <span class="testName">${escapeHtml(test.qualifiedName)}</span>
            <span class="testMeta">${escapeHtml(test.moduleName)}:${test.line}</span>
        </span>
        <span class="testTags">${tags}</span>
    </label>`;
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

function runSelectedRequestFromMessage(message: VbaTestsWebviewMessage): VbaTestsRunSelectedRequest {
    return {
        testIds: stringListFromUnknown(message.testIds),
        failFast: Boolean(message.failFast),
    };
}

function runModeRequestFromMessage(message: VbaTestsWebviewMessage): VbaTestsRunModeRequest {
    return {
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

function vbaTestsPanelKey(filePath: string): string {
    const normalized = path.normalize(filePath);
    return process.platform === 'win32' ? normalized.toLowerCase() : normalized;
}
