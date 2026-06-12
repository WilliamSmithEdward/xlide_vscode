import * as path from 'path';
import * as vscode from 'vscode';
import { measurePerformance } from './performanceTrace';
import { escapeAttr, escapeHtml, randomNonce, scriptJson } from './webview/html';
import {
    renderWebviewErrorPageHtml,
    webviewHeadHtml,
    WEBVIEW_TOAST_CSS,
    WEBVIEW_TOAST_HTML,
    WEBVIEW_TOAST_SCRIPT,
} from './webview/page';
import { bridgeWebviewMessages, createWebviewPanelRegistry } from './webview/panelRegistry';
import { DebouncedRefresher } from './webview/refresh';
import { WEBVIEW_BODY_CSS, WEBVIEW_PRIMARY_BUTTON_CSS, xlideAccentPaletteCss } from './webview/styles';
import { renderWebviewTemplate } from './webview/templates';
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

const openVbaTestsPanels = createWebviewPanelRegistry<OpenVbaTestsPanelEntry>();

export function openVbaTestsPanel(
    context: vscode.ExtensionContext,
    filePath: string,
    options: VbaTestsPanelOptions,
): vscode.WebviewPanel {
    const existing = openVbaTestsPanels.get(filePath);
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
    const panel = vscode.window.createWebviewPanel(
        'xlideVbaTests',
        `XLIDE Tests: ${path.basename(filePath)}`,
        vscode.ViewColumn.Active,
        {
            // No retainContextWhenHidden: every refresh reassigns webview.html,
            // and the script restores its filter state from vscode.getState.
            enableScripts: true,
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
    openVbaTestsPanels.set(filePath, entry);

    const refresher = new DebouncedRefresher({
        refresh: renderPanel,
        onError: (err) => {
            const error = errorMessage(err);
            void panel.webview.postMessage({ type: 'error', error });
        },
        defaultDelayMs: 250,
    });

    const runAndRefresh = async (
        operation: (() => Promise<void>) | undefined,
        missingMessage: string,
    ): Promise<void> => {
        if (!operation) {
            await panel.webview.postMessage({ type: 'error', error: missingMessage });
            return;
        }
        await refresher.runExclusive(async () => {
            await operation();
            await renderPanel();
        });
    };

    void renderPanel().catch((err) => {
        const error = errorMessage(err);
        panel.webview.html = renderVbaTestsErrorHtml(path.basename(filePath), error);
    });

    const messageSub = bridgeWebviewMessages(
        panel.webview,
        async (message: VbaTestsWebviewMessage) => {
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
        },
        () => renderPanel().catch(() => { /* keep existing error visible */ }),
    );

    const treeSub = entry.options.onDidChangeWorkbookTree?.(() => refresher.schedule());
    const panelDisposables = [
        messageSub,
        refresher,
        ...(treeSub ? [treeSub] : []),
    ];
    panel.onDidDispose(() => {
        disposed = true;
        openVbaTestsPanels.delete(filePath);
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

    return renderWebviewTemplate('assets/webview/vbaTests.html', {
        head: webviewHeadHtml(nonce, 'XLIDE Tests'),
        nonce,
        css: renderWebviewTemplate('assets/webview/vbaTests.css', {
            accentPalette: xlideAccentPaletteCss(),
            bodyCss: WEBVIEW_BODY_CSS,
            primaryButtonCss: WEBVIEW_PRIMARY_BUTTON_CSS,
            toastCss: WEBVIEW_TOAST_CSS,
        }),
        workbookName: escapeHtml(model.workbookName),
        supportState: escapeAttr(model.support.state),
        installTitle: escapeAttr(installTitle),
        installDisabled,
        installLabel: escapeHtml(model.support.actionLabel),
        supportTitle: escapeHtml(model.support.title),
        supportDescription: escapeHtml(model.support.description),
        runtimeState: escapeAttr(model.runtime.state),
        runtimeTitle: escapeHtml(model.runtime.title),
        runtimeDescription: escapeHtml(model.runtime.description),
        runDisabled,
        selectedRunTitle: escapeAttr(selectedRunTitle),
        selectedRunDisabled,
        currentScopeTitle: escapeAttr(currentScopeTitle),
        filterRunTitle: escapeAttr(filterRunTitle),
        filterRunDisabled,
        rerunFailedTitle: escapeAttr(rerunFailedTitle),
        rerunFailedDisabled,
        rerunFailedCount: hasLastFailed ? ` (${failedCount})` : '',
        runHelp: runHelp ? `<p class="helpText">${escapeHtml(runHelp)}</p>` : '',
        testSelection: renderTestSelection(model),
        tagFilters: renderTagFilters(model),
        toastHtml: WEBVIEW_TOAST_HTML,
        js: renderWebviewTemplate('assets/webview/vbaTests.js', {
            toastScript: WEBVIEW_TOAST_SCRIPT,
            workbookPathJson,
            tagNamesJson,
            testIdsJson,
            canRunJson,
            hasTagsJson,
            hasTestsJson,
            hasLastFailedJson,
        }),
    });
}


function renderVbaTestsErrorHtml(workbookName: string, error: string): string {
    return renderWebviewErrorPageHtml({
        title: 'XLIDE Tests Error',
        heading: 'XLIDE Tests Could Not Load',
        subtitle: workbookName,
        error,
    });
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

