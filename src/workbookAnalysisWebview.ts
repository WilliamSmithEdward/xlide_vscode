import * as vscode from 'vscode';
import * as path from 'path';
import type { WorkbookAnalysisProblem, WorkbookAnalysisResult } from './vbaWorkbookAnalysis';
import {
    buildWorkbookAnalysisPlainText,
    buildWorkbookAnalysisResultsModel,
    type WorkbookAnalysisResultRow,
    type WorkbookAnalysisResultsModel,
} from './workbookAnalysisResultsModel';
import {
    ANALYSIS_SEVERITIES,
    allowedAnalysisRuleSeverityOverrides,
    isAnalysisRuleTracked,
    type AnalysisRuleSeverityOverrides,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import { setGlobalAnalysisRuleTracked } from './analysisOptions';
import {
    clearWorkbookAnalysisRuleSeverityOverride,
    effectiveWorkbookAnalysisSettings,
    resetWorkbookAnalysisRuleTracking,
    resetWorkbookAnalysisSettings,
    resetWorkbookAnalysisRuleSeverities,
    resetWorkbookAnalysisVisibleSeverities,
    setWorkbookAnalysisRuleTracked,
    setWorkbookAnalysisRuleSeverityOverride,
    setWorkbookAnalysisVisibleSeverities,
    type EffectiveWorkbookAnalysisSettings,
} from './workbookAnalysisSettings';
import { settingsPathForWorkbook } from './workbookSettings';
import { decodeModuleUri, sameWorkbookPath, XLIDE_SCHEME } from './xlideFileSystem';

export type WorkbookAnalysisSuppressScope = 'block' | 'member' | 'module';

export interface WorkbookAnalysisResultsOptions {
    onOpenProblem?: (problem: WorkbookAnalysisProblem, analysisPanelColumn?: vscode.ViewColumn) => Promise<void>;
    onQuickFixProblem?: (
        problem: WorkbookAnalysisProblem,
        analysisPanelColumn?: vscode.ViewColumn,
        fixIndex?: number,
    ) => Promise<boolean>;
    onSuppressProblem?: (
        problem: WorkbookAnalysisProblem,
        scope: WorkbookAnalysisSuppressScope,
        analysisPanelColumn?: vscode.ViewColumn,
    ) => Promise<void>;
    onAskCopilot?: (problem: WorkbookAnalysisProblem, analysisPanelColumn?: vscode.ViewColumn) => Promise<void>;
    onRefreshResult?: () => Promise<WorkbookAnalysisResult>;
    onDidChangeWorkbookTree?: vscode.Event<unknown>;
}

interface WorkbookAnalysisMessage {
    type?: string;
    index?: number;
    text?: string;
    extension?: string;
    scope?: WorkbookAnalysisSuppressScope;
    severity?: string;
    severities?: string[];
    fixIndex?: number;
    suppressed?: boolean;
    tracked?: boolean;
    trackingScope?: 'workbook' | 'global';
    code?: string;
    moduleName?: string;
    moduleType?: string;
    line?: number;
    column?: number;
    endColumn?: number;
    message?: string;
}

interface OpenWorkbookAnalysisResultsPanelEntry {
    panel: vscode.WebviewPanel;
    options: WorkbookAnalysisResultsOptions;
    renderPanel: () => Promise<void>;
    setResult: (result: WorkbookAnalysisResult) => void;
}

const openWorkbookAnalysisResultsPanels = new Map<string, OpenWorkbookAnalysisResultsPanelEntry>();

export function openWorkbookAnalysisResults(
    context: vscode.ExtensionContext,
    result: WorkbookAnalysisResult,
    options: WorkbookAnalysisResultsOptions = {},
): vscode.WebviewPanel {
    const panelKey = workbookAnalysisResultsPanelKey(result.filePath);
    const existing = openWorkbookAnalysisResultsPanels.get(panelKey);
    if (existing) {
        existing.options = options;
        existing.setResult(result);
        existing.panel.reveal(vscode.ViewColumn.Active);
        void existing.renderPanel().catch((err) => {
            const error = err instanceof Error ? err.message : String(err);
            existing.panel.webview.html = renderWorkbookAnalysisErrorHtml(
                existing.panel.webview,
                path.basename(result.filePath),
                error,
            );
        });
        return existing.panel;
    }

    let currentResult = result;
    let currentModel = buildWorkbookAnalysisResultsModel(currentResult);
    let disposed = false;
    let refreshTimer: ReturnType<typeof setTimeout> | undefined;
    let refreshVersion = 0;
    let contextMenuOpen = false;
    let pendingRefreshAfterContextMenu = false;
    let severitySettingsSaveVersion = 0;
    let severitySettingsSaveQueue: Promise<void> = Promise.resolve();
    let ignoreOwnSeveritySettingsRefresh = false;
    let ignoreOwnSeveritySettingsRefreshTimer: ReturnType<typeof setTimeout> | undefined;
    const panel = vscode.window.createWebviewPanel(
        'xlideWorkbookAnalysisResults',
        `XLIDE Analysis: ${currentModel.workbookName}`,
        vscode.ViewColumn.Active,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
        },
    );
    const entry: OpenWorkbookAnalysisResultsPanelEntry = {
        panel,
        options,
        renderPanel: async () => { /* assigned below */ },
        setResult: (nextResult) => {
            currentResult = nextResult;
            currentModel = buildWorkbookAnalysisResultsModel(currentResult);
        },
    };

    const renderPanel = async (): Promise<void> => {
        currentModel = buildWorkbookAnalysisResultsModel(currentResult);
        const analysisSettings = await effectiveWorkbookAnalysisSettings(currentResult.filePath);
        if (disposed) { return; }
        panel.title = `XLIDE Analysis: ${currentModel.workbookName}`;
        panel.webview.html = renderWorkbookAnalysisResultsHtml(
            panel.webview,
            currentModel,
            analysisSettings,
        );
    };
    entry.renderPanel = renderPanel;
    openWorkbookAnalysisResultsPanels.set(panelKey, entry);

    const refreshPanel = async (requestVersion: number): Promise<void> => {
        if (disposed) { return; }
        if (requestVersion !== refreshVersion) { return; }
        if (contextMenuOpen) {
            pendingRefreshAfterContextMenu = true;
            return;
        }
        if (entry.options.onRefreshResult) {
            const nextResult = await entry.options.onRefreshResult();
            if (requestVersion !== refreshVersion || disposed) {
                return;
            }
            currentResult = nextResult;
        }
        if (contextMenuOpen) {
            pendingRefreshAfterContextMenu = true;
            return;
        }
        if (!disposed) {
            await renderPanel();
        }
    };

    const scheduleRefresh = (): void => {
        const requestVersion = ++refreshVersion;
        if (disposed) { return; }
        if (contextMenuOpen) {
            pendingRefreshAfterContextMenu = true;
            return;
        }
        if (refreshTimer) { clearTimeout(refreshTimer); }
        refreshTimer = setTimeout(() => {
            refreshTimer = undefined;
            void refreshPanel(requestVersion).catch((err) => {
                const error = err instanceof Error ? err.message : String(err);
                void panel.webview.postMessage({ type: 'error', error });
            });
        }, 350);
    };

    const refreshAfterAnalysisMutation = async (): Promise<void> => {
        const requestVersion = ++refreshVersion;
        if (disposed) { return; }
        if (refreshTimer) {
            clearTimeout(refreshTimer);
            refreshTimer = undefined;
        }
        contextMenuOpen = false;
        pendingRefreshAfterContextMenu = false;
        await refreshPanel(requestVersion);
    };

    const saveVisibleSeveritySettings = async (severities: string[] | undefined): Promise<void> => {
        const requestVersion = ++severitySettingsSaveVersion;
        severitySettingsSaveQueue = severitySettingsSaveQueue
            .catch(() => undefined)
            .then(async () => {
                if (requestVersion !== severitySettingsSaveVersion) {
                    return;
                }
                ignoreOwnSidecarRefreshBriefly();
                await setWorkbookAnalysisVisibleSeverities(currentResult.filePath, severities);
                if (!disposed && requestVersion === severitySettingsSaveVersion) {
                    await panel.webview.postMessage({ type: 'severitySettingsSaved' });
                }
            });
        await severitySettingsSaveQueue;
    };

    const ignoreOwnSidecarRefreshBriefly = (): void => {
        ignoreOwnSeveritySettingsRefresh = true;
        if (ignoreOwnSeveritySettingsRefreshTimer) {
            clearTimeout(ignoreOwnSeveritySettingsRefreshTimer);
        }
        ignoreOwnSeveritySettingsRefreshTimer = setTimeout(() => {
            ignoreOwnSeveritySettingsRefresh = false;
            ignoreOwnSeveritySettingsRefreshTimer = undefined;
        }, 1000);
    };

    const scheduleSidecarRefresh = (): void => {
        if (ignoreOwnSeveritySettingsRefresh) {
            return;
        }
        scheduleRefresh();
    };

    const problemAt = (index: unknown, suppressed?: boolean): WorkbookAnalysisProblem | undefined => {
        if (typeof index !== 'number') {
            return undefined;
        }
        return suppressed
            ? currentResult.suppressedProblems[index]
            : currentResult.problems[index];
    };

    const problemForOpenMessage = (message: WorkbookAnalysisMessage): WorkbookAnalysisProblem | undefined => {
        const indexedProblem = problemAt(message.index, message.suppressed);
        const rowProblem = problemFromOpenMessage(message);
        if (indexedProblem && (!rowProblem || sameProblemLocation(indexedProblem, rowProblem))) {
            return indexedProblem;
        }
        return rowProblem ?? indexedProblem;
    };

    void renderPanel().catch((err) => {
        const error = err instanceof Error ? err.message : String(err);
        panel.webview.html = renderWorkbookAnalysisErrorHtml(panel.webview, currentModel.workbookName, error);
    });
    const messageSub = panel.webview.onDidReceiveMessage(async (message: WorkbookAnalysisMessage) => {
        try {
            if (message.type === 'contextMenuOpened') {
                contextMenuOpen = true;
                if (refreshTimer) {
                    clearTimeout(refreshTimer);
                    refreshTimer = undefined;
                    pendingRefreshAfterContextMenu = true;
                }
                return;
            }
            if (message.type === 'contextMenuClosed') {
                contextMenuOpen = false;
                if (pendingRefreshAfterContextMenu) {
                    pendingRefreshAfterContextMenu = false;
                    scheduleRefresh();
                }
                return;
            }
            if (message.type === 'openProblem') {
                const problem = problemForOpenMessage(message);
                if (problem && entry.options.onOpenProblem) {
                    await entry.options.onOpenProblem(problem, panel.viewColumn);
                }
                return;
            }
            if (message.type === 'suppressProblem') {
                const problem = problemAt(message.index, message.suppressed);
                const scope = message.scope;
                if (problem && message.suppressed === true) {
                    await panel.webview.postMessage({ type: 'error', error: 'This analysis finding is already suppressed.' });
                    return;
                }
                if (problem && scope && !problem.suppressionScopes.includes(scope)) {
                    await panel.webview.postMessage({ type: 'error', error: `Ignore ${scope} is not valid for this analysis finding.` });
                    return;
                }
                if (problem && scope && entry.options.onSuppressProblem) {
                    await entry.options.onSuppressProblem(problem, scope, panel.viewColumn);
                    await panel.webview.postMessage({ type: 'suppressed', scope });
                    await refreshAfterAnalysisMutation();
                }
                return;
            }
            if (message.type === 'askCopilot') {
                const problem = problemAt(message.index, message.suppressed);
                if (problem && entry.options.onAskCopilot) {
                    await entry.options.onAskCopilot(problem, panel.viewColumn);
                }
                return;
            }
            if (message.type === 'quickFixProblem') {
                const problem = problemAt(message.index, message.suppressed);
                const applied = problem && entry.options.onQuickFixProblem
                    ? await entry.options.onQuickFixProblem(problem, panel.viewColumn, message.fixIndex)
                    : false;
                await panel.webview.postMessage({
                    type: applied ? 'quickFixed' : 'quickFixUnavailable',
                });
                if (applied) {
                    await refreshAfterAnalysisMutation();
                }
                return;
            }
            if (message.type === 'setRuleTracking') {
                const problem = problemAt(message.index, message.suppressed);
                const code = typeof message.code === 'string' ? message.code : problem?.code;
                if (code) {
                    const scope = message.trackingScope === 'global' ? 'global' : 'workbook';
                    const update = scope === 'global'
                        ? await setGlobalAnalysisRuleTracked(code, message.tracked === true)
                        : await setWorkbookAnalysisRuleTracked(currentResult.filePath, code, message.tracked === true);
                    await refreshAfterAnalysisMutation();
                    await panel.webview.postMessage({
                        type: 'ruleTrackingChanged',
                        scope,
                        code: update.code ?? code,
                        tracked: update.tracked,
                        untrackedRules: update.untrackedRules,
                    });
                }
                return;
            }
            if (message.type === 'setRuleSeverityOverride') {
                const code = typeof message.code === 'string' ? message.code : undefined;
                if (code && message.severity === 'default') {
                    await clearWorkbookAnalysisRuleSeverityOverride(currentResult.filePath, code);
                    await refreshAfterAnalysisMutation();
                } else if (code) {
                    await setWorkbookAnalysisRuleSeverityOverride(currentResult.filePath, code, message.severity);
                    await refreshAfterAnalysisMutation();
                }
                return;
            }
            if (message.type === 'updateSeveritySettings') {
                await saveVisibleSeveritySettings(message.severities);
                return;
            }
            if (message.type === 'resetAnalysisSeverities') {
                await resetWorkbookAnalysisVisibleSeverities(currentResult.filePath);
                await refreshAfterAnalysisMutation();
                return;
            }
            if (message.type === 'resetAnalysisRuleTracking') {
                await resetWorkbookAnalysisRuleTracking(currentResult.filePath);
                await refreshAfterAnalysisMutation();
                return;
            }
            if (message.type === 'resetAnalysisRuleSeverities') {
                await resetWorkbookAnalysisRuleSeverities(currentResult.filePath);
                await refreshAfterAnalysisMutation();
                return;
            }
            if (message.type === 'resetAnalysisSettings') {
                await resetWorkbookAnalysisSettings(currentResult.filePath);
                await refreshAfterAnalysisMutation();
                return;
            }
            if (message.type === 'copyText') {
                await vscode.env.clipboard.writeText(String(message.text ?? ''));
                await panel.webview.postMessage({ type: 'copied' });
                return;
            }
            if (message.type === 'exportText') {
                const extension = message.extension === 'json' ? 'json' : 'txt';
                const target = await vscode.window.showSaveDialog({
                    title: extension === 'json' ? 'Export XLIDE Analysis JSON' : 'Export XLIDE Analysis Report',
                    defaultUri: vscode.Uri.file(path.join(
                        path.dirname(currentModel.filePath),
                        `${sanitizeFileName(currentModel.workbookName)}.xlide-analysis.${extension}`,
                    )),
                    filters: extension === 'json' ? { JSON: ['json'] } : { Text: ['txt'] },
                });
                if (target) {
                    await vscode.workspace.fs.writeFile(
                        target,
                        Buffer.from(String(message.text ?? ''), 'utf8'),
                    );
                    await panel.webview.postMessage({ type: 'exported' });
                }
            }
        } catch (err) {
            const error = err instanceof Error ? err.message : String(err);
            await panel.webview.postMessage({ type: 'error', error });
        }
    });

    const configSub = vscode.workspace.onDidChangeConfiguration((e) => {
        if (e.affectsConfiguration('xlide.analysis') || e.affectsConfiguration('xlide.diagnostics')) {
            scheduleRefresh();
        }
    });
    const textChangeSub = vscode.workspace.onDidChangeTextDocument((e) => {
        if (isWorkbookDocument(e.document, currentResult.filePath)) {
            scheduleRefresh();
        }
    });
    const saveSub = vscode.workspace.onDidSaveTextDocument((document) => {
        if (isWorkbookDocument(document, currentResult.filePath)) {
            scheduleRefresh();
        }
    });
    const treeSub = entry.options.onDidChangeWorkbookTree?.(() => scheduleRefresh());
    const sidecarPath = settingsPathForWorkbook(currentResult.filePath);
    const sidecarWatcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(
        path.dirname(sidecarPath),
        path.basename(sidecarPath),
    ));
    const panelDisposables = [
        messageSub,
        configSub,
        textChangeSub,
        saveSub,
        sidecarWatcher,
        sidecarWatcher.onDidCreate(scheduleSidecarRefresh),
        sidecarWatcher.onDidChange(scheduleSidecarRefresh),
        sidecarWatcher.onDidDelete(scheduleSidecarRefresh),
        ...(treeSub ? [treeSub] : []),
    ];
    panel.onDidDispose(() => {
        disposed = true;
        openWorkbookAnalysisResultsPanels.delete(panelKey);
        if (refreshTimer) {
            clearTimeout(refreshTimer);
            refreshTimer = undefined;
        }
        if (ignoreOwnSeveritySettingsRefreshTimer) {
            clearTimeout(ignoreOwnSeveritySettingsRefreshTimer);
            ignoreOwnSeveritySettingsRefreshTimer = undefined;
        }
        for (const sub of panelDisposables) {
            sub.dispose();
        }
    });

    context.subscriptions.push(panel);
    return panel;
}

function isWorkbookDocument(document: vscode.TextDocument, filePath: string): boolean {
    if (document.uri.scheme !== XLIDE_SCHEME) {
        return false;
    }
    try {
        return sameWorkbookPath(decodeModuleUri(document.uri).xlsmPath, filePath);
    } catch {
        return false;
    }
}

export function renderWorkbookAnalysisResultsHtml(
    webview: vscode.Webview,
    model: WorkbookAnalysisResultsModel,
    analysisSettings: EffectiveWorkbookAnalysisSettings,
): string {
    const nonce = randomNonce();
    const visibleSeverities = analysisSettings.visibleSeverities;
    const untrackedRules = analysisSettings.untrackedRules;
    const ruleSeverityOverrides = analysisSettings.ruleSeverityOverrides;
    const modelJson = JSON.stringify({
        ...model,
        plainText: buildWorkbookAnalysisPlainText(model),
        visibleSeverities,
        untrackedRules,
        ruleSeverityOverrides,
        visibleSeveritiesSource: analysisSettings.visibleSeveritiesSource,
        untrackedRulesSource: analysisSettings.untrackedRulesSource,
        ruleSeverityOverridesSource: analysisSettings.ruleSeverityOverridesSource,
        analysisSettingsKey: workbookAnalysisSettingsKey(analysisSettings),
    }).replace(/</g, '\\u003c');
    const moduleOrder = new Map(model.groups.map((group, index) => [group.moduleName.toLowerCase(), index]));
    const allRows = [...model.rows, ...model.suppressedRows];
    const rowsHtml = allRows.length === 0
        ? '<div class="empty">No analysis findings.</div>'
        : allRows.map((row) => {
            const tracked = isAnalysisRuleTracked(row.code, untrackedRules);
            const trackingSource = analysisRuleTrackingSourceForRow(tracked, analysisSettings.untrackedRulesSource);
            const status = !tracked
                ? analysisRuleTrackingStatusLabel(trackingSource)
                : (row.suppressed ? 'Suppressed' : 'Tracked');
            const statusKey = !tracked ? 'untracked' : (row.suppressed ? 'suppressed' : 'tracked');
            return `
            <button
                class="problemRow severity-${escapeAttr(row.severity)}"
                type="button"
                data-open-index="${row.index}"
                data-suppressed="${row.suppressed ? 'yes' : 'no'}"
                data-status="${statusKey}"
                data-module="${escapeAttr(row.moduleName)}"
                data-module-type="${escapeAttr(row.moduleType)}"
                data-module-order="${moduleOrder.get(row.moduleName.toLowerCase()) ?? 9999}"
                data-severity="${escapeAttr(row.severity)}"
                data-compile="${row.vbeCompileEquivalent ? 'yes' : 'no'}"
                data-line="${row.line}"
                data-column="${row.column}"
                data-end-column="${row.endColumn}"
                data-rule="${escapeAttr(row.code || row.ruleTitle)}"
                data-rule-code="${escapeAttr(row.code)}"
                data-message="${escapeAttr(row.message)}"
                data-evidence="${escapeAttr(row.diagnosticKind)}"
                data-quick-fixes="${escapeAttr(JSON.stringify(row.quickFixTitles))}"
                data-suppression-scopes="${escapeAttr(JSON.stringify(row.suppressionScopes))}"
                data-tracked="${tracked ? 'yes' : 'no'}"
                data-tracking-source="${escapeAttr(trackingSource)}"
            >
                <span class="cell severity">${escapeHtml(row.severity)}</span>
                <span class="cell status">${escapeHtml(status)}</span>
                <span class="cell location">${escapeHtml(row.location)}</span>
                <span class="cell code">${escapeHtml(row.code || row.ruleTitle)}</span>
                <span class="cell kind">${escapeHtml(row.diagnosticKind)}</span>
                <span class="cell message">${escapeHtml(row.message)}</span>
            </button>
        `;
        }).join('');
    const moduleButtons = model.groups.length === 0
        ? '<button class="moduleFilter active" type="button" data-module-filter="all">All modules <span>0</span></button>'
        : [
            `<button class="moduleFilter active" type="button" data-module-filter="all">All modules <span>${model.totalProblems}</span></button>`,
            ...model.groups.map((group) => `
                <button class="moduleFilter" type="button" data-module-filter="${escapeAttr(group.moduleName)}">
                    <span class="moduleIdentity">
                        <span class="moduleIcon" title="${escapeAttr(group.moduleTypeLabel)}">${escapeHtml(group.moduleIcon)}</span>
                        <span class="moduleName">${escapeHtml(group.moduleName)}</span>
                    </span>
                    <span>${group.total}</span>
                </button>
            `),
        ].join('');
    const filterButtons = ANALYSIS_SEVERITIES.map((severity) => {
        const active = visibleSeverities.includes(severity);
        return `<button class="filterButton${active ? ' active' : ''}" type="button" data-severity-toggle="${severity}" aria-pressed="${active ? 'true' : 'false'}">${severityFilterLabel(severity)}</button>`;
    }).join('');
    const ruleSettingsHtml = analysisRuleSettingsHtml(allRows, untrackedRules);
    const ruleSeveritySettingsHtml = analysisRuleSeveritySettingsHtml(allRows, ruleSeverityOverrides);
    const severitySourceIsWorkbook = analysisSettings.visibleSeveritiesSource === 'workbook';
    const rulesSourceIsWorkbook = analysisSettings.untrackedRulesSource === 'workbook';
    const ruleSeveritiesSourceIsWorkbook = analysisSettings.ruleSeverityOverridesSource === 'workbook';
    const anyAnalysisOverride = severitySourceIsWorkbook || rulesSourceIsWorkbook || ruleSeveritiesSourceIsWorkbook;
    const informationCount = model.rows.filter((row) => row.severity === 'information').length;
    const untrackedCount = allRows.filter((row) => !isAnalysisRuleTracked(row.code, untrackedRules)).length;
    const suppressedCount = allRows.filter((row) => row.suppressed && isAnalysisRuleTracked(row.code, untrackedRules)).length;
    const summaryStatsHtml = [
        statHtml(String(model.errorCount), 'Errors'),
        statHtml(String(model.warningCount), 'Warnings'),
        statHtml(String(informationCount), 'Information'),
        statHtml(String(suppressedCount), 'Suppressed'),
        statHtml(String(untrackedCount), 'Untracked'),
        statHtml(String(model.moduleCount), 'Modules'),
    ].join('');

    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Analysis Results</title>
    <style>
        :root {
            color-scheme: light dark;
            --xlide-accent-blue: #2d5f94;
            --xlide-accent-blue-hover: #376fa8;
            --xlide-accent-background: color-mix(in srgb, var(--xlide-accent-blue) 82%, var(--vscode-editor-background));
            --xlide-accent-hover-background: color-mix(in srgb, var(--xlide-accent-blue-hover) 84%, var(--vscode-editor-background));
            --xlide-accent-border: color-mix(in srgb, var(--xlide-accent-blue) 78%, var(--vscode-panel-border));
        }
        * {
            box-sizing: border-box;
        }
        body {
            margin: 0;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
            font-size: var(--vscode-font-size);
        }
        button {
            font: inherit;
        }
        .shell {
            min-height: 100vh;
            display: grid;
            grid-template-rows: auto auto 1fr;
        }
        header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 14px 18px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        h1 {
            margin: 0;
            font-size: 16px;
            font-weight: 600;
        }
        .subtle {
            color: var(--vscode-descriptionForeground);
        }
        .headerActions {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            justify-content: flex-end;
        }
        .actionButton,
        .filterButton,
        .moduleFilter,
        .secondaryButton {
            border: 1px solid var(--vscode-button-border, var(--vscode-panel-border));
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
            border-radius: 4px;
            padding: 5px 10px;
            cursor: pointer;
        }
        .actionButton:hover,
        .filterButton:hover,
        .moduleFilter:hover,
        .secondaryButton:hover {
            background: var(--vscode-button-secondaryHoverBackground);
        }
        .secondaryButton {
            color: var(--vscode-foreground);
            background: transparent;
        }
        .filterButton.active,
        .moduleFilter.active {
            color: var(--vscode-button-foreground);
            border-color: var(--xlide-accent-border);
            background: var(--xlide-accent-background);
        }
        .filterButton.active:hover,
        .moduleFilter.active:hover {
            background: var(--xlide-accent-hover-background);
        }
        .filterButton.transient {
            margin-left: 6px;
        }
        .stats {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 8px 28px;
            padding: 9px 18px;
            background: var(--vscode-editor-background);
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        .stat {
            min-width: 0;
            display: grid;
            gap: 1px;
            min-inline-size: 96px;
            padding: 0;
        }
        .stat strong {
            display: block;
            font-size: 16px;
            font-weight: 600;
            line-height: 1.2;
        }
        .stat span {
            display: block;
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
            letter-spacing: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .body {
            min-height: 0;
            display: grid;
            grid-template-columns: minmax(180px, 260px) 1fr;
        }
        aside {
            min-width: 0;
            padding: 12px;
            border-right: 1px solid var(--vscode-panel-border);
            overflow: auto;
        }
        aside h2,
        main h2 {
            margin: 0 0 8px;
            font-size: 13px;
            font-weight: 600;
            letter-spacing: 0;
            color: var(--vscode-descriptionForeground);
        }
        .moduleList {
            display: grid;
            gap: 6px;
        }
        .moduleFilter {
            width: 100%;
            display: grid;
            grid-template-columns: 1fr auto;
            align-items: center;
            gap: 8px;
            text-align: left;
        }
        .moduleIdentity {
            min-width: 0;
            display: flex;
            align-items: center;
            gap: 7px;
        }
        .moduleIcon {
            width: 18px;
            height: 18px;
            display: inline-grid;
            place-items: center;
            flex: 0 0 auto;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 3px;
            color: var(--vscode-descriptionForeground);
            font-size: 11px;
            line-height: 1;
        }
        .moduleName {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        main {
            min-width: 0;
            display: grid;
            grid-template-rows: auto 1fr;
            overflow: hidden;
        }
        .toolbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            padding: 12px 14px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        .filters {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        .visibleCount {
            color: var(--vscode-descriptionForeground);
            white-space: nowrap;
        }
        .table {
            min-width: 0;
            overflow: auto;
        }
        .tableHeader,
        .problemRow {
            display: grid;
            grid-template-columns: 124px minmax(184px, 210px) 150px minmax(132px, 190px) 132px minmax(260px, 1fr);
            align-items: stretch;
            min-width: 1120px;
        }
        .tableHeader {
            position: sticky;
            top: 0;
            z-index: 1;
            background: var(--vscode-editor-background);
            color: var(--vscode-descriptionForeground);
            border-bottom: 1px solid var(--vscode-panel-border);
            font-weight: 600;
        }
        .sortHeader {
            border: 0;
            display: grid;
            grid-template-columns: minmax(0, auto) 12px;
            align-items: center;
            gap: 5px;
            color: inherit;
            background: transparent;
            text-align: left;
            cursor: pointer;
            font-weight: inherit;
        }
        .sortHeader:hover,
        .sortHeader:focus {
            outline: none;
            background: var(--vscode-list-hoverBackground);
        }
        .sortHeader[aria-sort="ascending"],
        .sortHeader[aria-sort="descending"] {
            color: var(--vscode-foreground);
        }
        .sortHeaderText {
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .sortIndicator {
            width: 10px;
            height: 10px;
            color: currentColor;
        }
        .sortIndicator::before {
            content: "";
            display: block;
            width: 0;
            height: 0;
            margin: 2px auto 0;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            opacity: 0;
        }
        .sortHeader[aria-sort="ascending"] .sortIndicator::before {
            border-bottom: 6px solid currentColor;
            opacity: 1;
        }
        .sortHeader[aria-sort="descending"] .sortIndicator::before {
            border-top: 6px solid currentColor;
            opacity: 1;
        }
        .problemRow {
            width: 100%;
            border: 0;
            border-bottom: 1px solid var(--vscode-panel-border);
            color: var(--vscode-foreground);
            background: transparent;
            text-align: left;
            cursor: pointer;
        }
        .problemRow:hover,
        .problemRow:focus {
            outline: none;
            background: var(--vscode-list-hoverBackground);
        }
        .problemRow.hiddenVisible {
            opacity: 0.52;
            background-image: repeating-linear-gradient(
                -45deg,
                transparent,
                transparent 9px,
                color-mix(in srgb, var(--vscode-descriptionForeground) 18%, transparent) 9px,
                color-mix(in srgb, var(--vscode-descriptionForeground) 18%, transparent) 10px
            );
        }
        .problemRow.suppressedVisible {
            opacity: 0.62;
        }
        .problemRow[hidden] {
            display: none !important;
        }
        .cell {
            min-width: 0;
            padding: 8px 10px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .message {
            white-space: normal;
            line-height: 1.35;
        }
        .severity {
            font-weight: 600;
            text-transform: capitalize;
            text-overflow: clip;
        }
        .severity-error .severity {
            color: var(--vscode-errorForeground);
        }
        .severity-warning .severity {
            color: var(--vscode-editorWarning-foreground);
        }
        .code,
        .kind,
        .location,
        .status {
            color: var(--vscode-descriptionForeground);
        }
        .empty {
            padding: 32px 14px;
            color: var(--vscode-descriptionForeground);
        }
        .toast {
            position: fixed;
            right: 14px;
            bottom: 14px;
            max-width: min(420px, calc(100vw - 28px));
            padding: 8px 10px;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 4px;
            color: var(--vscode-notifications-foreground);
            background: var(--vscode-notifications-background);
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.24);
        }
        .contextMenu {
            position: fixed;
            z-index: 10;
            min-width: 190px;
            padding: 4px;
            border: 1px solid var(--vscode-menu-border, var(--vscode-panel-border));
            border-radius: 4px;
            color: var(--vscode-menu-foreground, var(--vscode-foreground));
            background: var(--vscode-menu-background, var(--vscode-editor-background));
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28);
        }
        .contextMenu button {
            width: 100%;
            display: block;
            border: 0;
            border-radius: 3px;
            padding: 6px 9px;
            color: inherit;
            background: transparent;
            text-align: left;
            cursor: pointer;
        }
        .contextMenu button:hover,
        .contextMenu button:focus {
            outline: none;
            background: var(--vscode-menu-selectionBackground, var(--vscode-list-hoverBackground));
            color: var(--vscode-menu-selectionForeground, var(--vscode-foreground));
        }
        .contextMenu button:disabled {
            color: var(--vscode-disabledForeground);
            background: transparent;
            cursor: default;
        }
        .contextMenu button:disabled:hover,
        .contextMenu button:disabled:focus {
            color: var(--vscode-disabledForeground);
            background: transparent;
        }
        .contextMenu button[hidden],
        .contextDivider[hidden] {
            display: none;
        }
        .contextDivider {
            height: 1px;
            margin: 4px;
            background: var(--vscode-menu-separatorBackground, var(--vscode-panel-border));
        }
        .contextItemWithSubmenu {
            position: relative;
        }
        .contextItemWithSubmenu > button {
            display: flex;
            align-items: center;
            gap: 24px;
        }
        .submenuChevron {
            margin-left: auto;
            color: var(--vscode-descriptionForeground);
        }
        .contextSubmenu {
            position: absolute;
            display: none;
            top: 0;
            left: calc(100% + 4px);
            min-width: 260px;
            max-width: min(420px, calc(100vw - 32px));
            padding: 4px;
            border: 1px solid var(--vscode-menu-border, var(--vscode-panel-border));
            border-radius: 4px;
            color: var(--vscode-menu-foreground, var(--vscode-foreground));
            background: var(--vscode-menu-background, var(--vscode-editor-background));
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28);
        }
        .contextSubmenu button {
            white-space: normal;
            line-height: 1.25;
        }
        .contextItemWithSubmenu.hasQuickFixes:hover > .contextSubmenu {
            display: block;
        }
        .settingsBackdrop {
            position: fixed;
            inset: 0;
            z-index: 20;
            display: grid;
            place-items: center;
            padding: 24px;
            background: rgba(0, 0, 0, 0.28);
        }
        .settingsBackdrop[hidden] {
            display: none;
        }
        .settingsDialog {
            width: min(760px, 100%);
            max-height: min(760px, calc(100vh - 48px));
            display: grid;
            grid-template-rows: auto 1fr;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 6px;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            box-shadow: 0 16px 40px rgba(0, 0, 0, 0.34);
            overflow: hidden;
        }
        .settingsHeader {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 12px 14px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        .settingsHeader h2 {
            margin: 0;
            font-size: 15px;
        }
        .settingsBody {
            min-height: 0;
            overflow: auto;
            padding: 14px;
        }
        .settingsSection + .settingsSection {
            margin-top: 18px;
        }
        .settingsSectionHeader {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 14px;
            margin-bottom: 8px;
        }
        .settingsSection h3 {
            margin: 0 0 8px;
            color: var(--vscode-descriptionForeground);
            font-size: 13px;
            font-weight: 600;
        }
        .settingsSectionHeader h3 {
            margin-bottom: 2px;
        }
        .settingsSource {
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
        }
        .settingsResetButton {
            white-space: nowrap;
        }
        .settingsResetButton:disabled {
            opacity: 0.45;
            cursor: default;
        }
        .settingsChoices {
            display: grid;
            gap: 8px;
        }
        .settingsChoice,
        .settingsRuleRow {
            display: flex;
            align-items: center;
            gap: 10px;
            min-height: 32px;
            padding: 6px 8px;
            border: 1px solid var(--vscode-panel-border);
            border-radius: 4px;
        }
        .settingsChoice input,
        .settingsRuleRow input {
            width: 18px;
            height: 18px;
            margin: 0;
        }
        .settingsRuleList {
            display: grid;
            gap: 6px;
        }
        .settingsRuleMain {
            min-width: 0;
            flex: 1;
        }
        .settingsRuleTitle {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-weight: 600;
        }
        .settingsRuleMeta {
            margin-top: 2px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
        }
        .settingsRuleSelect {
            min-width: 150px;
            height: 30px;
            color: var(--vscode-dropdown-foreground);
            background: var(--vscode-dropdown-background);
            border: 1px solid var(--vscode-dropdown-border);
            border-radius: 4px;
            padding: 0 8px;
        }
        .settingsEmpty {
            color: var(--vscode-descriptionForeground);
        }
        .settingsFooterActions {
            display: flex;
            justify-content: flex-end;
            margin-top: 18px;
            padding-top: 14px;
            border-top: 1px solid var(--vscode-panel-border);
        }
        @media (max-width: 860px) {
            .stats {
                gap: 8px 18px;
            }
            .body {
                grid-template-columns: 1fr;
            }
            aside {
                border-right: 0;
                border-bottom: 1px solid var(--vscode-panel-border);
                max-height: 180px;
            }
        }
    </style>
</head>
<body>
    <div class="shell">
        <header>
            <div>
                <h1>XLIDE Analysis Results</h1>
                <div class="subtle">${escapeHtml(model.workbookName)}</div>
            </div>
            <div class="headerActions">
                <button class="actionButton" id="analysisSettings" type="button">Analysis Settings</button>
                <button class="actionButton" id="copyReport" type="button">Copy Report</button>
                <button class="actionButton" id="copyJson" type="button">Copy JSON</button>
                <button class="actionButton" id="exportReport" type="button">Export Report</button>
                <button class="actionButton" id="exportJson" type="button">Export JSON</button>
            </div>
        </header>
        <section class="stats" aria-label="Analysis summary">
            ${summaryStatsHtml}
        </section>
        <div class="body">
            <aside>
                <h2>Modules</h2>
                <div class="moduleList">${moduleButtons}</div>
            </aside>
            <main>
                <div class="toolbar">
                    <div class="filters" aria-label="Filters">
                        ${filterButtons}
                        <button class="filterButton transient" type="button" data-show-hidden aria-pressed="false">Show Non-Tracked / Suppressed</button>
                    </div>
                    <div class="visibleCount" id="visibleCount"></div>
                </div>
                <div class="table" role="table" aria-label="Analysis findings">
                    <div class="tableHeader" role="row">
                        ${[
        sortHeaderHtml('severity', 'Severity'),
        sortHeaderHtml('status', 'Status'),
        sortHeaderHtml('location', 'Location'),
        sortHeaderHtml('rule', 'Rule'),
        sortHeaderHtml('evidence', 'Evidence'),
        sortHeaderHtml('message', 'Message'),
    ].join('')}
                    </div>
                    ${rowsHtml}
                </div>
            </main>
        </div>
    </div>
    <div class="contextMenu" id="rowContextMenu" hidden>
        <button type="button" data-context-action="askCopilot">Ask Copilot</button>
        <div class="contextItemWithSubmenu" id="quickFixMenuItem">
            <button type="button" data-context-action="quickFix">
                <span>Quick Fix</span>
                <span class="submenuChevron" aria-hidden="true">&gt;</span>
            </button>
            <div class="contextSubmenu" id="quickFixSubmenu" hidden></div>
        </div>
        <div class="contextDivider" id="suppressionDivider"></div>
        <button type="button" data-context-action="suppressBlock" data-suppress-scope="block">Ignore Block</button>
        <button type="button" data-context-action="suppressMember" data-suppress-scope="member">Ignore Sub/Function</button>
        <button type="button" data-context-action="suppressModule" data-suppress-scope="module">Ignore Module</button>
        <div class="contextDivider" id="trackingDivider"></div>
        <button type="button" data-context-action="setRuleTrackingWorkbook" id="trackingWorkbookAction">Untrack In Workbook</button>
        <button type="button" data-context-action="setRuleTrackingGlobal" id="trackingGlobalAction">Untrack Globally</button>
    </div>
    <div class="settingsBackdrop" id="analysisSettingsDialog" hidden>
        <section class="settingsDialog" role="dialog" aria-modal="true" aria-labelledby="analysisSettingsTitle">
            <div class="settingsHeader">
                <h2 id="analysisSettingsTitle">Analysis Settings</h2>
                <button class="secondaryButton" id="closeAnalysisSettings" type="button">Close</button>
            </div>
            <div class="settingsBody">
                <section class="settingsSection">
                    <div class="settingsSectionHeader">
                        <div>
                            <h3>Severities</h3>
                            <div class="settingsSource">Source: ${settingsSourceLabel(analysisSettings.visibleSeveritiesSource)}</div>
                        </div>
                        <button class="secondaryButton settingsResetButton" type="button" data-reset-analysis="severities" ${severitySourceIsWorkbook ? '' : 'disabled'}>Use Global Default</button>
                    </div>
                    <div class="settingsChoices" aria-label="Severity visibility">
                        ${ANALYSIS_SEVERITIES.map((severity) => `
                            <label class="settingsChoice">
                                <input type="checkbox" data-settings-severity="${severity}" ${visibleSeverities.includes(severity) ? 'checked' : ''}>
                                <span>${severityFilterLabel(severity)}</span>
                            </label>
                        `).join('')}
                    </div>
                </section>
                <section class="settingsSection">
                    <div class="settingsSectionHeader">
                        <div>
                            <h3>Workbook Rule Tracking</h3>
                            <div class="settingsSource">Source: ${settingsSourceLabel(analysisSettings.untrackedRulesSource)}</div>
                        </div>
                        <button class="secondaryButton settingsResetButton" type="button" data-reset-analysis="rules" ${rulesSourceIsWorkbook ? '' : 'disabled'}>Use Global Default</button>
                    </div>
                    <div class="settingsRuleList" aria-label="Workbook tracked analysis rules">
                        ${ruleSettingsHtml}
                    </div>
                </section>
                <section class="settingsSection">
                    <div class="settingsSectionHeader">
                        <div>
                            <h3>Rule Severities</h3>
                            <div class="settingsSource">Source: ${settingsSourceLabel(analysisSettings.ruleSeverityOverridesSource)}</div>
                        </div>
                        <button class="secondaryButton settingsResetButton" type="button" data-reset-analysis="ruleSeverities" ${ruleSeveritiesSourceIsWorkbook ? '' : 'disabled'}>Use Global Default</button>
                    </div>
                    <div class="settingsRuleList" aria-label="Rule severity overrides">
                        ${ruleSeveritySettingsHtml}
                    </div>
                </section>
                <div class="settingsFooterActions">
                    <button class="secondaryButton settingsResetButton" type="button" data-reset-analysis="all" ${anyAnalysisOverride ? '' : 'disabled'}>Use All Global Defaults</button>
                </div>
            </div>
        </section>
    </div>
    <div class="toast" id="toast" hidden></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const model = ${modelJson};
        const severityIds = ['error', 'warning', 'information'];
        const untrackedRules = new Set(model.untrackedRules ?? []);
        const ruleSeverityOverrides = new Map(Object.entries(model.ruleSeverityOverrides ?? {}));
        const persistedState = vscode.getState?.() ?? {};
        const hasPersistedUiState = persistedState.analysisUiInitialized === true;
        const hasPersistedSeverityState = hasPersistedUiState &&
            persistedState.analysisSettingsKey === model.analysisSettingsKey &&
            Array.isArray(persistedState.visibleSeverities);
        const visibleSeverities = new Set(
            hasPersistedSeverityState
                ? normalizeSeverityList(persistedState.visibleSeverities)
                : normalizeSeverityList(model.visibleSeverities ?? severityIds)
        );
        let activeModule = hasPersistedUiState && typeof persistedState.activeModule === 'string'
            ? persistedState.activeModule
            : 'all';
        let showHiddenItems = hasPersistedUiState && persistedState.showHiddenItems === true;
        let sortKey = hasPersistedUiState && isSortKey(persistedState.sortKey)
            ? persistedState.sortKey
            : 'severity';
        let sortDirection = hasPersistedUiState && persistedState.sortDirection === 'desc'
            ? 'desc'
            : 'asc';
        let settingsOpen = hasPersistedUiState && persistedState.settingsOpen === true;
        const rows = Array.from(document.querySelectorAll('.problemRow'));
        const sortHeaders = Array.from(document.querySelectorAll('[data-sort]'));
        const table = document.querySelector('.table');
        const showHiddenButton = document.querySelector('[data-show-hidden]');
        const visibleCount = document.getElementById('visibleCount');
        const toast = document.getElementById('toast');
        const contextMenu = document.getElementById('rowContextMenu');
        const quickFixMenuItem = document.getElementById('quickFixMenuItem');
        const quickFixSubmenu = document.getElementById('quickFixSubmenu');
        const suppressionDivider = document.getElementById('suppressionDivider');
        const trackingDivider = document.getElementById('trackingDivider');
        const trackingWorkbookAction = document.getElementById('trackingWorkbookAction');
        const trackingGlobalAction = document.getElementById('trackingGlobalAction');
        const settingsDialog = document.getElementById('analysisSettingsDialog');
        let contextRow = null;
        let contextMenuVisible = false;
        let persistFiltersTimer = undefined;

        function showToast(message) {
            toast.textContent = message;
            toast.hidden = false;
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => { toast.hidden = true; }, 1800);
        }

        function rowMatchesModule(row) {
            return activeModule === 'all' || row.dataset.module === activeModule;
        }

        function updateRows() {
            let count = 0;
            let shownHidden = 0;
            for (const row of rows) {
                const moduleVisible = rowMatchesModule(row);
                const severityVisible = visibleSeverities.has(row.dataset.severity);
                const untracked = row.dataset.tracked === 'no';
                const suppressed = row.dataset.suppressed === 'yes';
                const hiddenByStatus = untracked || suppressed;
                const visible = moduleVisible &&
                    severityVisible &&
                    (!hiddenByStatus || showHiddenItems);
                const untrackedVisible = visible && untracked;
                const suppressedVisible = visible && suppressed && !untracked;
                row.hidden = !visible;
                row.classList.toggle('hiddenVisible', untrackedVisible);
                row.classList.toggle('suppressedVisible', suppressedVisible);
                if (visible && !hiddenByStatus) {
                    count += 1;
                }
                if (visible && hiddenByStatus) {
                    shownHidden += 1;
                }
            }
            const details = [];
            if (showHiddenItems) {
                details.push(\`\${shownHidden} non-tracked/suppressed visible\`);
            }
            visibleCount.textContent = details.length > 0
                ? \`\${count} shown, \${details.join(', ')}\`
                : \`\${count} shown\`;
        }

        function sortRows() {
            const direction = sortDirection === 'asc' ? 1 : -1;
            rows.sort((left, right) => direction * compareRows(left, right, sortKey));
            for (const row of rows) {
                table.appendChild(row);
            }
        }

        function compareRows(left, right, key) {
            if (key === 'severity') {
                const severityOrder = { error: 0, warning: 1, information: 2 };
                return compareNumber(severityOrder[left.dataset.severity] ?? 4, severityOrder[right.dataset.severity] ?? 4)
                    || compareLocation(left, right);
            }
            if (key === 'status') {
                const statusOrder = { tracked: 0, untracked: 1, suppressed: 2 };
                return compareNumber(statusOrder[left.dataset.status] ?? 9, statusOrder[right.dataset.status] ?? 9)
                    || compareLocation(left, right);
            }
            if (key === 'rule') {
                return compareText(left.dataset.rule, right.dataset.rule) || compareLocation(left, right);
            }
            if (key === 'message') {
                return compareText(left.dataset.message, right.dataset.message) || compareLocation(left, right);
            }
            if (key === 'evidence') {
                return compareText(left.dataset.evidence, right.dataset.evidence) || compareLocation(left, right);
            }
            return compareLocation(left, right);
        }

        function compareLocation(left, right) {
            return compareNumber(Number(left.dataset.moduleOrder), Number(right.dataset.moduleOrder))
                || compareText(left.dataset.module, right.dataset.module)
                || compareNumber(Number(left.dataset.line), Number(right.dataset.line))
                || compareNumber(Number(left.dataset.column), Number(right.dataset.column));
        }

        function compareNumber(left, right) {
            return left === right ? 0 : left < right ? -1 : 1;
        }

        function compareText(left, right) {
            return String(left ?? '').localeCompare(String(right ?? ''));
        }

        function setActive(buttons, activeButton) {
            for (const button of buttons) {
                button.classList.toggle('active', button === activeButton);
            }
        }

        function persistUiState() {
            vscode.setState?.({
                ...(vscode.getState?.() ?? {}),
                analysisUiInitialized: true,
                activeModule,
                showHiddenItems,
                sortKey,
                sortDirection,
                settingsOpen,
                visibleSeverities: Array.from(visibleSeverities),
                analysisSettingsKey: model.analysisSettingsKey,
            });
        }

        function syncHiddenToggleButton() {
            showHiddenButton?.classList.toggle('active', showHiddenItems);
            showHiddenButton?.setAttribute('aria-pressed', showHiddenItems ? 'true' : 'false');
        }

        function syncSeverityFilterButtons() {
            for (const button of document.querySelectorAll('[data-severity-toggle]')) {
                const active = visibleSeverities.has(button.dataset.severityToggle);
                button.classList.toggle('active', active);
                button.setAttribute('aria-pressed', active ? 'true' : 'false');
            }
            for (const checkbox of document.querySelectorAll('[data-settings-severity]')) {
                checkbox.checked = visibleSeverities.has(checkbox.dataset.settingsSeverity);
            }
        }

        function syncSortHeaders() {
            for (const header of sortHeaders) {
                const active = header.dataset.sort === sortKey;
                const sortValue = active
                    ? (sortDirection === 'asc' ? 'ascending' : 'descending')
                    : 'none';
                const label = header.dataset.sortLabel ?? header.textContent?.trim() ?? 'column';
                const nextDirection = active && sortDirection === 'asc' ? 'descending' : 'ascending';
                header.setAttribute('aria-sort', sortValue);
                header.title = active
                    ? \`Sorted \${sortValue}. Click to sort \${nextDirection}.\`
                    : \`Sort by \${label}\`;
            }
        }

        function syncModuleFilterButtons() {
            const buttons = Array.from(document.querySelectorAll('[data-module-filter]'));
            let activeButton = buttons.find(button => button.dataset.moduleFilter === activeModule);
            if (!activeButton) {
                activeModule = 'all';
                activeButton = buttons.find(button => button.dataset.moduleFilter === 'all');
            }
            if (activeButton) {
                setActive(buttons, activeButton);
            }
        }

        function syncSettingsDialog() {
            if (settingsDialog) {
                settingsDialog.hidden = !settingsOpen;
            }
        }

        function normalizeSeverityList(value) {
            return Array.isArray(value)
                ? value.filter(item => severityIds.includes(item))
                : severityIds;
        }

        function isSortKey(value) {
            return value === 'severity' ||
                value === 'status' ||
                value === 'location' ||
                value === 'rule' ||
                value === 'evidence' ||
                value === 'message';
        }

        function persistSeveritiesDebounced() {
            clearTimeout(persistFiltersTimer);
            persistFiltersTimer = setTimeout(() => {
                vscode.postMessage({
                    type: 'updateSeveritySettings',
                    severities: Array.from(visibleSeverities),
                });
            }, 250);
        }

        function hideContextMenu() {
            contextMenu.hidden = true;
            quickFixSubmenu.hidden = true;
            contextRow = null;
            notifyContextMenu(false);
        }

        function setSettingsOpen(open) {
            settingsOpen = open;
            persistUiState();
            syncSettingsDialog();
        }

        function showContextMenu(row, x, y) {
            contextRow = row;
            const quickFixButton = contextMenu.querySelector('[data-context-action="quickFix"]');
            const fixes = quickFixTitlesForRow(row);
            if (quickFixButton) {
                const enabled = fixes.length > 0;
                quickFixButton.disabled = !enabled;
                quickFixButton.setAttribute('aria-disabled', enabled ? 'false' : 'true');
                quickFixMenuItem?.classList.toggle('hasQuickFixes', enabled);
            }
            renderQuickFixSubmenu(fixes);
            const suppressionScopes = suppressionScopesForRow(row);
            const rowSuppressed = row.dataset.suppressed === 'yes';
            let visibleSuppressActions = 0;
            for (const button of contextMenu.querySelectorAll('[data-suppress-scope]')) {
                const visible = suppressionScopes.has(button.dataset.suppressScope);
                button.hidden = !visible;
                button.disabled = rowSuppressed;
                button.setAttribute('aria-disabled', rowSuppressed ? 'true' : 'false');
                if (visible) {
                    visibleSuppressActions += 1;
                }
            }
            if (suppressionDivider) {
                suppressionDivider.hidden = visibleSuppressActions === 0;
            }
            const visibleTrackingActions = syncTrackingActions(row);
            if (trackingDivider) {
                trackingDivider.hidden = visibleTrackingActions === 0;
            }
            contextMenu.hidden = false;
            const rect = contextMenu.getBoundingClientRect();
            contextMenu.style.left = Math.min(x, window.innerWidth - rect.width - 8) + 'px';
            contextMenu.style.top = Math.min(y, window.innerHeight - rect.height - 8) + 'px';
            notifyContextMenu(true);
        }

        function notifyContextMenu(visible) {
            if (contextMenuVisible === visible) {
                return;
            }
            contextMenuVisible = visible;
            vscode.postMessage({ type: visible ? 'contextMenuOpened' : 'contextMenuClosed' });
        }

        function contextProblemIndex() {
            return contextRow ? Number(contextRow.dataset.openIndex) : undefined;
        }

        function contextProblemSuppressed() {
            return contextRow?.dataset.suppressed === 'yes';
        }

        function contextProblemTracked() {
            return contextRow?.dataset.tracked !== 'no';
        }

        function syncTrackingActions(row) {
            const hasRuleCode = String(row.dataset.ruleCode ?? '').trim().length > 0;
            const tracked = row.dataset.tracked !== 'no';
            const source = row.dataset.trackingSource === 'workbook' ? 'workbook' : 'global';
            return configureTrackingAction(trackingWorkbookAction, {
                hidden: !hasRuleCode || (!tracked && source !== 'workbook'),
                label: tracked ? 'Untrack In Workbook' : 'Track In Workbook',
            }) + configureTrackingAction(trackingGlobalAction, {
                hidden: !hasRuleCode || (!tracked && source !== 'global'),
                label: tracked ? 'Untrack Globally' : 'Track Globally',
            });
        }

        function configureTrackingAction(button, options) {
            if (!button) {
                return 0;
            }
            button.hidden = options.hidden;
            button.textContent = options.label;
            return options.hidden ? 0 : 1;
        }

        function trackingStatusLabel(source) {
            return source === 'workbook'
                ? 'Untracked In Workbook'
                : 'Untracked Globally';
        }

        function quickFixTitlesForRow(row) {
            try {
                const parsed = JSON.parse(row.dataset.quickFixes ?? '[]');
                return Array.isArray(parsed) ? parsed.filter(item => typeof item === 'string') : [];
            } catch {
                return [];
            }
        }

        function suppressionScopesForRow(row) {
            try {
                const parsed = JSON.parse(row.dataset.suppressionScopes ?? '[]');
                return new Set(
                    Array.isArray(parsed)
                        ? parsed.filter(item => item === 'block' || item === 'member' || item === 'module')
                        : []
                );
            } catch {
                return new Set();
            }
        }

        function renderQuickFixSubmenu(fixes) {
            quickFixSubmenu.innerHTML = '';
            quickFixSubmenu.hidden = fixes.length === 0;
            fixes.forEach((title, index) => {
                const button = document.createElement('button');
                button.type = 'button';
                button.dataset.contextAction = 'applyQuickFix';
                button.dataset.fixIndex = String(index);
                button.textContent = title;
                quickFixSubmenu.appendChild(button);
            });
        }

        function postOpenProblem(row) {
            vscode.postMessage({
                type: 'openProblem',
                index: Number(row.dataset.openIndex),
                suppressed: row.dataset.suppressed === 'yes',
                moduleName: row.dataset.module,
                moduleType: row.dataset.moduleType,
                line: Number(row.dataset.line),
                column: Number(row.dataset.column),
                endColumn: Number(row.dataset.endColumn),
                severity: row.dataset.severity,
                code: row.dataset.ruleCode,
                message: row.dataset.message,
            });
        }

        document.addEventListener('click', (event) => {
            if (event.button === 2) {
                return;
            }
            const contextButton = event.target.closest?.('[data-context-action]');
            if (contextButton) {
                const action = contextButton.dataset.contextAction;
                if (action === 'quickFix') {
                    return;
                }
                const index = contextProblemIndex();
                const suppressed = contextProblemSuppressed();
                const currentlyTracked = contextProblemTracked();
                hideContextMenu();
                if (typeof index !== 'number' || Number.isNaN(index)) {
                    return;
                }
                if (action === 'applyQuickFix') {
                    vscode.postMessage({
                        type: 'quickFixProblem',
                        index,
                        suppressed,
                        fixIndex: Number(contextButton.dataset.fixIndex ?? 0),
                    });
                } else if (action === 'askCopilot') {
                    vscode.postMessage({ type: 'askCopilot', index, suppressed });
                } else if (action === 'setRuleTrackingWorkbook' || action === 'setRuleTrackingGlobal') {
                    vscode.postMessage({
                        type: 'setRuleTracking',
                        index,
                        suppressed,
                        tracked: !currentlyTracked,
                        trackingScope: action === 'setRuleTrackingGlobal' ? 'global' : 'workbook',
                    });
                } else if (action === 'suppressBlock') {
                    vscode.postMessage({ type: 'suppressProblem', index, suppressed, scope: 'block' });
                } else if (action === 'suppressMember') {
                    vscode.postMessage({ type: 'suppressProblem', index, suppressed, scope: 'member' });
                } else if (action === 'suppressModule') {
                    vscode.postMessage({ type: 'suppressProblem', index, suppressed, scope: 'module' });
                }
                return;
            }
            hideContextMenu();
            const settingsButton = event.target.closest?.('#analysisSettings');
            if (settingsButton) {
                setSettingsOpen(true);
                return;
            }
            const closeSettingsButton = event.target.closest?.('#closeAnalysisSettings');
            if (closeSettingsButton) {
                setSettingsOpen(false);
                return;
            }
            if (event.target === settingsDialog) {
                setSettingsOpen(false);
                return;
            }
            const resetAnalysisButton = event.target.closest?.('[data-reset-analysis]');
            if (resetAnalysisButton) {
                if (resetAnalysisButton.disabled) {
                    return;
                }
                const scope = resetAnalysisButton.dataset.resetAnalysis;
                if (scope === 'severities') {
                    vscode.postMessage({ type: 'resetAnalysisSeverities' });
                } else if (scope === 'rules') {
                    vscode.postMessage({ type: 'resetAnalysisRuleTracking' });
                } else if (scope === 'ruleSeverities') {
                    vscode.postMessage({ type: 'resetAnalysisRuleSeverities' });
                } else if (scope === 'all') {
                    vscode.postMessage({ type: 'resetAnalysisSettings' });
                }
                return;
            }
            const settingsSeverity = event.target.closest?.('[data-settings-severity]');
            if (settingsSeverity) {
                const id = settingsSeverity.dataset.settingsSeverity;
                if (settingsSeverity.checked) {
                    visibleSeverities.add(id);
                } else {
                    visibleSeverities.delete(id);
                }
                persistUiState();
                syncSeverityFilterButtons();
                updateRows();
                persistSeveritiesDebounced();
                return;
            }
            const settingsRule = event.target.closest?.('[data-settings-rule-code]');
            if (settingsRule) {
                const code = settingsRule.dataset.settingsRuleCode;
                const tracked = settingsRule.checked === true;
                persistUiState();
                vscode.postMessage({
                    type: 'setRuleTracking',
                    code,
                    tracked,
                    trackingScope: 'workbook',
                });
                return;
            }
            const settingsRuleSeverity = event.target.closest?.('[data-settings-rule-severity-code]');
            if (settingsRuleSeverity) {
                const code = settingsRuleSeverity.dataset.settingsRuleSeverityCode;
                const severity = settingsRuleSeverity.value;
                if (code) {
                    if (severity === 'default') {
                        ruleSeverityOverrides.delete(code);
                    } else {
                        ruleSeverityOverrides.set(code, severity);
                    }
                    vscode.postMessage({
                        type: 'setRuleSeverityOverride',
                        code,
                        severity,
                    });
                }
                return;
            }
            const filterButton = event.target.closest?.('[data-severity-toggle]');
            if (filterButton) {
                const id = filterButton.dataset.severityToggle;
                if (visibleSeverities.has(id)) {
                    visibleSeverities.delete(id);
                } else {
                    visibleSeverities.add(id);
                }
                persistUiState();
                syncSeverityFilterButtons();
                updateRows();
                persistSeveritiesDebounced();
                return;
            }
            const sortButton = event.target.closest?.('[data-sort]');
            if (sortButton) {
                const nextKey = sortButton.dataset.sort;
                if (sortKey === nextKey) {
                    sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
                } else {
                    sortKey = nextKey;
                    sortDirection = 'asc';
                }
                persistUiState();
                sortRows();
                syncSortHeaders();
                updateRows();
                return;
            }
            const showHiddenButton = event.target.closest?.('[data-show-hidden]');
            if (showHiddenButton) {
                showHiddenItems = !showHiddenItems;
                persistUiState();
                syncHiddenToggleButton();
                updateRows();
                return;
            }
            const moduleButton = event.target.closest?.('[data-module-filter]');
            if (moduleButton) {
                activeModule = moduleButton.dataset.moduleFilter;
                persistUiState();
                setActive(document.querySelectorAll('[data-module-filter]'), moduleButton);
                updateRows();
                return;
            }
            const problemRow = event.target.closest?.('[data-open-index]');
            if (problemRow) {
                postOpenProblem(problemRow);
            }
        });

        document.addEventListener('contextmenu', (event) => {
            const problemRow = event.target.closest?.('[data-open-index]');
            if (!problemRow) {
                event.preventDefault();
                hideContextMenu();
                return;
            }
            event.preventDefault();
            showContextMenu(problemRow, event.clientX, event.clientY);
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                if (settingsOpen) {
                    setSettingsOpen(false);
                    return;
                }
                hideContextMenu();
            }
        });

        window.addEventListener('scroll', hideContextMenu, true);

        document.getElementById('copyReport').addEventListener('click', () => {
            vscode.postMessage({ type: 'copyText', text: model.plainText });
        });
        document.getElementById('copyJson').addEventListener('click', () => {
            vscode.postMessage({ type: 'copyText', text: JSON.stringify(model, null, 2) });
        });
        document.getElementById('exportReport').addEventListener('click', () => {
            vscode.postMessage({ type: 'exportText', text: model.plainText, extension: 'txt' });
        });
        document.getElementById('exportJson').addEventListener('click', () => {
            vscode.postMessage({ type: 'exportText', text: JSON.stringify(model, null, 2), extension: 'json' });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'copied') {
                showToast('Copied');
            } else if (event.data?.type === 'exported') {
                showToast('Exported');
            } else if (event.data?.type === 'suppressed') {
                showToast('Analysis ignore directive inserted');
            } else if (event.data?.type === 'quickFixed') {
                showToast('Quick fix applied');
            } else if (event.data?.type === 'quickFixUnavailable') {
                showToast('No quick fix available');
            } else if (event.data?.type === 'ruleTrackingChanged') {
                const code = String(event.data.code ?? '').toLowerCase();
                const tracked = event.data.tracked === true;
                if (code) {
                    const trackingSource = event.data.scope === 'global' ? 'global' : 'workbook';
                    if (tracked) {
                        untrackedRules.delete(code);
                    } else {
                        untrackedRules.add(code);
                    }
                    for (const row of rows) {
                        if (String(row.dataset.ruleCode ?? '').toLowerCase() === code) {
                            row.dataset.tracked = tracked ? 'yes' : 'no';
                            row.dataset.trackingSource = tracked ? 'tracked' : trackingSource;
                            const suppressed = row.dataset.suppressed === 'yes';
                            row.dataset.status = !tracked ? 'untracked' : (suppressed ? 'suppressed' : 'tracked');
                            const statusCell = row.querySelector('.status');
                            if (statusCell) {
                                statusCell.textContent = !tracked
                                    ? trackingStatusLabel(trackingSource)
                                    : (suppressed ? 'Suppressed' : 'Tracked');
                            }
                        }
                    }
                    sortRows();
                    updateRows();
                }
                const scope = event.data.scope === 'global' ? 'globally' : 'in workbook';
                showToast(code
                    ? (tracked ? 'Tracked ' : 'Untracked ') + code + ' ' + scope
                    : 'Rule tracking updated');
            } else if (event.data?.type === 'error') {
                showToast(event.data.error || 'XLIDE action failed');
            }
        });

        sortRows();
        syncSortHeaders();
        syncModuleFilterButtons();
        syncSeverityFilterButtons();
        syncHiddenToggleButton();
        syncSettingsDialog();
        updateRows();
        persistUiState();
    </script>
</body>
</html>`;
}

function renderWorkbookAnalysisErrorHtml(
    webview: vscode.Webview,
    workbookName: string,
    error: string,
): string {
    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview.cspSource} 'unsafe-inline';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE Analysis Error</title>
    <style>
        body {
            margin: 0;
            padding: 24px;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font-family: var(--vscode-font-family);
        }
        main {
            max-width: 900px;
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
            background: color-mix(in srgb, var(--vscode-inputValidation-errorBackground) 40%, transparent);
            padding: 12px;
            border-radius: 4px;
            line-height: 1.45;
            white-space: pre-wrap;
        }
        .help {
            margin-top: 12px;
            color: var(--vscode-descriptionForeground);
        }
    </style>
</head>
<body>
    <main>
        <h1>XLIDE Analysis Could Not Load</h1>
        <div class="subtle">${escapeHtml(workbookName)}</div>
        <div class="error">${escapeHtml(error)}</div>
        <div class="help">Fix or delete the workbook settings sidecar, then run analysis again.</div>
    </main>
</body>
</html>`;
}

function statHtml(value: string, label: string): string {
    return `<div class="stat"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>`;
}

function sortHeaderHtml(key: string, label: string): string {
    return `<button
        class="cell sortHeader"
        type="button"
        role="columnheader"
        data-sort="${escapeAttr(key)}"
        data-sort-label="${escapeAttr(label)}"
        aria-sort="none"
        title="Sort by ${escapeAttr(label)}"
    ><span class="sortHeaderText">${escapeHtml(label)}</span><span class="sortIndicator" aria-hidden="true"></span></button>`;
}

function problemFromOpenMessage(message: WorkbookAnalysisMessage): WorkbookAnalysisProblem | undefined {
    const moduleName = typeof message.moduleName === 'string' && message.moduleName.trim()
        ? message.moduleName
        : undefined;
    const line = positiveIntegerFromUnknown(message.line);
    const column = positiveIntegerFromUnknown(message.column);
    const rawEndColumn = positiveIntegerFromUnknown(message.endColumn);
    if (!moduleName || line === undefined || column === undefined) {
        return undefined;
    }
    return {
        moduleName,
        moduleType: typeof message.moduleType === 'string' && message.moduleType.trim()
            ? message.moduleType
            : 'standard',
        line,
        column,
        endColumn: Math.max(column + 1, rawEndColumn ?? column + 1),
        severity: analysisProblemSeverityFromUnknown(message.severity),
        code: typeof message.code === 'string' ? message.code : undefined,
        message: typeof message.message === 'string' ? message.message : '',
        suppressionScopes: [],
    };
}

function sameProblemLocation(left: WorkbookAnalysisProblem, right: WorkbookAnalysisProblem): boolean {
    return left.moduleName.toLowerCase() === right.moduleName.toLowerCase() &&
        left.line === right.line &&
        left.column === right.column;
}

function positiveIntegerFromUnknown(value: unknown): number | undefined {
    return Number.isInteger(value) && Number(value) > 0 ? Number(value) : undefined;
}

function analysisProblemSeverityFromUnknown(value: unknown): WorkbookAnalysisProblem['severity'] {
    return value === 'error' || value === 'warning' || value === 'information'
        ? value
        : 'information';
}

function workbookAnalysisSettingsKey(settings: EffectiveWorkbookAnalysisSettings): string {
    return JSON.stringify({
        visibleSeverities: [...settings.visibleSeverities],
        visibleSeveritiesSource: settings.visibleSeveritiesSource,
        untrackedRules: [...settings.untrackedRules].sort((left, right) => left.localeCompare(right)),
        untrackedRulesSource: settings.untrackedRulesSource,
        ruleSeverityOverrides: Object.entries(settings.ruleSeverityOverrides)
            .sort(([left], [right]) => left.localeCompare(right)),
        ruleSeverityOverridesSource: settings.ruleSeverityOverridesSource,
    });
}

function settingsSourceLabel(source: EffectiveWorkbookAnalysisSettings['visibleSeveritiesSource']): string {
    switch (source) {
        case 'workbook':
            return 'Workbook override';
        case 'machine':
            return 'VS Code machine setting';
        case 'default':
            return 'Built-in default';
        case 'unknown':
            return 'Unknown';
    }
}

function analysisRuleTrackingSourceForRow(
    tracked: boolean,
    source: EffectiveWorkbookAnalysisSettings['untrackedRulesSource'],
): 'tracked' | 'workbook' | 'global' {
    if (tracked) {
        return 'tracked';
    }
    return source === 'workbook' ? 'workbook' : 'global';
}

function analysisRuleTrackingStatusLabel(source: 'tracked' | 'workbook' | 'global'): string {
    switch (source) {
        case 'tracked':
            return 'Tracked';
        case 'workbook':
            return 'Untracked In Workbook';
        case 'global':
            return 'Untracked Globally';
    }
}

interface AnalysisRuleSetting {
    code: string;
    title: string;
    total: number;
    modules: Set<string>;
    severities: Set<string>;
}

function analysisRuleSettingsHtml(
    rows: readonly WorkbookAnalysisResultRow[],
    untrackedRules: readonly string[],
): string {
    const rules = collectAnalysisRuleSettings(rows, untrackedRules);
    if (rules.size === 0) {
        return '<div class="settingsEmpty">No analysis rules in the current result.</div>';
    }
    return [...rules.values()]
        .sort((left, right) => left.code.localeCompare(right.code))
        .map((rule) => {
            const tracked = isAnalysisRuleTracked(rule.code, untrackedRules);
            const severitySummary = [...rule.severities]
                .sort((left, right) => severityOrderForWebview(left) - severityOrderForWebview(right))
                .map(severityLabelForWebview)
                .join(', ');
            const moduleSummary = [...rule.modules].sort().join(', ');
            const pieces = [
                `${rule.total} ${rule.total === 1 ? 'finding' : 'findings'}`,
                severitySummary,
                moduleSummary,
            ].filter(Boolean);
            return `
                <label class="settingsRuleRow">
                    <input type="checkbox" data-settings-rule-code="${escapeAttr(rule.code)}" ${tracked ? 'checked' : ''}>
                    <span class="settingsRuleMain">
                        <span class="settingsRuleTitle">${escapeHtml(rule.title)}</span>
                        <span class="settingsRuleMeta">${escapeHtml([rule.code, ...pieces].join(' | '))}</span>
                    </span>
                </label>
            `;
        })
        .join('');
}

function analysisRuleSeveritySettingsHtml(
    rows: readonly WorkbookAnalysisResultRow[],
    ruleSeverityOverrides: AnalysisRuleSeverityOverrides,
): string {
    const rules = [...collectAnalysisRuleSettings(rows, Object.keys(ruleSeverityOverrides)).values()]
        .filter((rule) => allowedAnalysisRuleSeverityOverrides(rule.code).length > 0)
        .sort((left, right) => left.code.localeCompare(right.code));
    if (rules.length === 0) {
        return '<div class="settingsEmpty">No configurable analysis rule severities in the current result.</div>';
    }
    return rules.map((rule) => {
        const allowed = allowedAnalysisRuleSeverityOverrides(rule.code);
        const current: string = ruleSeverityOverrides[rule.code] ?? 'default';
        const severitySummary = [...rule.severities]
            .sort((left, right) => severityOrderForWebview(left) - severityOrderForWebview(right))
            .map(severityLabelForWebview)
            .join(', ');
        const pieces = [
            `${rule.total} ${rule.total === 1 ? 'finding' : 'findings'}`,
            severitySummary,
            `Allowed: ${allowed.map(severityOverrideLabelForWebview).join(', ')}`,
        ].filter(Boolean);
        return `
            <label class="settingsRuleRow">
                <span class="settingsRuleMain">
                    <span class="settingsRuleTitle">${escapeHtml(rule.title)}</span>
                    <span class="settingsRuleMeta">${escapeHtml([rule.code, ...pieces].join(' | '))}</span>
                </span>
                <select class="settingsRuleSelect" data-settings-rule-severity-code="${escapeAttr(rule.code)}">
                    <option value="default" ${current === 'default' ? 'selected' : ''}>Default</option>
                    ${allowed.map((severity) => `
                        <option value="${escapeAttr(severity)}" ${current === severity ? 'selected' : ''}>${escapeHtml(severityOverrideLabelForWebview(severity))}</option>
                    `).join('')}
                </select>
            </label>
        `;
    }).join('');
}

function collectAnalysisRuleSettings(
    rows: readonly WorkbookAnalysisResultRow[],
    extraCodes: readonly string[],
): Map<string, AnalysisRuleSetting> {
    const rules = new Map<string, AnalysisRuleSetting>();
    for (const row of rows) {
        const code = normalizeRuleCodeForWebview(row.code);
        if (!code) {
            continue;
        }
        let setting = rules.get(code);
        if (!setting) {
            setting = {
                code,
                title: row.ruleTitle || row.code || code,
                total: 0,
                modules: new Set<string>(),
                severities: new Set<string>(),
            };
            rules.set(code, setting);
        }
        setting.total += 1;
        setting.modules.add(row.moduleName);
        setting.severities.add(row.severity);
    }
    for (const rule of extraCodes) {
        const code = normalizeRuleCodeForWebview(rule);
        if (code && !rules.has(code)) {
            rules.set(code, {
                code,
                title: code,
                total: 0,
                modules: new Set<string>(),
                severities: new Set<string>(),
            });
        }
    }
    return rules;
}

function normalizeRuleCodeForWebview(code: unknown): string | undefined {
    return typeof code === 'string' ? code.trim().toLowerCase() || undefined : undefined;
}

function severityOrderForWebview(severity: string): number {
    switch (severity) {
        case 'error':
            return 0;
        case 'warning':
            return 1;
        case 'information':
            return 2;
        default:
            return 9;
    }
}

function severityLabelForWebview(severity: string): string {
    switch (severity) {
        case 'error':
            return 'Error';
        case 'warning':
            return 'Warning';
        case 'information':
            return 'Information';
        default:
            return severity;
    }
}

function severityOverrideLabelForWebview(severity: string): string {
    return severity === 'off' ? 'Off' : severityLabelForWebview(severity);
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

function severityFilterLabel(severity: AnalysisSeverityFilter): string {
    switch (severity) {
        case 'error':
            return 'Errors';
        case 'warning':
            return 'Warnings';
        case 'information':
            return 'Information';
    }
}

function randomNonce(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let nonce = '';
    for (let i = 0; i < 32; i++) {
        nonce += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return nonce;
}

function sanitizeFileName(value: string): string {
    return value.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/[. ]+$/g, '');
}

function workbookAnalysisResultsPanelKey(filePath: string): string {
    const normalized = path.normalize(filePath);
    return process.platform === 'win32' ? normalized.toLowerCase() : normalized;
}
