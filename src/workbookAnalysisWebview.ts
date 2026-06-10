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
    isAnalysisRuleTracked,
    type AnalysisSeverityFilter,
} from './analysisSettingsCore';
import { diagnosticMetadataForCode } from './analyzer';
import { setGlobalAnalysisRuleTracked } from './analysisOptions';
import {
    effectiveWorkbookAnalysisSettings,
    resetWorkbookAnalysisRuleTracking,
    setWorkbookAnalysisRuleTracked,
    type EffectiveWorkbookAnalysisSettings,
} from './workbookAnalysisSettings';
import { sanitizeFileName } from './moduleExport';
import { settingsPathForWorkbook } from './workbookSettings';
import { decodeModuleUri, sameWorkbookPath, XLIDE_SCHEME } from './xlideFileSystem';
import { measurePerformance } from './performanceTrace';
import { escapeAttr, escapeHtml, randomNonce, scriptJson } from './webview/html';
import {
    renderWebviewErrorPageHtml,
    statHtml,
    webviewHeadHtml,
    WEBVIEW_TOAST_CSS,
    WEBVIEW_TOAST_HTML,
    WEBVIEW_TOAST_SCRIPT,
} from './webview/page';
import { bridgeWebviewMessages, createWebviewPanelRegistry } from './webview/panelRegistry';
import { DebouncedRefresher } from './webview/refresh';
import { errorMessage } from './util/errors';

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
    refresh: () => Promise<void>;
    setResult: (result: WorkbookAnalysisResult) => void;
    showErrorPage: (error: string) => void;
}

const WORKBOOK_ANALYSIS_REFRESH_DELAY_MS = 350;
const WORKBOOK_ANALYSIS_TEXT_CHANGE_REFRESH_DELAY_MS = 1200;

const openWorkbookAnalysisResultsPanels = createWebviewPanelRegistry<OpenWorkbookAnalysisResultsPanelEntry>();

export function openWorkbookAnalysisResults(
    context: vscode.ExtensionContext,
    result: WorkbookAnalysisResult,
    options: WorkbookAnalysisResultsOptions = {},
): vscode.WebviewPanel {
    const existing = openWorkbookAnalysisResultsPanels.get(result.filePath);
    if (existing) {
        existing.options = options;
        existing.setResult(result);
        existing.panel.reveal(vscode.ViewColumn.Active);
        void existing.refresh().catch((err) => {
            existing.showErrorPage(errorMessage(err));
        });
        return existing.panel;
    }

    let currentResult = result;
    let currentModel = buildWorkbookAnalysisResultsModel(currentResult);
    let disposed = false;
    let htmlRendered = false;
    let ignoreOwnSettingsRefresh = false;
    let ignoreOwnSettingsRefreshTimer: ReturnType<typeof setTimeout> | undefined;
    const panel = vscode.window.createWebviewPanel(
        'xlideWorkbookAnalysisResults',
        `XLIDE Analysis: ${currentModel.workbookName}`,
        vscode.ViewColumn.Active,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
        },
    );

    const updateResult = (nextResult: WorkbookAnalysisResult): void => {
        currentResult = nextResult;
        currentModel = buildWorkbookAnalysisResultsModel(currentResult);
    };

    /** Full document render: first paint and recovery from the error page. */
    const renderPanel = async (): Promise<void> => {
        await measurePerformance('workbookAnalysis.renderPanel', currentModel.workbookName, async () => {
        const analysisSettings = await effectiveWorkbookAnalysisSettings(currentResult.filePath);
        if (disposed) { return; }
        panel.title = `XLIDE Analysis: ${currentModel.workbookName}`;
        panel.webview.html = renderWorkbookAnalysisResultsHtml(
            currentModel,
            analysisSettings,
        );
        htmlRendered = true;
        });
    };

    /** Non-destructive refresh: posts the new model for client-side re-render. */
    const postModelUpdate = async (): Promise<void> => {
        await measurePerformance('workbookAnalysis.postModelUpdate', currentModel.workbookName, async () => {
        const analysisSettings = await effectiveWorkbookAnalysisSettings(currentResult.filePath);
        if (disposed) { return; }
        panel.title = `XLIDE Analysis: ${currentModel.workbookName}`;
        await panel.webview.postMessage({
            type: 'model',
            model: buildWorkbookAnalysisClientModel(currentModel, analysisSettings),
        });
        });
    };

    const refreshView = (): Promise<void> => htmlRendered ? postModelUpdate() : renderPanel();

    const showErrorPage = (error: string): void => {
        htmlRendered = false;
        panel.webview.html = renderWorkbookAnalysisErrorHtml(currentModel.workbookName, error);
    };

    const entry: OpenWorkbookAnalysisResultsPanelEntry = {
        panel,
        options,
        refresh: refreshView,
        setResult: updateResult,
        showErrorPage,
    };
    openWorkbookAnalysisResultsPanels.set(result.filePath, entry);

    const refresher = new DebouncedRefresher({
        refresh: async () => {
            if (entry.options.onRefreshResult) {
                const nextResult = await entry.options.onRefreshResult();
                if (disposed) { return; }
                updateResult(nextResult);
            }
            if (!disposed) {
                await refreshView();
            }
        },
        onError: (err) => {
            const error = errorMessage(err);
            void panel.webview.postMessage({ type: 'error', error });
        },
        defaultDelayMs: WORKBOOK_ANALYSIS_REFRESH_DELAY_MS,
    });

    const scheduleRefresh = (delayMs?: number): void => refresher.schedule(delayMs);

    const refreshAfterAnalysisMutation = (): Promise<void> => refresher.refreshNow();

    const ignoreOwnSidecarRefreshBriefly = (): void => {
        ignoreOwnSettingsRefresh = true;
        if (ignoreOwnSettingsRefreshTimer) {
            clearTimeout(ignoreOwnSettingsRefreshTimer);
        }
        ignoreOwnSettingsRefreshTimer = setTimeout(() => {
            ignoreOwnSettingsRefresh = false;
            ignoreOwnSettingsRefreshTimer = undefined;
        }, 1000);
    };

    const scheduleSidecarRefresh = (): void => {
        if (ignoreOwnSettingsRefresh) {
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

    const reportText = (json: boolean): string => json
        ? JSON.stringify(currentModel, null, 2)
        : buildWorkbookAnalysisPlainText(currentModel);

    void renderPanel().catch((err) => {
        showErrorPage(errorMessage(err));
    });
    const messageSub = bridgeWebviewMessages(
        panel.webview,
        async (message: WorkbookAnalysisMessage) => {
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
                    if (scope === 'workbook') {
                        ignoreOwnSidecarRefreshBriefly();
                    }
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
            if (message.type === 'resetAnalysisRuleTracking') {
                ignoreOwnSidecarRefreshBriefly();
                await resetWorkbookAnalysisRuleTracking(currentResult.filePath);
                await refreshAfterAnalysisMutation();
                return;
            }
            if (message.type === 'copyReport' || message.type === 'copyJson') {
                await vscode.env.clipboard.writeText(reportText(message.type === 'copyJson'));
                await panel.webview.postMessage({ type: 'copied' });
                return;
            }
            if (message.type === 'exportReport' || message.type === 'exportJson') {
                const json = message.type === 'exportJson';
                const extension = json ? 'json' : 'txt';
                const target = await vscode.window.showSaveDialog({
                    title: json ? 'Export XLIDE Analysis JSON' : 'Export XLIDE Analysis Report',
                    defaultUri: vscode.Uri.file(path.join(
                        path.dirname(currentModel.filePath),
                        `${sanitizeFileName(currentModel.workbookName)}.xlide-analysis.${extension}`,
                    )),
                    filters: json ? { JSON: ['json'] } : { Text: ['txt'] },
                });
                if (target) {
                    await vscode.workspace.fs.writeFile(
                        target,
                        Buffer.from(reportText(json), 'utf8'),
                    );
                    await panel.webview.postMessage({ type: 'exported' });
                }
            }
        },
    );

    const configSub = vscode.workspace.onDidChangeConfiguration((e) => {
        if (e.affectsConfiguration('xlide.analysis') || e.affectsConfiguration('xlide.diagnostics')) {
            scheduleRefresh();
        }
    });
    const textChangeSub = vscode.workspace.onDidChangeTextDocument((e) => {
        if (isWorkbookDocument(e.document, currentResult.filePath)) {
            scheduleRefresh(WORKBOOK_ANALYSIS_TEXT_CHANGE_REFRESH_DELAY_MS);
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
        openWorkbookAnalysisResultsPanels.delete(result.filePath);
        refresher.dispose();
        if (ignoreOwnSettingsRefreshTimer) {
            clearTimeout(ignoreOwnSettingsRefreshTimer);
            ignoreOwnSettingsRefreshTimer = undefined;
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

interface WorkbookAnalysisClientRow {
    index: number;
    suppressed: boolean;
    moduleName: string;
    moduleType: string;
    moduleOrder: number;
    severity: WorkbookAnalysisResultRow['severity'];
    vbeCompileEquivalent: boolean;
    line: number;
    column: number;
    endColumn: number;
    rule: string;
    ruleCode: string;
    message: string;
    evidence: string;
    quickFixTitles: string[];
    suppressionScopes: string[];
    tracked: boolean;
    trackingSource: 'tracked' | 'workbook' | 'global';
    statusKey: 'tracked' | 'untracked' | 'suppressed';
    statusLabel: string;
    location: string;
}

interface WorkbookAnalysisClientModel {
    workbookName: string;
    totalProblems: number;
    errorCount: number;
    warningCount: number;
    informationCount: number;
    suppressedCount: number;
    untrackedCount: number;
    moduleCount: number;
    groups: Array<{ moduleName: string; moduleIcon: string; moduleTypeLabel: string; total: number }>;
    rows: WorkbookAnalysisClientRow[];
    visibleSeverities: readonly AnalysisSeverityFilter[];
    untrackedRules: readonly string[];
    analysisSettingsKey: string;
    rulesSourceIsWorkbook: boolean;
    workbookUntrackedRules: Array<{ code: string; title: string }>;
}

/** Everything the webview script needs to render or re-render the panel body. */
function buildWorkbookAnalysisClientModel(
    model: WorkbookAnalysisResultsModel,
    analysisSettings: EffectiveWorkbookAnalysisSettings,
): WorkbookAnalysisClientModel {
    const untrackedRules = analysisSettings.untrackedRules;
    const moduleOrder = new Map(model.groups.map((group, index) => [group.moduleName.toLowerCase(), index]));
    const rows = [...model.rows, ...model.suppressedRows].map((row): WorkbookAnalysisClientRow => {
        const tracked = isAnalysisRuleTracked(row.code, untrackedRules);
        const trackingSource = analysisRuleTrackingSourceForRow(
            tracked,
            row.code,
            analysisSettings.workbookUntrackedRules,
        );
        return {
            index: row.index,
            suppressed: row.suppressed,
            moduleName: row.moduleName,
            moduleType: row.moduleType,
            moduleOrder: moduleOrder.get(row.moduleName.toLowerCase()) ?? 9999,
            severity: row.severity,
            vbeCompileEquivalent: row.vbeCompileEquivalent,
            line: row.line,
            column: row.column,
            endColumn: row.endColumn,
            rule: row.code || row.ruleTitle,
            ruleCode: row.code,
            message: row.message,
            evidence: row.diagnosticKind,
            quickFixTitles: row.quickFixTitles,
            suppressionScopes: row.suppressionScopes,
            tracked,
            trackingSource,
            statusKey: !tracked ? 'untracked' : (row.suppressed ? 'suppressed' : 'tracked'),
            statusLabel: !tracked
                ? analysisRuleTrackingStatusLabel(trackingSource)
                : (row.suppressed ? 'Suppressed' : 'Tracked'),
            location: row.location,
        };
    });
    return {
        workbookName: model.workbookName,
        totalProblems: model.totalProblems,
        errorCount: model.errorCount,
        warningCount: model.warningCount,
        informationCount: model.rows.filter((row) => row.severity === 'information').length,
        suppressedCount: rows.filter((row) => row.suppressed && row.tracked).length,
        untrackedCount: rows.filter((row) => !row.tracked).length,
        moduleCount: model.moduleCount,
        groups: model.groups.map((group) => ({
            moduleName: group.moduleName,
            moduleIcon: group.moduleIcon,
            moduleTypeLabel: group.moduleTypeLabel,
            total: group.total,
        })),
        rows,
        visibleSeverities: analysisSettings.visibleSeverities,
        untrackedRules: analysisSettings.untrackedRules,
        analysisSettingsKey: workbookAnalysisSettingsKey(analysisSettings),
        rulesSourceIsWorkbook: analysisSettings.untrackedRulesSource === 'workbook',
        workbookUntrackedRules: [...analysisSettings.workbookUntrackedRules]
            .sort((left, right) => left.localeCompare(right))
            .map((code) => ({ code, title: diagnosticMetadataForCode(code)?.title ?? code })),
    };
}

export function renderWorkbookAnalysisResultsHtml(
    model: WorkbookAnalysisResultsModel,
    analysisSettings: EffectiveWorkbookAnalysisSettings,
): string {
    const nonce = randomNonce();
    const clientModel = buildWorkbookAnalysisClientModel(model, analysisSettings);
    const modelJson = scriptJson(clientModel);
    const rowsHtml = clientModel.rows.length === 0
        ? '<div class="empty">No analysis findings.</div>'
        : clientModel.rows.map((row) => `
            <button
                class="problemRow severity-${escapeAttr(row.severity)}"
                type="button"
                data-open-index="${row.index}"
                data-suppressed="${row.suppressed ? 'yes' : 'no'}"
                data-status="${row.statusKey}"
                data-module="${escapeAttr(row.moduleName)}"
                data-module-type="${escapeAttr(row.moduleType)}"
                data-module-order="${row.moduleOrder}"
                data-severity="${escapeAttr(row.severity)}"
                data-compile="${row.vbeCompileEquivalent ? 'yes' : 'no'}"
                data-line="${row.line}"
                data-column="${row.column}"
                data-end-column="${row.endColumn}"
                data-rule="${escapeAttr(row.rule)}"
                data-rule-code="${escapeAttr(row.ruleCode)}"
                data-message="${escapeAttr(row.message)}"
                data-evidence="${escapeAttr(row.evidence)}"
                data-quick-fixes="${escapeAttr(JSON.stringify(row.quickFixTitles))}"
                data-suppression-scopes="${escapeAttr(JSON.stringify(row.suppressionScopes))}"
                data-tracked="${row.tracked ? 'yes' : 'no'}"
                data-tracking-source="${escapeAttr(row.trackingSource)}"
            >
                <span class="cell severity">${escapeHtml(row.severity)}</span>
                <span class="cell status">${escapeHtml(row.statusLabel)}</span>
                <span class="cell location">${escapeHtml(row.location)}</span>
                <span class="cell code">${escapeHtml(row.rule)}</span>
                <span class="cell kind">${escapeHtml(row.evidence)}</span>
                <span class="cell message">${escapeHtml(row.message)}</span>
            </button>
        `).join('');
    const moduleButtons = [
        `<button class="moduleFilter active" type="button" data-module-filter="all">All modules <span>${clientModel.totalProblems}</span></button>`,
        ...clientModel.groups.map((group) => `
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
        const active = clientModel.visibleSeverities.includes(severity);
        return `<button class="filterButton${active ? ' active' : ''}" type="button" data-severity-toggle="${severity}" aria-pressed="${active ? 'true' : 'false'}">${severityFilterLabel(severity)}</button>`;
    }).join('');
    const workbookUntrackedRulesHtml = workbookUntrackedRulesSettingsHtml(clientModel.workbookUntrackedRules);
    const rulesSourceIsWorkbook = clientModel.rulesSourceIsWorkbook;
    const summaryStatsHtml = [
        statHtml(String(clientModel.errorCount), 'Errors'),
        statHtml(String(clientModel.warningCount), 'Warnings'),
        statHtml(String(clientModel.informationCount), 'Information'),
        statHtml(String(clientModel.suppressedCount), 'Suppressed'),
        statHtml(String(clientModel.untrackedCount), 'Untracked'),
        statHtml(String(clientModel.moduleCount), 'Modules'),
    ].join('');

    return `<!DOCTYPE html>
<html lang="en">
<head>
    ${webviewHeadHtml(nonce, 'XLIDE Analysis Results')}
    <style nonce="${nonce}">
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
        ${WEBVIEW_TOAST_CSS}
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
            overflow-y: auto;
            overflow-x: hidden;
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
        .settingsHint {
            margin: -2px 0 8px;
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
            line-height: 1.4;
        }
        .settingsHeaderActions {
            display: flex;
            flex-wrap: wrap;
            justify-content: flex-end;
            gap: 8px;
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
        .settingsRuleTracking {
            display: flex;
            align-items: center;
            gap: 10px;
            min-width: 0;
            flex: 1;
        }
        .settingsRuleMain {
            min-width: 0;
            flex: 1;
        }
        .settingsRuleControls {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-left: auto;
        }
        .settingsRuleControlLabel,
        .settingsRuleFixed {
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
            white-space: nowrap;
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
        .settingsTable {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            border: 1px solid var(--vscode-panel-border);
        }
        .settingsTable th,
        .settingsTable td {
            padding: 7px 9px;
            border-bottom: 1px solid var(--vscode-panel-border);
            text-align: left;
            vertical-align: middle;
        }
        .settingsTable th {
            color: var(--vscode-descriptionForeground);
            font-size: 12px;
            font-weight: 600;
        }
        .settingsTable td {
            overflow-wrap: anywhere;
        }
        .settingsTableCode {
            width: 34%;
            color: var(--vscode-descriptionForeground);
            font-family: var(--vscode-editor-font-family);
        }
        .settingsTableAction {
            width: 96px;
            text-align: right;
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
                            <h3>Workbook Untracked Rules</h3>
                            <div class="settingsSource">Source: ${rulesSourceIsWorkbook ? 'Workbook settings' : 'No workbook override'}</div>
                        </div>
                        <button class="secondaryButton settingsResetButton" type="button" data-reset-analysis="rules" ${rulesSourceIsWorkbook ? '' : 'disabled'}>Track All</button>
                    </div>
                    <div id="workbookUntrackedRules">${workbookUntrackedRulesHtml}</div>
                </section>
            </div>
        </section>
    </div>
    ${WEBVIEW_TOAST_HTML}
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        ${WEBVIEW_TOAST_SCRIPT}
        let model = ${modelJson};
        const severityIds = ['error', 'warning', 'information'];
        let untrackedRules = new Set(model.untrackedRules ?? []);
        let analysisSettingsKey = model.analysisSettingsKey;
        let visibleSeverities = new Set(normalizeSeverityList(model.visibleSeverities ?? severityIds));
        let activeModule = 'all';
        let showHiddenItems = false;
        let sortKey = 'severity';
        let sortDirection = 'asc';
        let settingsOpen = false;
        let rows = Array.from(document.querySelectorAll('.problemRow'));
        const sortHeaders = Array.from(document.querySelectorAll('[data-sort]'));
        const table = document.querySelector('.table');
        const statsSection = document.querySelector('.stats');
        const moduleList = document.querySelector('.moduleList');
        const workbookSubtitle = document.querySelector('header .subtle');
        const showHiddenButton = document.querySelector('[data-show-hidden]');
        const visibleCount = document.getElementById('visibleCount');
        const contextMenu = document.getElementById('rowContextMenu');
        const quickFixMenuItem = document.getElementById('quickFixMenuItem');
        const quickFixSubmenu = document.getElementById('quickFixSubmenu');
        const suppressionDivider = document.getElementById('suppressionDivider');
        const trackingDivider = document.getElementById('trackingDivider');
        const trackingWorkbookAction = document.getElementById('trackingWorkbookAction');
        const trackingGlobalAction = document.getElementById('trackingGlobalAction');
        const settingsDialog = document.getElementById('analysisSettingsDialog');
        const settingsSource = document.querySelector('.settingsSource');
        const settingsResetButton = document.querySelector('[data-reset-analysis]');
        const workbookUntrackedRulesContainer = document.getElementById('workbookUntrackedRules');
        let contextRow = null;

        function applyModel(next) {
            model = next;
            untrackedRules = new Set(next.untrackedRules ?? []);
            if (next.analysisSettingsKey !== analysisSettingsKey) {
                analysisSettingsKey = next.analysisSettingsKey;
                visibleSeverities = new Set(normalizeSeverityList(next.visibleSeverities ?? severityIds));
            }
            if (workbookSubtitle) {
                workbookSubtitle.textContent = next.workbookName;
            }
            renderStats();
            renderModuleFilters();
            renderRows();
            renderWorkbookUntrackedRules();
            reattachContextRow();
            sortRows();
            syncSortHeaders();
            syncModuleFilterButtons();
            syncSeverityFilterButtons();
            syncHiddenToggleButton();
            updateRows();
        }

        function element(tag, className, text) {
            const node = document.createElement(tag);
            if (className) {
                node.className = className;
            }
            if (text !== undefined) {
                node.textContent = text;
            }
            return node;
        }

        function renderStats() {
            statsSection.textContent = '';
            for (const [value, label] of [
                [model.errorCount, 'Errors'],
                [model.warningCount, 'Warnings'],
                [model.informationCount, 'Information'],
                [model.suppressedCount, 'Suppressed'],
                [model.untrackedCount, 'Untracked'],
                [model.moduleCount, 'Modules'],
            ]) {
                const stat = element('div', 'stat');
                stat.appendChild(element('strong', undefined, String(value)));
                stat.appendChild(element('span', undefined, label));
                statsSection.appendChild(stat);
            }
        }

        function buildModuleFilterButton(filter, labelNode, count) {
            const button = element('button', 'moduleFilter');
            button.type = 'button';
            button.dataset.moduleFilter = filter;
            button.appendChild(labelNode);
            button.appendChild(element('span', undefined, String(count)));
            return button;
        }

        function renderModuleFilters() {
            moduleList.textContent = '';
            moduleList.appendChild(buildModuleFilterButton(
                'all',
                document.createTextNode('All modules '),
                model.totalProblems,
            ));
            for (const group of model.groups) {
                const identity = element('span', 'moduleIdentity');
                const icon = element('span', 'moduleIcon', group.moduleIcon);
                icon.title = group.moduleTypeLabel;
                identity.appendChild(icon);
                identity.appendChild(element('span', 'moduleName', group.moduleName));
                moduleList.appendChild(buildModuleFilterButton(group.moduleName, identity, group.total));
            }
        }

        function renderRows() {
            for (const row of rows) {
                row.remove();
            }
            const empty = table.querySelector('.empty');
            if (empty) {
                empty.remove();
            }
            rows = model.rows.map(buildProblemRow);
            if (rows.length === 0) {
                table.appendChild(element('div', 'empty', 'No analysis findings.'));
                return;
            }
            for (const row of rows) {
                table.appendChild(row);
            }
        }

        function buildProblemRow(row) {
            const button = element('button', 'problemRow severity-' + row.severity);
            button.type = 'button';
            button.dataset.openIndex = String(row.index);
            button.dataset.suppressed = row.suppressed ? 'yes' : 'no';
            button.dataset.status = row.statusKey;
            button.dataset.module = row.moduleName;
            button.dataset.moduleType = row.moduleType;
            button.dataset.moduleOrder = String(row.moduleOrder);
            button.dataset.severity = row.severity;
            button.dataset.compile = row.vbeCompileEquivalent ? 'yes' : 'no';
            button.dataset.line = String(row.line);
            button.dataset.column = String(row.column);
            button.dataset.endColumn = String(row.endColumn);
            button.dataset.rule = row.rule;
            button.dataset.ruleCode = row.ruleCode;
            button.dataset.message = row.message;
            button.dataset.evidence = row.evidence;
            button.dataset.quickFixes = JSON.stringify(row.quickFixTitles ?? []);
            button.dataset.suppressionScopes = JSON.stringify(row.suppressionScopes ?? []);
            button.dataset.tracked = row.tracked ? 'yes' : 'no';
            button.dataset.trackingSource = row.trackingSource;
            button.appendChild(element('span', 'cell severity', row.severity));
            button.appendChild(element('span', 'cell status', row.statusLabel));
            button.appendChild(element('span', 'cell location', row.location));
            button.appendChild(element('span', 'cell code', row.rule));
            button.appendChild(element('span', 'cell kind', row.evidence));
            button.appendChild(element('span', 'cell message', row.message));
            return button;
        }

        function renderWorkbookUntrackedRules() {
            settingsSource.textContent = 'Source: ' + (model.rulesSourceIsWorkbook ? 'Workbook settings' : 'No workbook override');
            settingsResetButton.disabled = !model.rulesSourceIsWorkbook;
            workbookUntrackedRulesContainer.textContent = '';
            const rules = model.workbookUntrackedRules ?? [];
            if (rules.length === 0) {
                workbookUntrackedRulesContainer.appendChild(
                    element('div', 'settingsEmpty', 'No workbook rules are manually untracked.'),
                );
                return;
            }
            const rulesTable = element('table', 'settingsTable');
            const head = element('thead');
            const headRow = element('tr');
            for (const [className, label] of [
                ['settingsTableCode', 'Rule'],
                [undefined, 'Title'],
                ['settingsTableAction', 'Action'],
            ]) {
                const cell = element('th', className, label);
                cell.scope = 'col';
                headRow.appendChild(cell);
            }
            head.appendChild(headRow);
            rulesTable.appendChild(head);
            const body = element('tbody');
            for (const rule of rules) {
                const ruleRow = element('tr');
                ruleRow.appendChild(element('td', 'settingsTableCode', rule.code));
                ruleRow.appendChild(element('td', undefined, rule.title));
                const action = element('td', 'settingsTableAction');
                const track = element('button', 'secondaryButton', 'Track');
                track.type = 'button';
                track.dataset.settingsTrackRuleCode = rule.code;
                action.appendChild(track);
                ruleRow.appendChild(action);
                body.appendChild(ruleRow);
            }
            rulesTable.appendChild(body);
            workbookUntrackedRulesContainer.appendChild(rulesTable);
        }

        function reattachContextRow() {
            if (!contextRow) {
                return;
            }
            const previous = contextRow.dataset;
            const match = rows.find((row) =>
                row.dataset.module === previous.module &&
                row.dataset.line === previous.line &&
                row.dataset.column === previous.column &&
                row.dataset.ruleCode === previous.ruleCode);
            if (match) {
                contextRow = match;
            } else {
                hideContextMenu();
            }
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

        function hideContextMenu() {
            contextMenu.hidden = true;
            quickFixSubmenu.hidden = true;
            contextRow = null;
        }

        function setSettingsOpen(open) {
            settingsOpen = open;
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
                if (scope === 'rules') {
                    vscode.postMessage({ type: 'resetAnalysisRuleTracking' });
                }
                return;
            }
            const settingsTrackRule = event.target.closest?.('[data-settings-track-rule-code]');
            if (settingsTrackRule) {
                const code = settingsTrackRule.dataset.settingsTrackRuleCode;
                vscode.postMessage({
                    type: 'setRuleTracking',
                    code,
                    tracked: true,
                    trackingScope: 'workbook',
                });
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
                syncSeverityFilterButtons();
                updateRows();
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
                sortRows();
                syncSortHeaders();
                updateRows();
                return;
            }
            const showHiddenButton = event.target.closest?.('[data-show-hidden]');
            if (showHiddenButton) {
                showHiddenItems = !showHiddenItems;
                syncHiddenToggleButton();
                updateRows();
                return;
            }
            const moduleButton = event.target.closest?.('[data-module-filter]');
            if (moduleButton) {
                activeModule = moduleButton.dataset.moduleFilter;
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
            vscode.postMessage({ type: 'copyReport' });
        });
        document.getElementById('copyJson').addEventListener('click', () => {
            vscode.postMessage({ type: 'copyJson' });
        });
        document.getElementById('exportReport').addEventListener('click', () => {
            vscode.postMessage({ type: 'exportReport' });
        });
        document.getElementById('exportJson').addEventListener('click', () => {
            vscode.postMessage({ type: 'exportJson' });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'model') {
                applyModel(event.data.model);
            } else if (event.data?.type === 'copied') {
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
    </script>
</body>
</html>`;
}

function renderWorkbookAnalysisErrorHtml(
    workbookName: string,
    error: string,
): string {
    return renderWebviewErrorPageHtml({
        title: 'XLIDE Analysis Error',
        heading: 'XLIDE Analysis Could Not Load',
        subtitle: workbookName,
        error,
        help: 'Fix or delete the workbook settings sidecar, then run analysis again.',
    });
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

function analysisRuleTrackingSourceForRow(
    tracked: boolean,
    code: string | undefined,
    workbookUntrackedRules: readonly string[],
): 'tracked' | 'workbook' | 'global' {
    if (tracked) {
        return 'tracked';
    }
    const normalizedCode = typeof code === 'string' ? code.trim().toLowerCase() : '';
    if (normalizedCode && workbookUntrackedRules.includes(normalizedCode)) {
        return 'workbook';
    }
    return 'global';
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

function workbookUntrackedRulesSettingsHtml(
    workbookUntrackedRules: ReadonlyArray<{ code: string; title: string }>,
): string {
    if (workbookUntrackedRules.length === 0) {
        return '<div class="settingsEmpty">No workbook rules are manually untracked.</div>';
    }
    return `
        <table class="settingsTable">
            <thead>
                <tr>
                    <th class="settingsTableCode" scope="col">Rule</th>
                    <th scope="col">Title</th>
                    <th class="settingsTableAction" scope="col">Action</th>
                </tr>
            </thead>
            <tbody>
                ${workbookUntrackedRules.map((rule) => `
                        <tr>
                            <td class="settingsTableCode">${escapeHtml(rule.code)}</td>
                            <td>${escapeHtml(rule.title)}</td>
                            <td class="settingsTableAction">
                                <button class="secondaryButton" type="button" data-settings-track-rule-code="${escapeAttr(rule.code)}">Track</button>
                            </td>
                        </tr>
                    `).join('')}
            </tbody>
        </table>
    `;
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


