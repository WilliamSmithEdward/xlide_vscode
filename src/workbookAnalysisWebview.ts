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
import { WEBVIEW_BODY_CSS, xlideAccentPaletteCss } from './webview/styles';
import { renderWebviewTemplate } from './webview/templates';
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
            // Required: refreshes arrive as postMessage model updates, so the
            // stored html goes stale and filter/sort state lives in the DOM.
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

    // For MUTATING row actions (suppress / quick-fix / ask-Copilot / rule
    // tracking) the client also sends the finding's stable identity. If a
    // background refresh shifted the problem array underneath the panel, the
    // problem now at the sent index will not match that identity - refuse to act
    // so we never suppress or rewrite the WRONG finding.
    const verifiedMutationProblem = (message: WorkbookAnalysisMessage): WorkbookAnalysisProblem | undefined => {
        const indexed = problemAt(message.index, message.suppressed);
        if (!indexed) {
            return undefined;
        }
        const identity = problemFromOpenMessage(message);
        if (identity && !sameProblemLocation(indexed, identity)) {
            return undefined;
        }
        return indexed;
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
                const problem = verifiedMutationProblem(message);
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
                const problem = verifiedMutationProblem(message);
                if (problem && entry.options.onAskCopilot) {
                    await entry.options.onAskCopilot(problem, panel.viewColumn);
                }
                return;
            }
            if (message.type === 'quickFixProblem') {
                const problem = verifiedMutationProblem(message);
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
                const problem = verifiedMutationProblem(message);
                const code = typeof message.code === 'string' ? message.code : problem?.code;
                if (code) {
                    const scope = message.trackingScope === 'global' ? 'global' : 'workbook';
                    if (scope === 'workbook') {
                        ignoreOwnSidecarRefreshBriefly();
                    }
                    const update = scope === 'global'
                        ? await setGlobalAnalysisRuleTracked(code, message.tracked === true)
                        : await setWorkbookAnalysisRuleTracked(currentResult.filePath, code, message.tracked === true);
                    // Re-arm the suppression window now that the sidecar write has
                    // resolved, so a slow filesystem-watcher notification after the
                    // initial window still does not trigger a redundant self-refresh.
                    if (scope === 'workbook') {
                        ignoreOwnSidecarRefreshBriefly();
                    }
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
                ignoreOwnSidecarRefreshBriefly();
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
        // Derive all three counts uniformly from the same row set so the
        // Information stat cannot silently desync from Errors/Warnings.
        errorCount: model.rows.filter((row) => row.severity === 'error').length,
        warningCount: model.rows.filter((row) => row.severity === 'warning').length,
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

    return renderWebviewTemplate('assets/webview/workbookAnalysis.html', {
        head: webviewHeadHtml(nonce, 'XLIDE Analysis Results'),
        nonce,
        css: renderWebviewTemplate('assets/webview/workbookAnalysis.css', {
            accentPalette: xlideAccentPaletteCss(),
            bodyCss: WEBVIEW_BODY_CSS,
            toastCss: WEBVIEW_TOAST_CSS,
        }),
        workbookName: escapeHtml(model.workbookName),
        summaryStats: summaryStatsHtml,
        moduleButtons,
        filterButtons,
        sortHeaders: [
            sortHeaderHtml('severity', 'Severity'),
            sortHeaderHtml('status', 'Status'),
            sortHeaderHtml('location', 'Location'),
            sortHeaderHtml('rule', 'Rule'),
            sortHeaderHtml('evidence', 'Evidence'),
            sortHeaderHtml('message', 'Message'),
        ].join(''),
        rows: rowsHtml,
        rulesSourceLabel: rulesSourceIsWorkbook ? 'Workbook settings' : 'No workbook override',
        rulesResetDisabled: rulesSourceIsWorkbook ? '' : 'disabled',
        workbookUntrackedRules: workbookUntrackedRulesHtml,
        toastHtml: WEBVIEW_TOAST_HTML,
        js: renderWebviewTemplate('assets/webview/workbookAnalysis.js', {
            toastScript: WEBVIEW_TOAST_SCRIPT,
            modelJson,
        }),
    });
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
        left.column === right.column &&
        // Include the rule code: distinct findings can share an exact line:column,
        // so a mutating action (suppress/quick-fix) must not act on a co-located
        // finding with a different code after a background refresh shifts indices.
        (left.code ?? '').trim().toLowerCase() === (right.code ?? '').trim().toLowerCase();
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


