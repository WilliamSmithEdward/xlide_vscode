import * as vscode from 'vscode';
import * as path from 'path';
import type { WorkbookAnalysisProblem, WorkbookAnalysisResult } from './vbaWorkbookAnalysis';
import {
    buildWorkbookAnalysisPlainText,
    buildWorkbookAnalysisResultsModel,
    type WorkbookAnalysisResultsModel,
} from './workbookAnalysisResultsModel';
import {
    addUntrackedAnalysisRuleToConfig,
    ANALYSIS_SEVERITIES,
    isAnalysisRuleTracked,
    normalizeAnalysisVisibleSeverities,
    type AnalysisSeverityFilter,
    untrackedAnalysisRulesFromConfig,
    visibleAnalysisSeveritiesFromConfig,
} from './analysisOptions';
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

export function openWorkbookAnalysisResults(
    context: vscode.ExtensionContext,
    result: WorkbookAnalysisResult,
    options: WorkbookAnalysisResultsOptions = {},
): vscode.WebviewPanel {
    let currentResult = result;
    let currentModel = buildWorkbookAnalysisResultsModel(currentResult);
    let disposed = false;
    let refreshTimer: ReturnType<typeof setTimeout> | undefined;
    let ignoreNextVisibleSeverityConfigChange = false;
    let contextMenuOpen = false;
    let pendingRefreshAfterContextMenu = false;
    const panel = vscode.window.createWebviewPanel(
        'xlideWorkbookAnalysisResults',
        `XLIDE Analysis: ${currentModel.workbookName}`,
        vscode.ViewColumn.Beside,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
        },
    );

    const renderPanel = (): void => {
        currentModel = buildWorkbookAnalysisResultsModel(currentResult);
        panel.title = `XLIDE Analysis: ${currentModel.workbookName}`;
        panel.webview.html = renderWorkbookAnalysisResultsHtml(
            panel.webview,
            currentModel,
            visibleAnalysisSeveritiesFromConfig(),
            untrackedAnalysisRulesFromConfig(),
        );
    };

    const refreshPanel = async (): Promise<void> => {
        if (disposed) { return; }
        if (contextMenuOpen) {
            pendingRefreshAfterContextMenu = true;
            return;
        }
        if (options.onRefreshResult) {
            currentResult = await options.onRefreshResult();
        }
        if (contextMenuOpen) {
            pendingRefreshAfterContextMenu = true;
            return;
        }
        if (!disposed) {
            renderPanel();
        }
    };

    const scheduleRefresh = (): void => {
        if (disposed) { return; }
        if (contextMenuOpen) {
            pendingRefreshAfterContextMenu = true;
            return;
        }
        if (refreshTimer) { clearTimeout(refreshTimer); }
        refreshTimer = setTimeout(() => {
            refreshTimer = undefined;
            void refreshPanel().catch((err) => {
                const error = err instanceof Error ? err.message : String(err);
                void panel.webview.postMessage({ type: 'error', error });
            });
        }, 350);
    };

    const problemAt = (index: unknown): WorkbookAnalysisProblem | undefined =>
        typeof index === 'number' ? currentResult.problems[index] : undefined;

    renderPanel();
    const messageSub = panel.webview.onDidReceiveMessage(async (message: {
        type?: string;
        index?: number;
        text?: string;
        extension?: string;
        scope?: WorkbookAnalysisSuppressScope;
        severities?: string[];
        fixIndex?: number;
    }) => {
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
                const problem = problemAt(message.index);
                if (problem && options.onOpenProblem) {
                    await options.onOpenProblem(problem, panel.viewColumn);
                }
                return;
            }
            if (message.type === 'suppressProblem') {
                const problem = problemAt(message.index);
                if (problem && message.scope && options.onSuppressProblem) {
                    await options.onSuppressProblem(problem, message.scope, panel.viewColumn);
                    await panel.webview.postMessage({ type: 'suppressed', scope: message.scope });
                    scheduleRefresh();
                }
                return;
            }
            if (message.type === 'askCopilot') {
                const problem = problemAt(message.index);
                if (problem && options.onAskCopilot) {
                    await options.onAskCopilot(problem, panel.viewColumn);
                }
                return;
            }
            if (message.type === 'quickFixProblem') {
                const problem = problemAt(message.index);
                const applied = problem && options.onQuickFixProblem
                    ? await options.onQuickFixProblem(problem, panel.viewColumn, message.fixIndex)
                    : false;
                await panel.webview.postMessage({
                    type: applied ? 'quickFixed' : 'quickFixUnavailable',
                });
                if (applied) {
                    scheduleRefresh();
                }
                return;
            }
            if (message.type === 'untrackProblem') {
                const problem = problemAt(message.index);
                if (problem?.code) {
                    const next = await addUntrackedAnalysisRuleToConfig(problem.code);
                    await panel.webview.postMessage({
                        type: 'untrackedRuleAdded',
                        code: problem.code,
                        untrackedRules: next,
                    });
                    scheduleRefresh();
                }
                return;
            }
            if (message.type === 'updateSeveritySettings') {
                const next = normalizeAnalysisVisibleSeverities(message.severities);
                ignoreNextVisibleSeverityConfigChange = true;
                await vscode.workspace
                    .getConfiguration('xlide')
                    .update('analysis.visibleSeverities', next, vscode.ConfigurationTarget.Global);
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
        if (e.affectsConfiguration('xlide.analysis.visibleSeverities') && ignoreNextVisibleSeverityConfigChange) {
            ignoreNextVisibleSeverityConfigChange = false;
            return;
        }
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
    const treeSub = options.onDidChangeWorkbookTree?.(() => scheduleRefresh());
    const panelDisposables = [
        messageSub,
        configSub,
        textChangeSub,
        saveSub,
        ...(treeSub ? [treeSub] : []),
    ];
    panel.onDidDispose(() => {
        disposed = true;
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

function renderWorkbookAnalysisResultsHtml(
    webview: vscode.Webview,
    model: WorkbookAnalysisResultsModel,
    visibleSeverities: readonly AnalysisSeverityFilter[],
    untrackedRules: readonly string[],
): string {
    const nonce = randomNonce();
    const modelJson = JSON.stringify({
        ...model,
        plainText: buildWorkbookAnalysisPlainText(model),
        visibleSeverities,
        untrackedRules,
    }).replace(/</g, '\\u003c');
    const moduleOrder = new Map(model.groups.map((group, index) => [group.moduleName.toLowerCase(), index]));
    const rowsHtml = model.rows.length === 0
        ? '<div class="empty">No unsuppressed analysis problems.</div>'
        : model.rows.map((row) => `
            <button
                class="problemRow severity-${escapeAttr(row.severity)}"
                type="button"
                data-open-index="${row.index}"
                data-module="${escapeAttr(row.moduleName)}"
                data-module-order="${moduleOrder.get(row.moduleName.toLowerCase()) ?? 9999}"
                data-severity="${escapeAttr(row.severity)}"
                data-compile="${row.vbeCompileEquivalent ? 'yes' : 'no'}"
                data-line="${row.line}"
                data-column="${row.column}"
                data-rule="${escapeAttr(row.code || row.ruleTitle)}"
                data-rule-code="${escapeAttr(row.code)}"
                data-message="${escapeAttr(row.message)}"
                data-evidence="${escapeAttr(row.diagnosticKind)}"
                data-quick-fixes="${escapeAttr(JSON.stringify(row.quickFixTitles))}"
                data-tracked="${isAnalysisRuleTracked(row.code, untrackedRules) ? 'yes' : 'no'}"
            >
                <span class="cell severity">${escapeHtml(row.severity)}</span>
                <span class="cell location">${escapeHtml(row.location)}</span>
                <span class="cell code">${escapeHtml(row.code || row.ruleTitle)}</span>
                <span class="cell message">${escapeHtml(row.message)}</span>
                <span class="cell kind">${escapeHtml(row.diagnosticKind)}</span>
            </button>
        `).join('');
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
        .moduleFilter {
            border: 1px solid var(--vscode-button-border, var(--vscode-panel-border));
            color: var(--vscode-button-secondaryForeground);
            background: var(--vscode-button-secondaryBackground);
            border-radius: 4px;
            padding: 5px 10px;
            cursor: pointer;
        }
        .actionButton:hover,
        .filterButton:hover,
        .moduleFilter:hover {
            background: var(--vscode-button-secondaryHoverBackground);
        }
        .filterButton.active,
        .moduleFilter.active {
            color: var(--vscode-button-foreground);
            background: var(--vscode-button-background);
        }
        .filterButton.transient {
            margin-left: 6px;
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(6, minmax(96px, 1fr));
            gap: 1px;
            background: var(--vscode-panel-border);
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        .stat {
            min-width: 0;
            padding: 10px 14px;
            background: var(--vscode-editor-background);
        }
        .stat strong {
            display: block;
            font-size: 18px;
            font-weight: 600;
            line-height: 1.2;
        }
        .stat span {
            display: block;
            color: var(--vscode-descriptionForeground);
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
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
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
            grid-template-columns: 124px 150px minmax(132px, 190px) minmax(260px, 1fr) 132px;
            align-items: stretch;
            min-width: 936px;
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
        .location {
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
        @media (max-width: 860px) {
            .stats {
                grid-template-columns: repeat(2, minmax(0, 1fr));
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
                <button class="actionButton" id="copyReport" type="button">Copy Report</button>
                <button class="actionButton" id="copyJson" type="button">Copy JSON</button>
                <button class="actionButton" id="exportReport" type="button">Export Report</button>
                <button class="actionButton" id="exportJson" type="button">Export JSON</button>
            </div>
        </header>
        <section class="stats" aria-label="Analysis summary">
            ${statHtml(String(model.totalProblems), 'Problems')}
            ${statHtml(String(model.errorCount), 'Errors')}
            ${statHtml(String(model.warningCount), 'Warnings')}
            ${statHtml(String(model.moduleCount), 'Modules')}
            ${statHtml(String(model.vbeCompileEquivalentCount), 'VBE-equivalent')}
            ${statHtml(String(model.suppressedCount), 'Suppressed')}
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
                        <button class="filterButton transient" type="button" data-show-hidden aria-pressed="false">Show Untracked Items</button>
                    </div>
                    <div class="visibleCount" id="visibleCount"></div>
                </div>
                <div class="table" role="table" aria-label="Analysis problems">
                    <div class="tableHeader" role="row">
                        <button class="cell sortHeader" type="button" role="columnheader" data-sort="severity">Severity</button>
                        <button class="cell sortHeader" type="button" role="columnheader" data-sort="location">Location</button>
                        <button class="cell sortHeader" type="button" role="columnheader" data-sort="rule">Rule</button>
                        <button class="cell sortHeader" type="button" role="columnheader" data-sort="message">Message</button>
                        <button class="cell sortHeader" type="button" role="columnheader" data-sort="evidence">Evidence</button>
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
        <div class="contextDivider"></div>
        <button type="button" data-context-action="suppressBlock">Ignore Block</button>
        <button type="button" data-context-action="suppressMember">Ignore Sub/Function</button>
        <button type="button" data-context-action="suppressModule">Ignore Module</button>
        <div class="contextDivider"></div>
        <button type="button" data-context-action="untrack">Untrack</button>
    </div>
    <div class="toast" id="toast" hidden></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const model = ${modelJson};
        const severityIds = ['error', 'warning', 'information'];
        const visibleSeverities = new Set(model.visibleSeverities ?? severityIds);
        const untrackedRules = new Set(model.untrackedRules ?? []);
        let activeModule = 'all';
        let showHiddenItems = false;
        let sortKey = 'severity';
        let sortDirection = 'asc';
        const rows = Array.from(document.querySelectorAll('.problemRow'));
        const table = document.querySelector('.table');
        const visibleCount = document.getElementById('visibleCount');
        const toast = document.getElementById('toast');
        const contextMenu = document.getElementById('rowContextMenu');
        const quickFixMenuItem = document.getElementById('quickFixMenuItem');
        const quickFixSubmenu = document.getElementById('quickFixSubmenu');
        let contextRow = null;
        let contextMenuVisible = false;
        let persistFiltersTimer = undefined;

        function showToast(message) {
            toast.textContent = message;
            toast.hidden = false;
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => { toast.hidden = true; }, 1800);
        }

        function rowMatches(row) {
            return rowMatchesModule(row) && rowMatchesVisibleFilters(row);
        }

        function rowMatchesModule(row) {
            return activeModule === 'all' || row.dataset.module === activeModule;
        }

        function rowMatchesVisibleFilters(row) {
            return visibleSeverities.has(row.dataset.severity) && row.dataset.tracked !== 'no';
        }

        function updateRows() {
            let count = 0;
            let shownHidden = 0;
            for (const row of rows) {
                const moduleVisible = rowMatchesModule(row);
                const visible = moduleVisible && rowMatchesVisibleFilters(row);
                const untrackedVisible = moduleVisible &&
                    showHiddenItems &&
                    visibleSeverities.has(row.dataset.severity) &&
                    row.dataset.tracked === 'no';
                row.hidden = !visible && !untrackedVisible;
                row.classList.toggle('hiddenVisible', untrackedVisible);
                if (visible) {
                    count += 1;
                } else if (untrackedVisible) {
                    shownHidden += 1;
                }
            }
            visibleCount.textContent = showHiddenItems
                ? \`\${count} shown, \${shownHidden} untracked visible\`
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
            contextMenu.hidden = false;
            const rect = contextMenu.getBoundingClientRect();
            contextMenu.style.left = Math.min(x, window.innerWidth - rect.width - 8) + 'px';
            contextMenu.style.top = Math.min(y, window.innerHeight - rect.height - 8) + 'px';
            contextMenu.querySelector('button')?.focus();
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

        function quickFixTitlesForRow(row) {
            try {
                const parsed = JSON.parse(row.dataset.quickFixes ?? '[]');
                return Array.isArray(parsed) ? parsed.filter(item => typeof item === 'string') : [];
            } catch {
                return [];
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
                hideContextMenu();
                if (typeof index !== 'number' || Number.isNaN(index)) {
                    return;
                }
                if (action === 'applyQuickFix') {
                    vscode.postMessage({
                        type: 'quickFixProblem',
                        index,
                        fixIndex: Number(contextButton.dataset.fixIndex ?? 0),
                    });
                } else if (action === 'askCopilot') {
                    vscode.postMessage({ type: 'askCopilot', index });
                } else if (action === 'untrack') {
                    vscode.postMessage({ type: 'untrackProblem', index });
                } else if (action === 'suppressBlock') {
                    vscode.postMessage({ type: 'suppressProblem', index, scope: 'block' });
                } else if (action === 'suppressMember') {
                    vscode.postMessage({ type: 'suppressProblem', index, scope: 'member' });
                } else if (action === 'suppressModule') {
                    vscode.postMessage({ type: 'suppressProblem', index, scope: 'module' });
                }
                return;
            }
            hideContextMenu();
            const filterButton = event.target.closest?.('[data-severity-toggle]');
            if (filterButton) {
                const id = filterButton.dataset.severityToggle;
                if (visibleSeverities.has(id)) {
                    visibleSeverities.delete(id);
                } else {
                    visibleSeverities.add(id);
                }
                filterButton.classList.toggle('active', visibleSeverities.has(id));
                filterButton.setAttribute('aria-pressed', visibleSeverities.has(id) ? 'true' : 'false');
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
                sortRows();
                updateRows();
                return;
            }
            const showHiddenButton = event.target.closest?.('[data-show-hidden]');
            if (showHiddenButton) {
                showHiddenItems = !showHiddenItems;
                showHiddenButton.classList.toggle('active', showHiddenItems);
                showHiddenButton.setAttribute('aria-pressed', showHiddenItems ? 'true' : 'false');
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
                vscode.postMessage({
                    type: 'openProblem',
                    index: Number(problemRow.dataset.openIndex),
                });
            }
        });

        document.addEventListener('contextmenu', (event) => {
            const problemRow = event.target.closest?.('[data-open-index]');
            if (!problemRow) {
                hideContextMenu();
                return;
            }
            event.preventDefault();
            showContextMenu(problemRow, event.clientX, event.clientY);
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
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
            } else if (event.data?.type === 'untrackedRuleAdded') {
                const code = String(event.data.code ?? '').toLowerCase();
                if (code) {
                    untrackedRules.add(code);
                    for (const row of rows) {
                        if (String(row.dataset.ruleCode ?? '').toLowerCase() === code) {
                            row.dataset.tracked = 'no';
                        }
                    }
                    updateRows();
                }
                showToast(code ? 'Untracked ' + code : 'Rule untracked');
            } else if (event.data?.type === 'error') {
                showToast(event.data.error || 'XLIDE action failed');
            }
        });

        sortRows();
        updateRows();
    </script>
</body>
</html>`;
}

function statHtml(value: string, label: string): string {
    return `<div class="stat"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>`;
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
