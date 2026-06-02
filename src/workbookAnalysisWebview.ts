import * as vscode from 'vscode';
import * as path from 'path';
import type { WorkbookAnalysisProblem, WorkbookAnalysisResult } from './vbaWorkbookAnalysis';
import {
    buildWorkbookAnalysisPlainText,
    buildWorkbookAnalysisResultsModel,
    type WorkbookAnalysisResultsModel,
} from './workbookAnalysisResultsModel';

export interface WorkbookAnalysisResultsOptions {
    onOpenProblem?: (problem: WorkbookAnalysisProblem, analysisPanelColumn?: vscode.ViewColumn) => Promise<void>;
}

export function openWorkbookAnalysisResults(
    context: vscode.ExtensionContext,
    result: WorkbookAnalysisResult,
    options: WorkbookAnalysisResultsOptions = {},
): vscode.WebviewPanel {
    const model = buildWorkbookAnalysisResultsModel(result);
    const panel = vscode.window.createWebviewPanel(
        'xlideWorkbookAnalysisResults',
        `XLIDE Analysis: ${model.workbookName}`,
        vscode.ViewColumn.Beside,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
        },
    );

    panel.webview.html = renderWorkbookAnalysisResultsHtml(panel.webview, model);
    const messageSub = panel.webview.onDidReceiveMessage(async (message: {
        type?: string;
        index?: number;
        text?: string;
        extension?: string;
    }) => {
        try {
            if (message.type === 'openProblem') {
                const problem = typeof message.index === 'number'
                    ? result.problems[message.index]
                    : undefined;
                if (problem && options.onOpenProblem) {
                    await options.onOpenProblem(problem, panel.viewColumn);
                }
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
                        path.dirname(model.filePath),
                        `${sanitizeFileName(model.workbookName)}.xlide-analysis.${extension}`,
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

    context.subscriptions.push(panel, messageSub);
    return panel;
}

function renderWorkbookAnalysisResultsHtml(webview: vscode.Webview, model: WorkbookAnalysisResultsModel): string {
    const nonce = randomNonce();
    const modelJson = JSON.stringify({
        ...model,
        plainText: buildWorkbookAnalysisPlainText(model),
    }).replace(/</g, '\\u003c');
    const rowsHtml = model.rows.length === 0
        ? '<div class="empty">No unsuppressed analysis problems.</div>'
        : model.rows.map((row) => `
            <button
                class="problemRow severity-${escapeAttr(row.severity)}"
                type="button"
                data-open-index="${row.index}"
                data-module="${escapeAttr(row.moduleName)}"
                data-severity="${escapeAttr(row.severity)}"
                data-compile="${row.vbeCompileEquivalent ? 'yes' : 'no'}"
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
                    <span class="moduleName">${escapeHtml(group.moduleName)}</span>
                    <span>${group.total}</span>
                </button>
            `),
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
            grid-template-columns: 88px 150px minmax(132px, 190px) minmax(260px, 1fr) 132px;
            align-items: stretch;
            min-width: 900px;
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
                        <button class="filterButton active" type="button" data-filter="all">All</button>
                        <button class="filterButton" type="button" data-filter="error">Errors</button>
                        <button class="filterButton" type="button" data-filter="warning">Warnings</button>
                        <button class="filterButton" type="button" data-filter="compile">VBE</button>
                        <button class="filterButton" type="button" data-filter="guidance">Guidance</button>
                    </div>
                    <div class="visibleCount" id="visibleCount"></div>
                </div>
                <div class="table" role="table" aria-label="Analysis problems">
                    <div class="tableHeader" role="row">
                        <span class="cell" role="columnheader">Severity</span>
                        <span class="cell" role="columnheader">Location</span>
                        <span class="cell" role="columnheader">Rule</span>
                        <span class="cell" role="columnheader">Message</span>
                        <span class="cell" role="columnheader">Evidence</span>
                    </div>
                    ${rowsHtml}
                </div>
            </main>
        </div>
    </div>
    <div class="toast" id="toast" hidden></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const model = ${modelJson};
        let activeFilter = 'all';
        let activeModule = 'all';
        const rows = Array.from(document.querySelectorAll('.problemRow'));
        const visibleCount = document.getElementById('visibleCount');
        const toast = document.getElementById('toast');

        function showToast(message) {
            toast.textContent = message;
            toast.hidden = false;
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => { toast.hidden = true; }, 1800);
        }

        function rowMatches(row) {
            if (activeModule !== 'all' && row.dataset.module !== activeModule) {
                return false;
            }
            if (activeFilter === 'all') {
                return true;
            }
            if (activeFilter === 'compile') {
                return row.dataset.compile === 'yes';
            }
            if (activeFilter === 'guidance') {
                return row.dataset.compile !== 'yes';
            }
            return row.dataset.severity === activeFilter;
        }

        function updateRows() {
            let count = 0;
            for (const row of rows) {
                const visible = rowMatches(row);
                row.hidden = !visible;
                if (visible) {
                    count += 1;
                }
            }
            visibleCount.textContent = \`\${count} shown\`;
        }

        function setActive(buttons, activeButton) {
            for (const button of buttons) {
                button.classList.toggle('active', button === activeButton);
            }
        }

        document.addEventListener('click', (event) => {
            const filterButton = event.target.closest?.('[data-filter]');
            if (filterButton) {
                activeFilter = filterButton.dataset.filter;
                setActive(document.querySelectorAll('[data-filter]'), filterButton);
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
            } else if (event.data?.type === 'error') {
                showToast(event.data.error || 'XLIDE action failed');
            }
        });

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
