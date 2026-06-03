import * as vscode from 'vscode';
import type { VbaTestRunReport, VbaTestRunSummary } from './vbaTestRunner';
import { summarizeVbaTestRun } from './vbaTestRunner';

export function openVbaTestResults(
    context: vscode.ExtensionContext,
    report: VbaTestRunReport,
): vscode.WebviewPanel {
    const panel = vscode.window.createWebviewPanel(
        'xlideVbaTestResults',
        `XLIDE Tests: ${report.workbookName}`,
        vscode.ViewColumn.Beside,
        {
            enableScripts: false,
            retainContextWhenHidden: true,
        },
    );
    panel.webview.html = renderVbaTestResultsHtml(panel.webview, report);
    context.subscriptions.push(panel);
    return panel;
}

export function renderVbaTestResultsHtml(
    webviewOrReport: vscode.Webview | VbaTestRunReport,
    maybeReport?: VbaTestRunReport,
): string {
    const report = maybeReport ?? webviewOrReport as VbaTestRunReport;
    const webview = maybeReport ? webviewOrReport as vscode.Webview : undefined;
    const summary = summarizeVbaTestRun(report);
    const rows = report.results.map((result) => `
        <tr class="${escapeAttr(result.status)}">
            <td><span class="status">${escapeHtml(statusLabel(result.status))}</span></td>
            <td>
                <div class="testName">${escapeHtml(result.test.qualifiedName)}</div>
                <div class="meta">${escapeHtml(`${result.test.moduleName}:${result.test.line}:${result.test.column}`)}</div>
            </td>
            <td>${escapeHtml(`${result.durationMs} ms`)}</td>
            <td>${result.error ? escapeHtml(result.error) : ''}</td>
        </tr>
    `).join('');

    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${webview?.cspSource ?? "'unsafe-inline'"} 'unsafe-inline';">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLIDE VBA Test Results</title>
    <style>
        :root {
            color-scheme: dark light;
        }
        body {
            margin: 0;
            padding: 24px;
            color: var(--vscode-foreground);
            background: var(--vscode-editor-background);
            font: 13px/1.45 var(--vscode-font-family);
        }
        .shell {
            max-width: 1120px;
            margin: 0 auto;
        }
        .header {
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            gap: 16px;
            padding-bottom: 18px;
            border-bottom: 1px solid var(--vscode-panel-border);
        }
        h1 {
            margin: 0;
            font-size: 22px;
            line-height: 1.2;
        }
        .subtitle,
        .meta,
        .contract,
        .empty {
            color: var(--vscode-descriptionForeground);
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(4, minmax(92px, 1fr));
            gap: 10px;
            margin: 20px 0;
        }
        .stat {
            border: 1px solid var(--vscode-panel-border);
            background: var(--vscode-editorWidget-background, var(--vscode-sideBar-background));
            border-radius: 6px;
            padding: 12px;
        }
        .stat strong {
            display: block;
            font-size: 20px;
            line-height: 1.1;
        }
        .stat span {
            color: var(--vscode-descriptionForeground);
        }
        table {
            width: 100%;
            border-collapse: collapse;
            border: 1px solid var(--vscode-panel-border);
        }
        th,
        td {
            padding: 10px 12px;
            border-bottom: 1px solid var(--vscode-panel-border);
            text-align: left;
            vertical-align: top;
        }
        th {
            color: var(--vscode-descriptionForeground);
            background: var(--vscode-editorWidget-background, var(--vscode-sideBar-background));
            font-weight: 600;
        }
        tr:last-child td {
            border-bottom: 0;
        }
        .testName {
            font-weight: 650;
        }
        .status {
            display: inline-block;
            min-width: 68px;
            border-radius: 999px;
            padding: 2px 8px;
            text-align: center;
            font-weight: 650;
        }
        .passed .status {
            color: var(--vscode-testing-iconPassed, #73c991);
            background: color-mix(in srgb, var(--vscode-testing-iconPassed, #73c991) 16%, transparent);
        }
        .failed .status {
            color: var(--vscode-testing-iconFailed, #f48771);
            background: color-mix(in srgb, var(--vscode-testing-iconFailed, #f48771) 16%, transparent);
        }
        .skipped .status {
            color: var(--vscode-testing-iconSkipped, #cca700);
            background: color-mix(in srgb, var(--vscode-testing-iconSkipped, #cca700) 16%, transparent);
        }
        .contract {
            margin-top: 16px;
        }
        @media (max-width: 720px) {
            body {
                padding: 16px;
            }
            .header {
                display: block;
            }
            .stats {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }
        }
    </style>
</head>
<body>
    <main class="shell">
        <header class="header">
            <div>
                <h1>XLIDE VBA Test Results</h1>
                <div class="subtitle">${escapeHtml(report.workbookName)}</div>
            </div>
            <div class="subtitle">${escapeHtml(`${report.durationMs} ms total`)}</div>
        </header>
        ${renderSummary(summary)}
        ${rows
        ? `<table aria-label="VBA Test Results">
            <thead>
                <tr>
                    <th>Status</th>
                    <th>Test</th>
                    <th>Duration</th>
                    <th>Details</th>
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>`
        : `<div class="empty">No VBA tests were discovered.</div>`}
        <div class="contract">${escapeHtml(report.discovery.contract)}</div>
    </main>
</body>
</html>`;
}

function renderSummary(summary: VbaTestRunSummary): string {
    return `<section class="stats" aria-label="VBA Test Summary">
        ${statHtml(summary.total, 'Tests')}
        ${statHtml(summary.passed, 'Passed')}
        ${statHtml(summary.failed, 'Failed')}
        ${statHtml(summary.skipped, 'Skipped')}
    </section>`;
}

function statHtml(value: number, label: string): string {
    return `<div class="stat"><strong>${escapeHtml(String(value))}</strong><span>${escapeHtml(label)}</span></div>`;
}

function statusLabel(status: string): string {
    switch (status) {
        case 'passed':
            return 'Passed';
        case 'failed':
            return 'Failed';
        case 'skipped':
            return 'Skipped';
        default:
            return status;
    }
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
