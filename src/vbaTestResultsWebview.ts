import * as path from 'path';
import * as vscode from 'vscode';
import type { VbaTestRunReport, VbaTestRunSummary } from './vbaTestRunner';
import { describeVbaTestSelection, summarizeVbaTestRun, vbaTestFailureMessage } from './vbaTestRunner';

export interface VbaTestResultsOptions {
    onRerunFailed?: () => Promise<void>;
}

interface OpenVbaTestResultsPanelEntry {
    panel: vscode.WebviewPanel;
    options: VbaTestResultsOptions;
}

interface VbaTestResultsRenderOptions {
    canRerunFailed?: boolean;
}

interface VbaTestResultsWebviewMessage {
    type?: string;
}

const openVbaTestResultsPanels = new Map<string, OpenVbaTestResultsPanelEntry>();

export function openVbaTestResults(
    context: vscode.ExtensionContext,
    report: VbaTestRunReport,
    options: VbaTestResultsOptions = {},
): vscode.WebviewPanel {
    const panelKey = vbaTestResultsPanelKey(report.filePath);
    const existing = openVbaTestResultsPanels.get(panelKey);
    if (existing) {
        existing.options = options;
        existing.panel.title = `XLIDE Test Results: ${report.workbookName}`;
        existing.panel.webview.html = renderVbaTestResultsHtml(existing.panel.webview, report, {
            canRerunFailed: Boolean(options.onRerunFailed),
        });
        existing.panel.reveal(vscode.ViewColumn.Beside);
        return existing.panel;
    }

    const panel = vscode.window.createWebviewPanel(
        'xlideVbaTestResults',
        `XLIDE Test Results: ${report.workbookName}`,
        vscode.ViewColumn.Beside,
        {
            enableScripts: true,
            retainContextWhenHidden: true,
        },
    );
    const entry: OpenVbaTestResultsPanelEntry = {
        panel,
        options,
    };
    openVbaTestResultsPanels.set(panelKey, entry);
    panel.onDidDispose(() => {
        openVbaTestResultsPanels.delete(panelKey);
    });
    panel.webview.onDidReceiveMessage(async (message: VbaTestResultsWebviewMessage) => {
        if (message.type !== 'rerunFailed') {
            return;
        }
        try {
            if (!entry.options.onRerunFailed) {
                await panel.webview.postMessage({ type: 'error', error: 'XLIDE rerun failed is not available.' });
                return;
            }
            await entry.options.onRerunFailed();
            await panel.webview.postMessage({ type: 'rerunComplete' });
        } catch (err) {
            const error = err instanceof Error ? err.message : String(err);
            await panel.webview.postMessage({ type: 'error', error });
        }
    });
    panel.webview.html = renderVbaTestResultsHtml(panel.webview, report, {
        canRerunFailed: Boolean(options.onRerunFailed),
    });
    context.subscriptions.push(panel);
    return panel;
}

export function renderVbaTestResultsHtml(
    webviewOrReport: vscode.Webview | VbaTestRunReport,
    maybeReport?: VbaTestRunReport,
    options: VbaTestResultsRenderOptions = {},
): string {
    const report = maybeReport ?? webviewOrReport as VbaTestRunReport;
    const webview = maybeReport ? webviewOrReport as vscode.Webview : undefined;
    const nonce = randomNonce();
    const cspSource = webview?.cspSource ?? 'vscode-resource:';
    const summary = summarizeVbaTestRun(report);
    const selectionDescription = describeVbaTestSelection(report.discovery.selection);
    const timing = runTimingSummary(report);
    const rerunFailedCount = report.results.filter((result) => isRerunnableFailureStatus(result.status)).length;
    const canRerunFailed = Boolean(options.canRerunFailed && rerunFailedCount > 0);
    const rows = report.results.map((result) => `
        <tr class="${escapeAttr(result.status)}">
            <td><span class="status">${escapeHtml(statusLabel(result.status))}</span></td>
            <td>
                <div class="testName">${escapeHtml(result.test.qualifiedName)}</div>
                <div class="meta">${escapeHtml(`${result.test.moduleName}:${result.test.line}:${result.test.column}`)}</div>
            </td>
            <td class="tagCell">${testMetadataHtml(result.test.metadata)}</td>
            <td>${escapeHtml(`${result.durationMs} ms`)}</td>
            <td class="detailsCell">${result.error ? escapeHtml(vbaTestFailureMessage(result.error)) : ''}</td>
        </tr>
    `).join('');

    return /* html */`<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';">
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
        .headerRight {
            display: grid;
            gap: 12px;
            justify-items: end;
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
        .runTiming {
            display: grid;
            gap: 4px;
            min-width: 260px;
            text-align: right;
        }
        .runTimingRow {
            display: grid;
            grid-template-columns: 72px 1fr;
            gap: 10px;
            align-items: baseline;
        }
        .runTimingLabel {
            color: var(--vscode-descriptionForeground);
        }
        .runTimingValue {
            font-weight: 650;
        }
        .actions {
            display: flex;
            gap: 8px;
            justify-content: flex-end;
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
        button:disabled {
            cursor: not-allowed;
            opacity: 0.55;
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
        .stats {
            display: grid;
            grid-template-columns: repeat(6, minmax(92px, 1fr));
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
            table-layout: fixed;
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
        th:nth-child(1),
        td:nth-child(1) {
            width: 90px;
        }
        th:nth-child(3),
        td:nth-child(3) {
            width: 210px;
        }
        th:nth-child(4),
        td:nth-child(4) {
            width: 84px;
        }
        .testName,
        .meta {
            overflow-wrap: anywhere;
        }
        .detailsCell {
            white-space: pre-wrap;
            overflow-wrap: anywhere;
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
        .timeout .status,
        .host-error .status {
            color: var(--vscode-errorForeground, #f48771);
            background: color-mix(in srgb, var(--vscode-errorForeground, #f48771) 16%, transparent);
        }
        .skipped .status {
            color: var(--vscode-testing-iconSkipped, #cca700);
            background: color-mix(in srgb, var(--vscode-testing-iconSkipped, #cca700) 16%, transparent);
        }
        .xfail .status {
            color: var(--vscode-testing-iconSkipped, #cca700);
            background: color-mix(in srgb, var(--vscode-testing-iconSkipped, #cca700) 16%, transparent);
        }
        .xpass .status {
            color: var(--vscode-testing-iconQueued, #4fc1ff);
            background: color-mix(in srgb, var(--vscode-testing-iconQueued, #4fc1ff) 16%, transparent);
        }
        .tagSet {
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
            margin: 0;
        }
        .tag {
            border: 1px solid color-mix(in srgb, var(--vscode-badge-foreground, #ffffff) 28%, transparent);
            border-radius: 999px;
            padding: 1px 7px;
            color: var(--vscode-badge-foreground, #ffffff);
            background: var(--vscode-badge-background, var(--vscode-button-background));
            font-weight: 600;
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
            .runTiming {
                margin-top: 12px;
                text-align: left;
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
                <div class="subtitle">${escapeHtml(
                    selectionDescription
                        ? `${report.workbookName} - ${selectionDescription}`
                        : report.workbookName,
                )}</div>
            </div>
            <div class="headerRight">
                <div class="runTiming" aria-label="Run timing">
                    <div class="runTimingRow">
                        <span class="runTimingLabel">Started</span>
                        <span class="runTimingValue" title="${escapeAttr(timing.startedIso)}">${escapeHtml(timing.startedLabel)}</span>
                    </div>
                    <div class="runTimingRow">
                        <span class="runTimingLabel">Stopped</span>
                        <span class="runTimingValue" title="${escapeAttr(timing.stoppedIso)}">${escapeHtml(timing.stoppedLabel)}</span>
                    </div>
                    <div class="runTimingRow">
                        <span class="runTimingLabel">Elapsed</span>
                        <span class="runTimingValue">${escapeHtml(`${report.durationMs} ms total`)}</span>
                    </div>
                </div>
                ${canRerunFailed ? `<div class="actions">
                    <button type="button" data-action="rerunFailed" title="${escapeAttr(`Rerun ${rerunFailedCount} failed, timed out, host-error, or unexpected-pass test${rerunFailedCount === 1 ? '' : 's'} from this result.`)}">Rerun Failed (${rerunFailedCount})</button>
                </div>` : ''}
            </div>
        </header>
        ${renderSummary(summary)}
        ${rows
        ? `<table aria-label="VBA Test Results">
            <thead>
                <tr>
                    <th>Status</th>
                    <th>Test</th>
                    <th>Tags</th>
                    <th>Duration</th>
                    <th>Details</th>
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>`
        : `<div class="empty">${escapeHtml(emptyResultsMessage(report, selectionDescription))}</div>`}
        <div class="contract">${escapeHtml(report.discovery.contract)}</div>
    </main>
    <div class="toast" id="toast"></div>
    <script nonce="${nonce}">
        const vscode = acquireVsCodeApi();
        const toast = document.getElementById('toast');
        let toastTimer;
        let running = false;

        function showToast(message) {
            toast.textContent = message;
            toast.classList.add('visible');
            clearTimeout(toastTimer);
            toastTimer = setTimeout(() => toast.classList.remove('visible'), 2600);
        }

        function setRunning(next) {
            running = next;
            document.querySelectorAll('button[data-action="rerunFailed"]').forEach((button) => {
                button.disabled = running;
            });
        }

        document.addEventListener('click', (event) => {
            const button = event.target.closest?.('button[data-action="rerunFailed"]');
            if (!button || button.disabled) {
                return;
            }
            setRunning(true);
            vscode.postMessage({ type: 'rerunFailed' });
        });

        window.addEventListener('message', (event) => {
            if (event.data?.type === 'error') {
                setRunning(false);
                showToast(event.data.error || 'XLIDE test action failed');
            } else if (event.data?.type === 'rerunComplete') {
                setRunning(false);
            }
        });
    </script>
</body>
</html>`;
}

function emptyResultsMessage(report: VbaTestRunReport, selectionDescription: string): string {
    if (selectionDescription && report.discovery.unfilteredTestCount > 0) {
        return 'No VBA tests matched the selected filters.';
    }
    return 'No VBA tests were discovered.';
}

function renderSummary(summary: VbaTestRunSummary): string {
    return `<section class="stats" aria-label="VBA Test Summary">
        ${statHtml(summary.total, 'Tests')}
        ${statHtml(summary.passed, 'Passed')}
        ${statHtml(summary.failed, 'Failed')}
        ${statHtml(summary.timeout, 'Timeout')}
        ${statHtml(summary.hostError, 'Host Errors')}
        ${statHtml(summary.skipped, 'Skipped')}
        ${statHtml(summary.xfail, 'XFail')}
        ${statHtml(summary.xpass, 'XPass')}
    </section>`;
}

function statHtml(value: number, label: string): string {
    return `<div class="stat"><strong>${escapeHtml(String(value))}</strong><span>${escapeHtml(label)}</span></div>`;
}

function runTimingSummary(report: VbaTestRunReport): {
    startedLabel: string;
    startedIso: string;
    stoppedLabel: string;
    stoppedIso: string;
} {
    const startedAt = new Date(report.startedAt);
    if (!Number.isFinite(startedAt.getTime())) {
        return {
            startedLabel: 'Unknown',
            startedIso: '',
            stoppedLabel: 'Unknown',
            stoppedIso: '',
        };
    }
    const stoppedAt = new Date(startedAt.getTime() + Math.max(0, report.durationMs));
    return {
        startedLabel: formatRunTimestamp(startedAt),
        startedIso: startedAt.toISOString(),
        stoppedLabel: formatRunTimestamp(stoppedAt),
        stoppedIso: stoppedAt.toISOString(),
    };
}

function formatRunTimestamp(value: Date): string {
    return value.toLocaleString(undefined, {
        year: 'numeric',
        month: 'short',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
    });
}

function statusLabel(status: string): string {
    switch (status) {
        case 'passed':
            return 'Passed';
        case 'failed':
            return 'Failed';
        case 'timeout':
            return 'Timeout';
        case 'host-error':
            return 'Host Error';
        case 'skipped':
            return 'Skipped';
        case 'xfail':
            return 'XFail';
        case 'xpass':
            return 'XPass';
        default:
            return status;
    }
}

function isRerunnableFailureStatus(status: string): boolean {
    return status === 'failed' ||
        status === 'timeout' ||
        status === 'host-error' ||
        status === 'xpass';
}

function testMetadataHtml(metadata: {
    tags: readonly string[];
    owner?: string;
    requirement?: string;
    timeoutMs?: number;
    expectedError?: string;
    skipReason?: string;
    xfailReason?: string;
}): string {
    const tags = [
        ...metadata.tags,
        metadata.owner ? `owner:${metadata.owner}` : '',
        metadata.requirement ? `req:${metadata.requirement}` : '',
        metadata.timeoutMs ? `timeout:${metadata.timeoutMs}ms` : '',
        metadata.expectedError ? `expected-error:${metadata.expectedError}` : '',
        metadata.xfailReason ? 'xfail' : '',
    ].filter(Boolean);
    if (tags.length === 0) {
        return '';
    }
    return `<div class="tagSet">${tags.map((tag) => `<span class="tag">${escapeHtml(tag)}</span>`).join('')}</div>`;
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

function vbaTestResultsPanelKey(filePath: string): string {
    const normalized = path.normalize(filePath);
    return process.platform === 'win32' ? normalized.toLowerCase() : normalized;
}
