import * as vscode from 'vscode';
import type { VbaTestCase, VbaTestRunItem, VbaTestRunReport, VbaTestRunSummary } from './vbaTestRunner';
import { describeVbaTestSelection, summarizeVbaTestRun } from './vbaTestRunner';
import { escapeAttr, escapeHtml, randomNonce } from './webview/html';
import {
    statHtml,
    webviewHeadHtml,
    WEBVIEW_TOAST_CSS,
    WEBVIEW_TOAST_HTML,
    WEBVIEW_TOAST_SCRIPT,
} from './webview/page';
import { bridgeWebviewMessages, createWebviewPanelRegistry } from './webview/panelRegistry';
import { WEBVIEW_BODY_CSS, WEBVIEW_PRIMARY_BUTTON_CSS, xlideAccentPaletteCss } from './webview/styles';
import { renderWebviewTemplate } from './webview/templates';

export interface VbaTestResultsOptions {
    onRerunFailed?: () => Promise<void>;
    onOpenTest?: (test: VbaTestCase) => Promise<void>;
}

interface OpenVbaTestResultsPanelEntry {
    panel: vscode.WebviewPanel;
    options: VbaTestResultsOptions;
    report: VbaTestRunReport;
}

interface VbaTestResultsRenderOptions {
    canRerunFailed?: boolean;
}

interface VbaTestResultsWebviewMessage {
    type?: string;
    index?: number;
}

const openVbaTestResultsPanels = createWebviewPanelRegistry<OpenVbaTestResultsPanelEntry>();

export function openVbaTestResults(
    context: vscode.ExtensionContext,
    report: VbaTestRunReport,
    options: VbaTestResultsOptions = {},
): vscode.WebviewPanel {
    const existing = openVbaTestResultsPanels.get(report.filePath);
    if (existing) {
        existing.options = options;
        existing.report = report;
        existing.panel.title = `XLIDE Test Results: ${report.workbookName}`;
        existing.panel.webview.html = renderVbaTestResultsHtml(report, {
            canRerunFailed: Boolean(options.onRerunFailed),
        });
        existing.panel.reveal(vscode.ViewColumn.Active);
        return existing.panel;
    }

    const panel = vscode.window.createWebviewPanel(
        'xlideVbaTestResults',
        `XLIDE Test Results: ${report.workbookName}`,
        vscode.ViewColumn.Active,
        {
            enableScripts: true,
            // Required: setVbaTestResultsRunning posts running-state updates
            // while the panel may be hidden; a dead iframe would drop them.
            retainContextWhenHidden: true,
        },
    );
    const entry: OpenVbaTestResultsPanelEntry = {
        panel,
        options,
        report,
    };
    openVbaTestResultsPanels.set(report.filePath, entry);
    const messageSub = bridgeWebviewMessages(panel.webview, async (message: VbaTestResultsWebviewMessage) => {
        if (message.type === 'openTest') {
            const test = typeof message.index === 'number'
                ? entry.report.results[message.index]?.test
                : undefined;
            if (!test || !entry.options.onOpenTest) {
                await panel.webview.postMessage({ type: 'error', error: 'XLIDE test navigation is not available.' });
                return;
            }
            await entry.options.onOpenTest(test);
            return;
        }
        if (message.type !== 'rerunFailed') {
            return;
        }
        if (!entry.options.onRerunFailed) {
            await panel.webview.postMessage({ type: 'error', error: 'XLIDE rerun failed is not available.' });
            return;
        }
        await entry.options.onRerunFailed();
        await panel.webview.postMessage({ type: 'rerunComplete' });
    });
    panel.onDidDispose(() => {
        messageSub.dispose();
        openVbaTestResultsPanels.delete(report.filePath);
    });
    panel.webview.html = renderVbaTestResultsHtml(report, {
        canRerunFailed: Boolean(options.onRerunFailed),
    });
    context.subscriptions.push(panel);
    return panel;
}

export function setVbaTestResultsRunning(filePath: string, running: boolean): void {
    const entry = openVbaTestResultsPanels.get(filePath);
    if (!entry) {
        return;
    }
    void entry.panel.webview.postMessage({ type: 'setRunning', running });
}

export function renderVbaTestResultsHtml(
    report: VbaTestRunReport,
    options: VbaTestResultsRenderOptions = {},
): string {
    const nonce = randomNonce();
    const summary = summarizeVbaTestRun(report);
    const selectionDescription = describeVbaTestSelection(report.discovery.selection);
    const timing = runTimingSummary(report);
    const rerunFailedCount = report.results.filter((result) => isRerunnableFailureStatus(result.status)).length;
    const canRerunFailed = Boolean(options.canRerunFailed && rerunFailedCount > 0);
    const rows = report.results.map((result, index) => `
        <tr class="${escapeAttr(result.status)}">
            <td><span class="status">${escapeHtml(statusLabel(result.status))}</span></td>
            <td>
                <button
                    class="testNameLink"
                    type="button"
                    data-open-test-index="${index}"
                    title="Open ${escapeAttr(result.test.qualifiedName)}"
                >${escapeHtml(result.test.qualifiedName)}</button>
                <div class="meta">${escapeHtml(`${result.test.moduleName}:${result.test.line}:${result.test.column}`)}</div>
            </td>
            <td class="tagCell">${testMetadataHtml(result.test.metadata)}</td>
            <td>${escapeHtml(`${result.durationMs} ms`)}</td>
            <td class="detailsCell">${resultDetailsHtml(result)}</td>
        </tr>
    `).join('');

    return renderWebviewTemplate('assets/webview/vbaTestResults.html', {
        head: webviewHeadHtml(nonce, 'XLIDE VBA Test Results'),
        nonce,
        css: renderWebviewTemplate('assets/webview/vbaTestResults.css', {
            accentPalette: xlideAccentPaletteCss(),
            bodyCss: WEBVIEW_BODY_CSS,
            primaryButtonCss: WEBVIEW_PRIMARY_BUTTON_CSS,
            toastCss: WEBVIEW_TOAST_CSS,
        }),
        subtitle: escapeHtml(
            selectionDescription
                ? `${report.workbookName} - ${selectionDescription}`
                : report.workbookName,
        ),
        startedIso: escapeAttr(timing.startedIso),
        startedLabel: escapeHtml(timing.startedLabel),
        stoppedIso: escapeAttr(timing.stoppedIso),
        stoppedLabel: escapeHtml(timing.stoppedLabel),
        elapsed: escapeHtml(`${report.durationMs} ms total`),
        rerunActions: canRerunFailed ? `<div class="actions">
                    <button type="button" data-action="rerunFailed" title="${escapeAttr(`Rerun ${rerunFailedCount} failed, timed out, host-error, or unexpected-pass test${rerunFailedCount === 1 ? '' : 's'} from this result.`)}">Rerun Failed (${rerunFailedCount})</button>
                </div>` : '',
        summary: renderSummary(summary),
        results: rows
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
            : `<div class="empty">${escapeHtml(emptyResultsMessage(report, selectionDescription))}</div>`,
        contract: escapeHtml(report.discovery.contract),
        toastHtml: WEBVIEW_TOAST_HTML,
        js: renderWebviewTemplate('assets/webview/vbaTestResults.js', {
            toastScript: WEBVIEW_TOAST_SCRIPT,
        }),
    });
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
    const startedLabel = formatRunTimestamp(startedAt);
    // durationMs is normally finite, but a single non-finite value upstream would
    // make stoppedAt an Invalid Date and toISOString() throw, taking down the whole
    // panel render. Degrade the stopped fields to the started value instead.
    const stoppedFinite = Number.isFinite(stoppedAt.getTime());
    return {
        startedLabel,
        startedIso: startedAt.toISOString(),
        stoppedLabel: stoppedFinite ? formatRunTimestamp(stoppedAt) : startedLabel,
        stoppedIso: stoppedFinite ? stoppedAt.toISOString() : '',
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
    expectedError?: number | 'any';
    skipReason?: string;
    xfailReason?: string;
}): string {
    const tags = [
        ...metadata.tags,
        metadata.owner ? `owner:${metadata.owner}` : '',
        metadata.requirement ? `req:${metadata.requirement}` : '',
        metadata.timeoutMs ? `timeout:${metadata.timeoutMs}ms` : '',
        metadata.expectedError ? `expected-error:${metadata.expectedError}` : '',
    ].filter(Boolean);
    if (tags.length === 0) {
        return '';
    }
    return `<div class="tagSet">${tags.map((tag) => `<span class="tag">${escapeHtml(tag)}</span>`).join('')}</div>`;
}

// result.error is already cleaned by the execution layer (vbaTestFailureMessages.ts).
function resultDetailsHtml(result: VbaTestRunItem): string {
    const details = result.error ? `<div>${escapeHtml(result.error)}</div>` : '';
    const output = result.output?.length
        ? `<div class="testOutput"><div class="outputLabel">Output</div>${result.output
            .map((line) => `<div class="outputLine">${escapeHtml(line)}</div>`)
            .join('')}</div>`
        : '';
    return `${details}${output}`;
}

