import * as vscode from 'vscode';
import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
    ViewColumn: { Beside: 2 },
    window: {
        createWebviewPanel: vi.fn(),
    },
}));

import { openVbaTestResults, renderVbaTestResultsHtml } from '../src/vbaTestResultsWebview';
import type { VbaTestRunReport } from '../src/vbaTestRunner';

describe('VBA test results webview', () => {
    beforeEach(() => {
        vi.clearAllMocks();
    });

    it('reuses an existing results panel for the same workbook', () => {
        const panel = fakeWebviewPanel();
        vi.mocked(vscode.window.createWebviewPanel).mockReturnValue(panel as unknown as vscode.WebviewPanel);
        const context = { subscriptions: [] } as unknown as vscode.ExtensionContext;
        const firstReport = reportFixture();
        const secondReport = { ...reportFixture(), durationMs: 99 };

        const first = openVbaTestResults(context, firstReport);
        const second = openVbaTestResults(context, secondReport);

        expect(first).toBe(second);
        expect(vscode.window.createWebviewPanel).toHaveBeenCalledTimes(1);
        expect(panel.reveal).toHaveBeenCalledWith(vscode.ViewColumn.Beside);
        expect(panel.webview.html).toContain('99 ms total');
        panel.disposePanel();
    });

    it('renders summary stats, discovered tests, and escaped failure details', () => {
        const html = renderVbaTestResultsHtml(reportFixture());

        expect(html).toContain('XLIDE VBA Test Results');
        expect(html).toContain('Book.xlsm');
        expect(html).toContain('Started');
        expect(html).toContain('Stopped');
        expect(html).toContain('Elapsed');
        expect(html).toContain('title="2026-01-01T00:00:00.000Z"');
        expect(html).toContain('title="2026-01-01T00:00:00.045Z"');
        expect(html).toContain('Tests.TestPass');
        expect(html).toContain('Tests.TestFail');
        expect(html).toContain('Passed');
        expect(html).toContain('Failed');
        expect(html).toContain('Timeout');
        expect(html).toContain('Host Errors');
        expect(html).toContain('XFail');
        expect(html).toContain('<th>Tags</th>');
        expect(html).toContain('class="tagCell"');
        expect(html).toContain('class="detailsCell"');
        expect(html).toContain('white-space: pre-wrap');
        expect(html).toContain('var(--vscode-badge-foreground');
        expect(html).toContain('owner:finance');
        expect(html).toContain('smoke');
        expect(html).toContain('&lt;boom&gt;');
        expect(html).toContain('@xlide-test');
    });
});

function fakeWebviewPanel(): vscode.WebviewPanel & { disposePanel: () => void } {
    let disposeHandler: (() => void) | undefined;
    return {
        title: '',
        webview: {
            cspSource: 'vscode-resource:',
            html: '',
        },
        reveal: vi.fn(),
        onDidDispose: vi.fn((handler: () => void) => {
            disposeHandler = handler;
            return { dispose: vi.fn() };
        }),
        disposePanel: () => disposeHandler?.(),
    } as unknown as vscode.WebviewPanel & { disposePanel: () => void };
}

function reportFixture(): VbaTestRunReport {
    const pass = {
        id: 'Tests.TestPass',
        moduleName: 'Tests',
        moduleType: 'standard',
        procedureName: 'TestPass',
        qualifiedName: 'Tests.TestPass',
        line: 2,
        column: 1,
        annotationLine: 1,
        metadata: {
            tags: ['smoke'],
            owner: 'finance',
            requirement: 'INV-104',
        },
    };
    const fail = {
        ...pass,
        id: 'Tests.TestFail',
        procedureName: 'TestFail',
        qualifiedName: 'Tests.TestFail',
        line: 6,
        annotationLine: 5,
        metadata: { tags: [] },
    };
    const xfail = {
        ...pass,
        id: 'Tests.TestExpectedFailure',
        procedureName: 'TestExpectedFailure',
        qualifiedName: 'Tests.TestExpectedFailure',
        line: 10,
        annotationLine: 9,
        metadata: { tags: ['known-bug'], xfailReason: 'Pending fix' },
    };
    const timeout = {
        ...pass,
        id: 'Tests.TestTimeout',
        procedureName: 'TestTimeout',
        qualifiedName: 'Tests.TestTimeout',
        line: 14,
        annotationLine: 13,
        metadata: { tags: ['slow'], timeoutMs: 5000 },
    };
    return {
        filePath: 'C:/work/Book.xlsm',
        workbookName: 'Book.xlsm',
        startedAt: '2026-01-01T00:00:00.000Z',
        durationMs: 45,
        discovery: {
            filePath: 'C:/work/Book.xlsm',
            tests: [pass, fail, xfail, timeout],
            unfilteredTestCount: 4,
            modulesScanned: 1,
            modulesIgnored: 0,
            contract: "Standard-module no-argument Sub procedures with '@xlide-test'.",
        },
        results: [
            { test: pass, status: 'passed', durationMs: 10 },
            { test: fail, status: 'failed', durationMs: 20, error: 'Err.Raise <boom>' },
            { test: xfail, status: 'xfail', durationMs: 14, error: 'Known issue' },
            { test: timeout, status: 'timeout', durationMs: 5000, error: 'Timed out' },
        ],
    };
}
