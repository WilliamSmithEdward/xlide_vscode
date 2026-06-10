import * as vscode from 'vscode';
import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
    ViewColumn: { Active: -1 },
    window: {
        createWebviewPanel: vi.fn(),
    },
}));

import { openVbaTestResults, renderVbaTestResultsHtml, setVbaTestResultsRunning } from '../src/vbaTestResultsWebview';
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
        expect(panel.reveal).toHaveBeenCalledWith(vscode.ViewColumn.Active);
        expect(panel.webview.html).toContain('99 ms total');
        panel.disposePanel();
    });

    it('posts running state updates to an open results panel', () => {
        const panel = fakeWebviewPanel();
        vi.mocked(vscode.window.createWebviewPanel).mockReturnValue(panel as unknown as vscode.WebviewPanel);
        const context = { subscriptions: [] } as unknown as vscode.ExtensionContext;
        const report = reportFixture();

        openVbaTestResults(context, report);
        setVbaTestResultsRunning(report.filePath, true);
        setVbaTestResultsRunning(report.filePath, false);

        expect(panel.webview.postMessage).toHaveBeenCalledWith({ type: 'setRunning', running: true });
        expect(panel.webview.postMessage).toHaveBeenCalledWith({ type: 'setRunning', running: false });
        panel.disposePanel();
    });

    it('renders a rerun failed action when the command surface provides one', () => {
        const html = renderVbaTestResultsHtml(reportFixture(), {
            canRerunFailed: true,
        });

        expect(html).toContain('data-action="rerunFailed"');
        expect(html).toContain('Rerun Failed (2)');
        expect(html).toContain('vscode.postMessage({ type: \'rerunFailed\' });');
        expect(html).toContain("event.data?.type === 'setRunning'");
        expect(html).toContain('setRunning(Boolean(event.data.running));');
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
        expect(html).toContain('class="testOutput"');
        expect(html).toContain('Output');
        expect(html).toContain('created invoice');
        expect(html).toContain('class="testNameLink"');
        expect(html).toContain('data-open-test-index="0"');
        expect(html).toContain('cursor: pointer');
        expect(html).toContain('type: \'openTest\'');
        expect(html).toContain('white-space: pre-wrap');
        expect(html).toContain('color: var(--vscode-descriptionForeground)');
        expect(html).toContain('background: color-mix(in srgb, var(--xlide-accent-blue) 14%, transparent)');
        expect(html).toContain('owner:finance');
        expect(html).toContain('smoke');
        expect(html).toContain('<span class="tag">known-bug</span>');
        expect(html).not.toContain('<span class="tag">xfail</span>');
        expect(html).toContain('&lt;boom&gt;');
        expect(html).toContain('@xlide-test');
    });

    it('renders developer-friendly details for raw Excel automation run failures', () => {
        const report = reportFixture();
        report.results[1] = {
            ...report.results[1],
            error: [
                'RUN_FAILED|Exception calling "Run" with "1" argument(s):',
                '"Exception from HRESULT: 0x800A9C68"',
                'HRESULT: 0x80131501',
                'Inner: Exception from HRESULT: 0x800A9C68',
                'Inner HRESULT: 0x800A9C68',
                'At C:\\Users\\William\\AppData\\Local\\Temp\\xlide-vba-test-host-DDT6Nq\\run-vba-tests.ps1:677 char:4587',
                '+ ... tPrefix, $excelId, $macroName) }; $excel.Run($macroRef);',
                '+                       ~~~~~~~~~~~~~~~~~~~~',
            ].join('\n'),
        };

        const html = renderVbaTestResultsHtml(report);

        expect(html).toContain('Excel could not run the test macro. Check for VBA compile errors, macro security prompts, or a missing test procedure. HRESULT: 0x800A9C68.');
        expect(html).not.toContain('run-vba-tests.ps1');
        expect(html).not.toContain('$excel.Run');
        expect(html).not.toContain('0x80131501');
    });

    it('opens the clicked test from the latest rendered report', async () => {
        const panel = fakeWebviewPanel();
        vi.mocked(vscode.window.createWebviewPanel).mockReturnValue(panel as unknown as vscode.WebviewPanel);
        const context = { subscriptions: [] } as unknown as vscode.ExtensionContext;
        const firstOpen = vi.fn(async () => undefined);
        const secondOpen = vi.fn(async () => undefined);
        const firstReport = reportFixture();
        const secondReport = {
            ...reportFixture(),
            results: [{
                ...reportFixture().results[0],
                test: {
                    ...reportFixture().results[0].test,
                    qualifiedName: 'Tests.Latest',
                    procedureName: 'Latest',
                    line: 42,
                },
            }],
        };

        openVbaTestResults(context, firstReport, { onOpenTest: firstOpen });
        openVbaTestResults(context, secondReport, { onOpenTest: secondOpen });
        await panel.emitMessage({ type: 'openTest', index: 0 });

        expect(firstOpen).not.toHaveBeenCalled();
        expect(secondOpen).toHaveBeenCalledWith(expect.objectContaining({
            qualifiedName: 'Tests.Latest',
            line: 42,
        }));
        panel.disposePanel();
    });
});

function fakeWebviewPanel(): vscode.WebviewPanel & {
    disposePanel: () => void;
    emitMessage: (message: unknown) => Promise<void>;
} {
    let disposeHandler: (() => void) | undefined;
    let messageHandler: ((message: unknown) => unknown) | undefined;
    return {
        title: '',
        webview: {
            cspSource: 'vscode-resource:',
            html: '',
            onDidReceiveMessage: vi.fn((handler: (message: unknown) => unknown) => {
                messageHandler = handler;
                return { dispose: vi.fn() };
            }),
            postMessage: vi.fn(async () => true),
        },
        reveal: vi.fn(),
        onDidDispose: vi.fn((handler: () => void) => {
            disposeHandler = handler;
            return { dispose: vi.fn() };
        }),
        disposePanel: () => disposeHandler?.(),
        emitMessage: async (message: unknown) => {
            await messageHandler?.(message);
        },
    } as unknown as vscode.WebviewPanel & {
        disposePanel: () => void;
        emitMessage: (message: unknown) => Promise<void>;
    };
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
            { test: pass, status: 'passed', durationMs: 10, output: ['created invoice'] },
            { test: fail, status: 'failed', durationMs: 20, error: 'Err.Raise <boom>' },
            { test: xfail, status: 'xfail', durationMs: 14, error: 'Known issue' },
            { test: timeout, status: 'timeout', durationMs: 5000, error: 'Timed out' },
        ],
    };
}
