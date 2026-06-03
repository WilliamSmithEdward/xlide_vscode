import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
    ViewColumn: { Beside: 2 },
    window: {
        createWebviewPanel: vi.fn(),
    },
}));

import { renderVbaTestResultsHtml } from '../src/vbaTestResultsWebview';
import type { VbaTestRunReport } from '../src/vbaTestRunner';

describe('VBA test results webview', () => {
    it('renders summary stats, discovered tests, and escaped failure details', () => {
        const html = renderVbaTestResultsHtml(reportFixture());

        expect(html).toContain('XLIDE VBA Test Results');
        expect(html).toContain('Book.xlsm');
        expect(html).toContain('Tests.TestPass');
        expect(html).toContain('Tests.TestFail');
        expect(html).toContain('Passed');
        expect(html).toContain('Failed');
        expect(html).toContain('Timeout');
        expect(html).toContain('Host Errors');
        expect(html).toContain('XFail');
        expect(html).toContain('owner:finance');
        expect(html).toContain('smoke');
        expect(html).toContain('&lt;boom&gt;');
        expect(html).toContain('@xlide-test');
    });
});

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
