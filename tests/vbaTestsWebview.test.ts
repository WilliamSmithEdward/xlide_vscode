import * as vscode from 'vscode';
import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
    ViewColumn: { Beside: 2 },
    window: {
        createWebviewPanel: vi.fn(),
    },
}));

import { openVbaTestsPanel, renderVbaTestsHtml, type VbaTestsPanelModel } from '../src/vbaTestsWebview';

describe('VBA tests webview', () => {
    it('reuses an existing tests panel for the same workbook', () => {
        const panel = fakeWebviewPanel();
        vi.mocked(vscode.window.createWebviewPanel).mockReturnValue(panel as unknown as vscode.WebviewPanel);
        const context = { subscriptions: [] } as unknown as vscode.ExtensionContext;
        const options = { getModel: async () => model(installedSupport()) };

        const first = openVbaTestsPanel(context, 'C:\\work\\Book.xlsm', options);
        const second = openVbaTestsPanel(context, 'C:\\work\\Book.xlsm', options);

        expect(first).toBe(second);
        expect(vscode.window.createWebviewPanel).toHaveBeenCalledTimes(1);
        expect(panel.reveal).toHaveBeenCalledWith(vscode.ViewColumn.Beside);
        panel.disposePanel();
    });

    it('blocks run actions until the bundled support module is installed', () => {
        const html = renderVbaTestsHtml(model({
            state: 'missing',
            title: 'XlideAssert.bas Not Installed',
            description: 'The bundled test support module must be installed before XLIDE can run workbook tests.',
            actionLabel: 'Install',
            canInstall: true,
            canRun: false,
        }));

        expect(html).toContain('XLIDE Unit Tests');
        expect(html).toContain('Book.xlsm');
        expect(html).toContain('XlideAssert.bas Not Installed');
        expect(html).toContain('Excel COM Ready');
        expect(html).toContain('data-action="installSupport"');
        expect(html).not.toContain('data-action="refresh"');
        expect(html).toContain('data-action="runAll" disabled');
        expect(html).toMatch(/data-action="runWithFilters"[^>]*disabled/);
        expect(html).toContain('class="statusGrid"');
    });

    it('enables run actions when test support is installed', () => {
        const html = renderVbaTestsHtml(model({
            state: 'installed',
            title: 'XlideAssert.bas Installed',
            description: 'Workbook tests can run through the XLIDE-owned read-only Excel test host.',
            actionLabel: 'Installed',
            canInstall: false,
            canRun: true,
        }));

        expect(html).toContain('XlideAssert.bas Installed');
        expect(html).toContain('data-action="installSupport" title="Workbook tests can run through the XLIDE-owned read-only Excel test host." disabled');
        expect(html).toContain('data-action="runAll" >Run All Tests</button>');
        expect(html).toContain('data-action="runWithFilters" title="Run selected tag filters" >Run With Filters</button>');
        expect(html).toContain('Include Tags');
        expect(html).toContain('Exclude Tags');
        expect(html).toContain('data-filter-action="selectAll" data-filter-kind="include"');
        expect(html).toContain('data-filter-kind="include" data-tag="smoke" checked');
        expect(html).toContain('data-filter-kind="exclude" data-tag="fast"');
        expect(html).toContain('Fail Fast');
        expect(html).toContain('let running = false;');
        expect(html).toContain('setRunning(true);');
    });

    it('disables filtered runs when the workbook has no tag filters', () => {
        const html = renderVbaTestsHtml(model({
            state: 'installed',
            title: 'XlideAssert.bas Installed',
            description: 'Workbook tests can run through the XLIDE-owned read-only Excel test host.',
            actionLabel: 'Installed',
            canInstall: false,
            canRun: true,
        }, undefined, {
            totalTests: 1,
            taggedTests: 0,
            untaggedTests: 1,
            tags: [],
        }));

        expect(html).toContain('data-action="runAll" >Run All Tests</button>');
        expect(html).toContain('data-action="runWithFilters" title="No test tags discovered in this workbook." disabled');
        expect(html).toContain('No test tags discovered');
    });

    it('blocks run actions when Excel COM is unavailable', () => {
        const html = renderVbaTestsHtml(model({
            state: 'installed',
            title: 'XlideAssert.bas Installed',
            description: 'Workbook tests can run through the XLIDE-owned read-only Excel test host.',
            actionLabel: 'Installed',
            canInstall: false,
            canRun: true,
        }, {
            state: 'missing',
            title: 'Excel COM Not Found',
            description: 'Install Microsoft Excel before running workbook tests through XLIDE.',
            canRun: false,
        }));

        expect(html).toContain('XlideAssert.bas Installed');
        expect(html).toContain('Excel COM Not Found');
        expect(html).toContain('data-action="runAll" disabled');
        expect(html).toMatch(/data-action="runWithFilters"[^>]*disabled/);
        expect(html).toContain('Install Microsoft Excel before running workbook tests through XLIDE.');
    });

    it('keeps support install available when Excel COM is unavailable', () => {
        const html = renderVbaTestsHtml(model({
            state: 'missing',
            title: 'XlideAssert.bas Not Installed',
            description: 'The bundled test support module is installed with pyopenvba before XLIDE runs workbook tests.',
            actionLabel: 'Install',
            canInstall: true,
            canRun: false,
        }, {
            state: 'missing',
            title: 'Excel COM Not Found',
            description: 'Install Microsoft Excel before running workbook tests through XLIDE.',
            canRun: false,
        }));

        expect(html).toContain('data-action="installSupport" title="Install Book.xlsm" >Install</button>');
        expect(html).toContain('data-action="runAll" disabled');
        expect(html).toMatch(/data-action="runWithFilters"[^>]*disabled/);
    });
});

function fakeWebviewPanel(): vscode.WebviewPanel & { disposePanel: () => void } {
    let disposeHandler: (() => void) | undefined;
    return {
        title: '',
        webview: {
            cspSource: 'vscode-resource:',
            html: '',
            postMessage: vi.fn(async () => true),
            onDidReceiveMessage: vi.fn(() => ({ dispose: vi.fn() })),
        },
        reveal: vi.fn(),
        onDidDispose: vi.fn((handler: () => void) => {
            disposeHandler = handler;
            return { dispose: vi.fn() };
        }),
        disposePanel: () => disposeHandler?.(),
    } as unknown as vscode.WebviewPanel & { disposePanel: () => void };
}

function installedSupport(): VbaTestsPanelModel['support'] {
    return {
        state: 'installed',
        title: 'XlideAssert.bas Installed',
        description: 'Workbook tests can run through the XLIDE-owned read-only Excel test host.',
        actionLabel: 'Installed',
        canInstall: false,
        canRun: true,
    };
}

function model(
    support: VbaTestsPanelModel['support'],
    runtime: VbaTestsPanelModel['runtime'] = {
        state: 'installed',
        title: 'Excel COM Ready',
        description: 'Microsoft Excel is registered for COM automation on this machine.',
        canRun: true,
    },
    discovery: VbaTestsPanelModel['discovery'] = {
        totalTests: 3,
        taggedTests: 2,
        untaggedTests: 1,
        tags: [
            { name: 'fast', testCount: 1 },
            { name: 'smoke', testCount: 2 },
        ],
    },
): VbaTestsPanelModel {
    return {
        filePath: 'C:\\work\\Book.xlsm',
        workbookName: 'Book.xlsm',
        support,
        runtime,
        discovery,
    };
}
