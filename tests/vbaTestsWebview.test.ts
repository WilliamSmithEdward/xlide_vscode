import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
    ViewColumn: { Beside: 2 },
    window: {
        createWebviewPanel: vi.fn(),
    },
}));

import { renderVbaTestsHtml, type VbaTestsPanelModel } from '../src/vbaTestsWebview';

describe('VBA tests webview', () => {
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
        expect(html).toContain('data-action="runWithFilters" disabled');
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
        expect(html).toContain('data-action="runWithFilters" >Run With Filters</button>');
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
        expect(html).toContain('data-action="runWithFilters" disabled');
        expect(html).toContain('Install Microsoft Excel before running workbook tests through XLIDE.');
    });
});

function model(
    support: VbaTestsPanelModel['support'],
    runtime: VbaTestsPanelModel['runtime'] = {
        state: 'installed',
        title: 'Excel COM Ready',
        description: 'Microsoft Excel is registered for COM automation on this machine.',
        canRun: true,
    },
): VbaTestsPanelModel {
    return {
        filePath: 'C:\\work\\Book.xlsm',
        workbookName: 'Book.xlsm',
        support,
        runtime,
    };
}
