import { describe, expect, it, vi } from 'vitest';
import type { WorkbookAnalysisResult } from '../src/vbaWorkbookAnalysis';
import { buildWorkbookAnalysisResultsModel } from '../src/workbookAnalysisResultsModel';
import { renderWorkbookAnalysisResultsHtml } from '../src/workbookAnalysisWebview';

vi.mock('vscode', () => ({
    workspace: {
        getConfiguration: () => ({
            get: (_key: string, fallback: unknown) => fallback,
            inspect: () => ({}),
        }),
    },
}));

function resultFixture(): WorkbookAnalysisResult {
    return {
        filePath: 'C:/work/Book.xlsm',
        moduleCount: 1,
        errorCount: 1,
        warningCount: 0,
        summary: {
            byCategory: { semantic: 1 },
            byDiagnosticKind: { compile: 1 },
            vbeCompileEquivalentCount: 1,
            nonVbeCompileEquivalentCount: 0,
            suppressedCount: 0,
        },
        problems: [{
            moduleName: 'Module1',
            moduleType: 'standard',
            line: 4,
            column: 2,
            endColumn: 8,
            severity: 'error',
            code: 'undeclared-variable',
            ruleTitle: 'Undeclared variable',
            category: 'semantic',
            vbeCompileEquivalent: true,
            diagnosticKind: 'compile',
            message: 'Variable is not declared.',
        }],
        suppressedProblems: [],
    };
}

describe('workbook analysis webview', () => {
    it('renders scope-explicit rule tracking controls', () => {
        const html = renderWorkbookAnalysisResultsHtml(
            { cspSource: 'test-csp' } as never,
            buildWorkbookAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: [],
                untrackedRulesSource: 'default',
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('data-context-action="setRuleTrackingWorkbook"');
        expect(html).toContain('data-context-action="setRuleTrackingGlobal"');
        expect(html).toContain('Untrack In Workbook');
        expect(html).toContain('Untrack Globally');
        expect(html).toContain('Workbook Rule Tracking');
        expect(html).toContain('trackingScope: \'workbook\'');
        expect(html).toContain('trackingScope: action === \'setRuleTrackingGlobal\' ? \'global\' : \'workbook\'');
        expect(html).not.toContain('id="trackingAction"');
    });
});
