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

function suppressedOnlyResultFixture(): WorkbookAnalysisResult {
    const fixture = resultFixture();
    fixture.errorCount = 0;
    fixture.warningCount = 0;
    fixture.summary.suppressedCount = 1;
    fixture.problems = [];
    fixture.suppressedProblems = [{
        moduleName: 'UserForm1',
        moduleType: 'userform',
        line: 1,
        column: 1,
        endColumn: 52,
        severity: 'warning',
        code: 'option-explicit-missing',
        ruleTitle: 'Option Explicit missing',
        category: 'style',
        vbeCompileEquivalent: false,
        diagnosticKind: 'style-policy',
        message: 'Option Explicit is not specified.',
        suppressed: true,
    }];
    return fixture;
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
        expect(html).toContain('Track In Workbook');
        expect(html).toContain('Track Globally');
        expect(html).toContain('Workbook Rule Tracking');
        expect(html).toContain('data-tracking-source="tracked"');
        expect(html).toContain('id="trackingDivider"');
        expect(html).toContain('syncTrackingActions(row)');
        expect(html).toContain('grid-template-columns: 124px minmax(184px, 210px)');
        expect(html).toContain('trackingScope: \'workbook\'');
        expect(html).toContain('trackingScope: action === \'setRuleTrackingGlobal\' ? \'global\' : \'workbook\'');
        expect(html).not.toContain('id="trackingAction"');
    });

    it('labels globally untracked rule rows distinctly', () => {
        const html = renderWorkbookAnalysisResultsHtml(
            { cspSource: 'test-csp' } as never,
            buildWorkbookAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: ['undeclared-variable'],
                untrackedRulesSource: 'machine',
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('Untracked Globally');
        expect(html).toContain('data-tracked="no"');
        expect(html).toContain('data-tracking-source="global"');
    });

    it('labels workbook-untracked rule rows distinctly', () => {
        const html = renderWorkbookAnalysisResultsHtml(
            { cspSource: 'test-csp' } as never,
            buildWorkbookAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: ['undeclared-variable'],
                untrackedRulesSource: 'workbook',
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('Untracked In Workbook');
        expect(html).toContain('data-tracked="no"');
        expect(html).toContain('data-tracking-source="workbook"');
    });

    it('gives untracked status precedence over suppressed findings', () => {
        const html = renderWorkbookAnalysisResultsHtml(
            { cspSource: 'test-csp' } as never,
            buildWorkbookAnalysisResultsModel(suppressedOnlyResultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: ['option-explicit-missing'],
                untrackedRulesSource: 'machine',
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('<strong>0</strong><span>Suppressed</span>');
        expect(html).toContain('<strong>1</strong><span>Untracked</span>');
        expect(html).toContain('data-suppressed="yes"');
        expect(html).toContain('data-status="untracked"');
        expect(html).toContain('<span class="cell status">Untracked Globally</span>');
        expect(html).not.toContain('<span class="cell status">Suppressed</span>');
    });
});
