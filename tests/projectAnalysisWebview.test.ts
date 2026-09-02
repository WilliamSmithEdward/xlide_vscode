import * as vscode from 'vscode';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { ProjectAnalysisResult } from '../src/vbaProjectWideAnalysis';
import { buildProjectAnalysisResultsModel } from '../src/projectAnalysisResultsModel';
import { openProjectAnalysisResults, renderProjectAnalysisResultsHtml } from '../src/projectAnalysisWebview';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

function resultFixture(): ProjectAnalysisResult {
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

function suppressedOnlyResultFixture(): ProjectAnalysisResult {
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

describe('project analysis webview', () => {
    beforeEach(() => {
        vi.clearAllMocks();
    });

    it('reuses an existing analysis panel for the project in the active editor group', () => {
        const panel = fakeWebviewPanel();
        vi.mocked(vscode.window.createWebviewPanel).mockReturnValue(panel as unknown as vscode.WebviewPanel);
        const context = { subscriptions: [] } as unknown as vscode.ExtensionContext;

        const first = openProjectAnalysisResults(context, resultFixture());
        const second = openProjectAnalysisResults(context, resultFixture());

        expect(first).toBe(second);
        expect(vscode.window.createWebviewPanel).toHaveBeenCalledTimes(1);
        expect(vscode.window.createWebviewPanel).toHaveBeenCalledWith(
            'xlideProjectAnalysisResults',
            'XLIDE Analysis: Book.xlsm',
            vscode.ViewColumn.Active,
            expect.objectContaining({
                enableScripts: true,
                retainContextWhenHidden: true,
            }),
        );
        expect(panel.reveal).toHaveBeenCalledWith(vscode.ViewColumn.Active);
        panel.disposePanel();
    });

    it('renders scope-explicit rule tracking controls', () => {
        const html = renderProjectAnalysisResultsHtml(
            buildProjectAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: [],
                untrackedRulesSource: 'default',
                projectUntrackedRules: [],
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('data-context-action="setRuleTrackingProject"');
        expect(html).toContain('data-context-action="setRuleTrackingGlobal"');
        expect(html).toContain('Untrack In Project');
        expect(html).toContain('Untrack Globally');
        expect(html).toContain('Track In Project');
        expect(html).toContain('Track Globally');
        expect(html).toContain('Project Untracked Rules');
        expect(html).toContain('No project rules are manually untracked.');
        expect(html).toContain('Track All');
        expect(html).not.toContain('Default Visible Severities');
        expect(html).not.toContain('Rule Behavior');
        expect(html).not.toContain('data-settings-rule-severity-code');
        expect(html).not.toContain('data-settings-severity');
        expect(html).not.toContain('Workbook Rule Tracking');
        expect(html).not.toContain('Rule Severities');
        expect(html).toContain('data-tracking-source="tracked"');
        expect(html).toContain('id="trackingDivider"');
        expect(html).toContain('syncTrackingActions(row)');
        expect(html).toContain('grid-template-columns: 124px minmax(184px, 210px)');
        expect(html).toContain('trackingScope: \'project\'');
        expect(html).toContain('trackingScope: action === \'setRuleTrackingGlobal\' ? \'global\' : \'project\'');
        expect(html).not.toContain('id="trackingAction"');
    });

    it('renders analysis table headers as sortable controls', () => {
        const html = renderProjectAnalysisResultsHtml(
            buildProjectAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: [],
                untrackedRulesSource: 'default',
                projectUntrackedRules: [],
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('data-sort="severity"');
        expect(html).toContain('data-sort="message"');
        expect(html).toContain('class="sortIndicator" aria-hidden="true"');
        expect(html).toContain('syncSortHeaders();');
    });

    it('posts stable finding location data when opening a row', () => {
        const html = renderProjectAnalysisResultsHtml(
            buildProjectAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: [],
                untrackedRulesSource: 'default',
                projectUntrackedRules: [],
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('data-module="Module1"');
        expect(html).toContain('data-module-type="standard"');
        expect(html).toContain('data-line="4"');
        expect(html).toContain('data-column="2"');
        expect(html).toContain('data-end-column="8"');
        expect(html).toContain('moduleName: row.dataset.module');
        expect(html).toContain('endColumn: Number(row.dataset.endColumn)');
    });

    it('labels globally untracked rule rows distinctly', () => {
        const html = renderProjectAnalysisResultsHtml(
            buildProjectAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: ['undeclared-variable'],
                untrackedRulesSource: 'machine',
                projectUntrackedRules: [],
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('Untracked Globally');
        expect(html).toContain('data-tracked="no"');
        expect(html).toContain('data-tracking-source="global"');
    });

    it('labels project-untracked rule rows distinctly', () => {
        const html = renderProjectAnalysisResultsHtml(
            buildProjectAnalysisResultsModel(resultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: ['undeclared-variable'],
                untrackedRulesSource: 'project',
                projectUntrackedRules: ['undeclared-variable'],
                ruleSeverityOverrides: {},
                ruleSeverityOverridesSource: 'default',
            },
        );

        expect(html).toContain('Untracked In File');
        expect(html).toContain('data-tracked="no"');
        expect(html).toContain('data-tracking-source="project"');
        expect(html).toContain('<td class="settingsTableCode">undeclared-variable</td>');
        expect(html).toContain('data-settings-track-rule-code="undeclared-variable"');
    });

    it('gives untracked status precedence over suppressed findings', () => {
        const html = renderProjectAnalysisResultsHtml(
            buildProjectAnalysisResultsModel(suppressedOnlyResultFixture()),
            {
                visibleSeverities: ['error', 'warning', 'information'],
                visibleSeveritiesSource: 'default',
                untrackedRules: ['option-explicit-missing'],
                untrackedRulesSource: 'machine',
                projectUntrackedRules: [],
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

function fakeWebviewPanel(): vscode.WebviewPanel & { disposePanel: () => void } {
    let disposeHandler: (() => void) | undefined;
    return {
        title: '',
        viewColumn: vscode.ViewColumn.Active,
        webview: {
            cspSource: 'test-csp',
            html: '',
            postMessage: vi.fn(),
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
