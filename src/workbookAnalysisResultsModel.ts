import * as path from 'path';
import type {
    WorkbookAnalysisProblem,
    WorkbookAnalysisResult,
    WorkbookAnalysisSeverity,
    WorkbookAnalysisSummaryCategory,
    WorkbookAnalysisSummaryKind,
} from './vbaWorkbookAnalysis';
import { compareVbaModulesForTreeOrder, moduleTypeBadge, moduleTypeLabel } from './moduleDisplay';

export interface WorkbookAnalysisResultRow {
    index: number;
    moduleName: string;
    moduleType: string;
    moduleIcon: string;
    moduleTypeLabel: string;
    location: string;
    line: number;
    column: number;
    endColumn: number;
    severity: WorkbookAnalysisSeverity;
    code: string;
    ruleTitle: string;
    category: WorkbookAnalysisSummaryCategory;
    diagnosticKind: WorkbookAnalysisSummaryKind;
    vbeCompileEquivalent: boolean;
    quickFixAvailable: boolean;
    quickFixTitles: string[];
    message: string;
    specReference: string;
}

export interface WorkbookAnalysisModuleGroup {
    moduleName: string;
    moduleType: string;
    moduleIcon: string;
    moduleTypeLabel: string;
    total: number;
    errorCount: number;
    warningCount: number;
    informationCount: number;
    hintCount: number;
    rows: WorkbookAnalysisResultRow[];
}

export interface WorkbookAnalysisResultsModel {
    filePath: string;
    workbookName: string;
    moduleCount: number;
    totalProblems: number;
    errorCount: number;
    warningCount: number;
    suppressedCount: number;
    vbeCompileEquivalentCount: number;
    nonVbeCompileEquivalentCount: number;
    byCategory: Array<{ name: WorkbookAnalysisSummaryCategory; count: number }>;
    byDiagnosticKind: Array<{ name: WorkbookAnalysisSummaryKind; count: number }>;
    rows: WorkbookAnalysisResultRow[];
    groups: WorkbookAnalysisModuleGroup[];
}

export function buildWorkbookAnalysisResultsModel(result: WorkbookAnalysisResult): WorkbookAnalysisResultsModel {
    const rows = result.problems.map(problemToRow);
    const groupsByModule = new Map<string, WorkbookAnalysisModuleGroup>();

    for (const row of rows) {
        const key = row.moduleName.toLowerCase();
        let group = groupsByModule.get(key);
        if (!group) {
            group = {
                moduleName: row.moduleName,
                moduleType: row.moduleType,
                moduleIcon: row.moduleIcon,
                moduleTypeLabel: row.moduleTypeLabel,
                total: 0,
                errorCount: 0,
                warningCount: 0,
                informationCount: 0,
                hintCount: 0,
                rows: [],
            };
            groupsByModule.set(key, group);
        }
        group.total += 1;
        group.rows.push(row);
        switch (row.severity) {
            case 'error':
                group.errorCount += 1;
                break;
            case 'warning':
                group.warningCount += 1;
                break;
            case 'information':
                group.informationCount += 1;
                break;
            case 'hint':
                group.hintCount += 1;
                break;
        }
    }

    const groups = [...groupsByModule.values()].sort((left, right) =>
        compareVbaModulesForTreeOrder(left, right),
    );

    return {
        filePath: result.filePath,
        workbookName: path.basename(result.filePath),
        moduleCount: result.moduleCount,
        totalProblems: rows.length,
        errorCount: result.errorCount,
        warningCount: result.warningCount,
        suppressedCount: result.summary.suppressedCount,
        vbeCompileEquivalentCount: result.summary.vbeCompileEquivalentCount,
        nonVbeCompileEquivalentCount: result.summary.nonVbeCompileEquivalentCount,
        byCategory: countEntries(result.summary.byCategory),
        byDiagnosticKind: countEntries(result.summary.byDiagnosticKind),
        rows,
        groups,
    };
}

export function buildWorkbookAnalysisPlainText(model: WorkbookAnalysisResultsModel): string {
    const lines = [
        `XLIDE Analysis Results - ${model.workbookName}`,
        `${model.totalProblems} problem(s), ${model.errorCount} error(s), ${model.warningCount} warning(s), ${model.moduleCount} module(s) checked.`,
        `VBE compile-equivalent: ${model.vbeCompileEquivalentCount}; XLIDE guidance/risk: ${model.nonVbeCompileEquivalentCount}.`,
    ];

    if (model.suppressedCount > 0) {
        lines.push(`Suppressed by XLIDE analysis directives: ${model.suppressedCount}.`);
    }
    lines.push(`Categories: ${formatCountEntries(model.byCategory)}.`);
    lines.push(`Evidence: ${formatCountEntries(model.byDiagnosticKind)}.`);
    lines.push('');

    for (const group of model.groups) {
        lines.push(`${group.moduleName} (${group.moduleType})`);
        for (const row of group.rows) {
            const code = row.code ? ` [${row.code}]` : '';
            lines.push(`  ${row.severity.toUpperCase()} ${row.location} ${row.message}${code}`);
        }
    }

    return lines.join('\n');
}

function problemToRow(problem: WorkbookAnalysisProblem, index: number): WorkbookAnalysisResultRow {
    return {
        index,
        moduleName: problem.moduleName,
        moduleType: problem.moduleType,
        moduleIcon: moduleTypeBadge(problem.moduleType),
        moduleTypeLabel: moduleTypeLabel(problem.moduleType),
        location: `${problem.moduleName}:${problem.line}:${problem.column}`,
        line: problem.line,
        column: problem.column,
        endColumn: problem.endColumn,
        severity: problem.severity,
        code: problem.code ?? '',
        ruleTitle: problem.ruleTitle ?? problem.code ?? 'Diagnostic',
        category: problem.category ?? 'uncategorized',
        diagnosticKind: problem.diagnosticKind ?? 'unknown',
        vbeCompileEquivalent: problem.vbeCompileEquivalent === true,
        quickFixAvailable: problem.quickFixAvailable === true,
        quickFixTitles: problem.quickFixTitles ?? [],
        message: problem.message,
        specReference: problem.specReference ?? '',
    };
}

function countEntries<K extends string>(counts: Partial<Record<K, number>>): Array<{ name: K; count: number }> {
    return Object.entries(counts)
        .filter((entry): entry is [K, number] => typeof entry[1] === 'number' && entry[1] > 0)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([name, count]) => ({ name, count }));
}

function formatCountEntries<K extends string>(entries: readonly { name: K; count: number }[]): string {
    return entries.length === 0
        ? 'none'
        : entries.map((entry) => `${entry.name} ${entry.count}`).join(', ');
}
