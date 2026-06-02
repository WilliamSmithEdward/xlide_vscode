// Workbook-wide VBA analysis. Reads every module's source from a workbook and
// runs the same two analysis passes the live editor uses - the structural
// block-balance analyzer (analyzeVbaStructure) and the high-confidence semantic rule
// engine (analyzeModule) - then flattens the findings into a single, sorted
// list of problems with 1-based line/column locations suitable for both the
// Output channel (with clickable file links) and the AI agent tool.
//
// This module owns no `vscode` UI surface beyond reading configuration, so the
// pure analysis stays reusable and testable.

import * as vscode from 'vscode';
import type { PythonBridge } from './pythonBridge';
import {
    diagnosticMetadataForCode,
    DiagnosticCategory,
    DiagnosticEvidenceKind,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    resolveDiagnosticCodeActions,
    type VbaDiagnosticData,
} from './analyzer';
import { lineStartOffsets } from './vbaStructuralAnalysis';
import { analyzeVbaModuleSource, type VbaModuleAnalysisDiagnostic } from './vbaModuleAnalysis';
import {
    buildVbaProjectIndex,
    moduleKindFromType,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
} from './vbaProjectAnalysis';
import { compareVbaModulesForTreeOrder } from './moduleDisplay';
import { openModuleSourceForWorkbook } from './vbaOpenDocuments';
import {
    validAnalysisSuppressionScopesForDiagnostic,
    type AnalysisSuppressionScope,
} from './analysisSuppressionScopes';
import { effectiveWorkbookAnalysisSettings } from './workbookAnalysisSettings';

export type WorkbookAnalysisSeverity = 'error' | 'warning' | 'information';
export type WorkbookAnalysisSummaryCategory = DiagnosticCategory | 'uncategorized';
export type WorkbookAnalysisSummaryKind = DiagnosticEvidenceKind | 'unknown';
export type { AnalysisSuppressionScope } from './analysisSuppressionScopes';

/** A single analysis finding located within one module of a workbook. */
export interface WorkbookAnalysisProblem {
    moduleName: string;
    moduleType: string;
    /** 1-based line number of the finding. */
    line: number;
    /** 1-based start column of the finding. */
    column: number;
    /** 1-based end column (exclusive) of the finding. */
    endColumn: number;
    severity: WorkbookAnalysisSeverity;
    /** Stable rule code shared by structural and semantic diagnostics. */
    code?: string;
    /** Human-readable title from the shared diagnostic metadata catalogue. */
    ruleTitle?: string;
    /** Broad diagnostic bucket used for summaries and future filtering. */
    category?: DiagnosticCategory;
    /** True when this problem should match a VBE compile failure. */
    vbeCompileEquivalent?: boolean;
    /** Evidence bucket for compile/runtime/style summary reporting. */
    diagnosticKind?: DiagnosticEvidenceKind;
    /** Optional authority or oracle note behind the diagnostic. */
    specReference?: string;
    /** Resolver metadata used by shared quick-fix actions. */
    data?: VbaDiagnosticData;
    expectedClose?: string;
    insertLine?: number;
    quickFixAvailable?: boolean;
    quickFixTitles?: string[];
    suppressionScopes: AnalysisSuppressionScope[];
    /** True when the finding is hidden by an XLIDE analysis suppression directive. */
    suppressed?: boolean;
    message: string;
}

/** Aggregate metadata summary for a workbook analysis run. */
export interface WorkbookAnalysisSummary {
    byCategory: Partial<Record<WorkbookAnalysisSummaryCategory, number>>;
    byDiagnosticKind: Partial<Record<WorkbookAnalysisSummaryKind, number>>;
    vbeCompileEquivalentCount: number;
    nonVbeCompileEquivalentCount: number;
    suppressedCount: number;
}

/** Aggregate result of analyzing an entire workbook. */
export interface WorkbookAnalysisResult {
    filePath: string;
    moduleCount: number;
    problems: WorkbookAnalysisProblem[];
    suppressedProblems: WorkbookAnalysisProblem[];
    errorCount: number;
    warningCount: number;
    summary: WorkbookAnalysisSummary;
}

interface RawModule {
    name: string;
    type: string;
    documentType?: EventHandlerDocumentType;
    source: string;
}

/** Converts a 0-based character offset to a 1-based {line, column} pair. */
function offsetToLineColumn(
    starts: number[],
    offset: number,
): { line: number; column: number } {
    let lo = 0;
    let hi = starts.length - 1;
    while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (starts[mid] <= offset) { lo = mid; } else { hi = mid - 1; }
    }
    return { line: lo + 1, column: offset - starts[lo] + 1 };
}

function severityFromRule(s: RuleSeverity): WorkbookAnalysisSeverity {
    return s;
}

function metadataFieldsForCode(
    code: string | undefined,
): Pick<
    WorkbookAnalysisProblem,
    'ruleTitle' | 'category' | 'vbeCompileEquivalent' | 'diagnosticKind' | 'specReference'
> {
    const meta = diagnosticMetadataForCode(code);
    if (!meta) {
        return {};
    }
    return {
        ruleTitle: meta.title,
        category: meta.category,
        vbeCompileEquivalent: meta.vbeCompileEquivalent,
        diagnosticKind: meta.diagnosticKind,
        specReference: meta.specReference,
    };
}

function incrementCount<K extends string>(
    counts: Partial<Record<K, number>>,
    key: K,
): void {
    counts[key] = (counts[key] ?? 0) + 1;
}

export function summarizeWorkbookAnalysisProblems(
    problems: readonly WorkbookAnalysisProblem[],
    suppressedCount: number,
): WorkbookAnalysisSummary {
    const byCategory: Partial<Record<WorkbookAnalysisSummaryCategory, number>> = {};
    const byDiagnosticKind: Partial<Record<WorkbookAnalysisSummaryKind, number>> = {};
    let vbeCompileEquivalentCount = 0;
    let nonVbeCompileEquivalentCount = 0;

    for (const problem of problems) {
        incrementCount(byCategory, problem.category ?? 'uncategorized');
        incrementCount(byDiagnosticKind, problem.diagnosticKind ?? 'unknown');
        if (problem.vbeCompileEquivalent) {
            vbeCompileEquivalentCount++;
        } else {
            nonVbeCompileEquivalentCount++;
        }
    }

    return {
        byCategory,
        byDiagnosticKind,
        vbeCompileEquivalentCount,
        nonVbeCompileEquivalentCount,
        suppressedCount,
    };
}

export function workbookProblemsForModule(
    moduleName: string,
    moduleType: string,
    source: string,
    diagnostics: readonly VbaModuleAnalysisDiagnostic[],
    options: { suppressed?: boolean } = {},
): WorkbookAnalysisProblem[] {
    const starts = lineStartOffsets(source);
    return diagnostics.map((diagnostic) => {
        const start = offsetToLineColumn(starts, diagnostic.span.start);
        const end = offsetToLineColumn(starts, diagnostic.span.end);
        const suppressionScopes = validAnalysisSuppressionScopesForDiagnostic(
            source,
            diagnostic.code,
            diagnostic.span.start,
        );
        const quickFixes = diagnostic.code
            ? resolveDiagnosticCodeActions(source, {
                code: diagnostic.code,
                message: diagnostic.message,
                span: diagnostic.span,
                expectedClose: diagnostic.expectedClose,
                insertLine: diagnostic.insertLine,
                data: diagnostic.data,
                includeSuppressionAction: false,
            })
            : [];
        return {
            moduleName,
            moduleType,
            line: start.line,
            column: start.column,
            endColumn: end.line === start.line ? end.column : start.column + 1,
            severity: severityFromRule(diagnostic.severity),
            code: diagnostic.code,
            data: diagnostic.data,
            expectedClose: diagnostic.expectedClose,
            insertLine: diagnostic.insertLine,
            quickFixAvailable: quickFixes.length > 0,
            quickFixTitles: quickFixes.map((fix) => fix.title),
            suppressionScopes,
            suppressed: options.suppressed === true,
            ...metadataFieldsForCode(diagnostic.code),
            message: diagnostic.message,
        };
    });
}

function sortWorkbookProblems(problems: WorkbookAnalysisProblem[]): void {
    problems.sort((a, b) => {
        const moduleOrder = compareVbaModulesForTreeOrder(a, b);
        if (moduleOrder !== 0) { return moduleOrder; }
        if (a.line !== b.line) { return a.line - b.line; }
        return a.column - b.column;
    });
}

/** Loads every module's source from the workbook (best-effort per module). */
async function loadWorkbookModules(
    bridge: PythonBridge,
    filePath: string,
): Promise<RawModule[]> {
    const list = await bridge.call<Array<{
        name: string;
        type: string;
        documentType?: EventHandlerDocumentType;
    }>>(
        'listModules',
        { path: filePath },
    );
    const out: RawModule[] = [];
    for (const entry of list) {
        try {
            const result = await bridge.call<{ source: string }>(
                'readModule',
                { path: filePath, module: entry.name },
            );
            out.push({
                name: entry.name,
                type: entry.type,
                documentType: entry.documentType,
                source: result.source,
            });
        } catch {
            // Skip modules that fail to read; analysis is best-effort.
        }
    }
    return out;
}

/**
 * Analyzes every module in a workbook and returns the flattened, sorted problem
 * list. Never throws on a per-module analysis failure - those modules simply
 * contribute no problems.
 */
export async function analyzeWorkbook(
    bridge: PythonBridge,
    filePath: string,
): Promise<WorkbookAnalysisResult> {
    const modules = await loadWorkbookModules(bridge, filePath);
    const openDocuments = vscode.workspace.textDocuments ?? [];
    for (const mod of modules) {
        mod.source = openModuleSourceForWorkbook(filePath, mod.name, openDocuments) ?? mod.source;
    }

    const project = buildVbaProjectIndex(modules.map((mod) => ({
        moduleName: mod.name,
        source: mod.source,
        type: mod.type,
        documentType: mod.documentType,
    })));
    const projectProcedures = projectProcedureSignatures(project);

    const analysisSettings = await effectiveWorkbookAnalysisSettings(filePath);

    const problems: WorkbookAnalysisProblem[] = [];
    const suppressedProblems: WorkbookAnalysisProblem[] = [];

    for (const mod of modules) {
        const projectOptions = projectAnalysisOptionsForModule(project, mod.name, projectProcedures);
        const moduleAnalysis = analyzeVbaModuleSource({
            source: mod.source,
            moduleName: mod.name,
            moduleKind: moduleKindFromType(mod.type),
            documentType: mod.documentType,
            severityOverrides: analysisSettings.ruleSeverityOverrides,
            ...projectOptions,
        });
        problems.push(...workbookProblemsForModule(
            mod.name,
            mod.type,
            mod.source,
            moduleAnalysis.diagnostics,
        ));
        suppressedProblems.push(...workbookProblemsForModule(
            mod.name,
            mod.type,
            mod.source,
            moduleAnalysis.suppressedDiagnostics,
            { suppressed: true },
        ));
    }

    sortWorkbookProblems(problems);
    sortWorkbookProblems(suppressedProblems);

    const errorCount = problems.filter((p) => p.severity === 'error').length;
    const warningCount = problems.filter((p) => p.severity === 'warning').length;
    const summary = summarizeWorkbookAnalysisProblems(problems, suppressedProblems.length);

    return {
        filePath,
        moduleCount: modules.length,
        problems,
        suppressedProblems,
        errorCount,
        warningCount,
        summary,
    };
}
