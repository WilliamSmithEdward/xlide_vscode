// Workbook-wide VBA linting. Reads every module's source from a workbook and
// runs the same two analysis passes the live editor uses - the structural
// block-balance linter (lintVbaSource) and the high-confidence semantic rule
// engine (analyzeModule) - then flattens the findings into a single, sorted
// list of problems with 1-based line/column locations suitable for both the
// Output channel (with clickable file links) and the AI agent tool.
//
// This module owns no `vscode` UI surface beyond reading configuration, so the
// pure analysis stays reusable and testable.

import * as vscode from 'vscode';
import type { PythonBridge } from './pythonBridge';
import {
    analyzeModule,
    diagnosticMetadataForCode,
    DiagnosticCategory,
    DiagnosticEvidenceKind,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    ModuleSymbolKind,
    ProjectIndex,
    scanLintSuppressions,
    SeverityOverrides,
    VbaDiagnostic,
} from './analyzer';
import { lineStartOffsets, lintVbaSource } from './vbaLinter';

export type WorkbookLintSeverity = 'error' | 'warning' | 'information' | 'hint';
export type WorkbookLintSummaryCategory = DiagnosticCategory | 'uncategorized';
export type WorkbookLintSummaryKind = DiagnosticEvidenceKind | 'unknown';

/** A single lint finding located within one module of a workbook. */
export interface WorkbookLintProblem {
    moduleName: string;
    moduleType: string;
    /** 1-based line number of the finding. */
    line: number;
    /** 1-based start column of the finding. */
    column: number;
    /** 1-based end column (exclusive) of the finding. */
    endColumn: number;
    severity: WorkbookLintSeverity;
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
    message: string;
}

/** Aggregate metadata summary for a workbook lint run. */
export interface WorkbookLintSummary {
    byCategory: Partial<Record<WorkbookLintSummaryCategory, number>>;
    byDiagnosticKind: Partial<Record<WorkbookLintSummaryKind, number>>;
    vbeCompileEquivalentCount: number;
    nonVbeCompileEquivalentCount: number;
    suppressedCount: number;
}

/** Aggregate result of linting an entire workbook. */
export interface WorkbookLintResult {
    filePath: string;
    moduleCount: number;
    problems: WorkbookLintProblem[];
    errorCount: number;
    warningCount: number;
    summary: WorkbookLintSummary;
}

interface RawModule {
    name: string;
    type: string;
    documentType?: EventHandlerDocumentType;
    source: string;
}

function moduleKindFromType(type?: string): ModuleSymbolKind {
    switch (type) {
        case 'class': return 'class';
        case 'document': return 'document';
        case 'userform': return 'userform';
        default: return 'standard';
    }
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

function severityFromRule(s: RuleSeverity): WorkbookLintSeverity {
    return s;
}

function metadataFieldsForCode(
    code: string | undefined,
): Pick<
    WorkbookLintProblem,
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

function summarizeProblems(
    problems: readonly WorkbookLintProblem[],
    suppressedCount: number,
): WorkbookLintSummary {
    const byCategory: Partial<Record<WorkbookLintSummaryCategory, number>> = {};
    const byDiagnosticKind: Partial<Record<WorkbookLintSummaryKind, number>> = {};
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
            // Skip modules that fail to read; lint is best-effort.
        }
    }
    return out;
}

/**
 * Lints every module in a workbook and returns the flattened, sorted problem
 * list. Never throws on a per-module analysis failure - those modules simply
 * contribute no problems.
 */
export async function lintWorkbook(
    bridge: PythonBridge,
    filePath: string,
): Promise<WorkbookLintResult> {
    const modules = await loadWorkbookModules(bridge, filePath);

    // Project-wide names enable cross-module call and Option Explicit checks,
    // filtered per caller by VBA visibility.
    const project = new ProjectIndex();
    for (const mod of modules) {
        try {
            project.setModule({
                moduleName: mod.name,
                moduleKind: moduleKindFromType(mod.type),
                source: mod.source,
            });
        } catch {
            // Ignore parse failures for the cross-module name set.
        }
    }
    let projectProcedures: ReturnType<ProjectIndex['procedureSignatures']> | undefined;
    try {
        projectProcedures = project.procedureSignatures();
    } catch {
        projectProcedures = undefined;
    }

    const config = vscode.workspace.getConfiguration('xlide');
    const optionExplicit = config.get<string>('diagnostics.optionExplicit', 'warning');
    const severities: SeverityOverrides = {
        optionExplicitMissing:
            optionExplicit === 'off' ? 'off' : (optionExplicit as RuleSeverity),
    };

    const problems: WorkbookLintProblem[] = [];
    let suppressedCount = 0;

    for (const mod of modules) {
        const starts = lineStartOffsets(mod.source);
        const suppressions = scanLintSuppressions(mod.source);
        let moduleSuppressedCount = 0;

        for (const d of suppressions.diagnostics) {
            const start = offsetToLineColumn(starts, d.span.start);
            const end = offsetToLineColumn(starts, d.span.end);
            problems.push({
                moduleName: mod.name,
                moduleType: mod.type,
                line: start.line,
                column: start.column,
                endColumn: end.line === start.line ? end.column : start.column + 1,
                severity: severityFromRule(d.severity),
                code: d.code,
                ...metadataFieldsForCode(d.code),
                message: d.message,
            });
        }

        // Structural block-balance pass (already line/column based, 0-based).
        try {
            for (const p of lintVbaSource(mod.source)) {
                const span = {
                    start: (starts[p.line] ?? 0) + p.startCol,
                    end: (starts[p.line] ?? 0) + p.endCol,
                };
                if (suppressions.isDiagnosticSuppressed(p.code, span)) {
                    moduleSuppressedCount++;
                    continue;
                }
                problems.push({
                    moduleName: mod.name,
                    moduleType: mod.type,
                    line: p.line + 1,
                    column: p.startCol + 1,
                    endColumn: p.endCol + 1,
                    severity: p.severity,
                    code: p.code,
                    ...metadataFieldsForCode(p.code),
                    message: p.message,
                });
            }
        } catch {
            // Structural linter is defensive; ignore any failure.
        }

        // Semantic rule pass (offset spans -> line/column).
        let semantic: VbaDiagnostic[];
        let knownProcedures: ReadonlySet<string> | undefined;
        let knownIdentifiers: ReadonlySet<string> | undefined;
        let knownNonTypeNames: ReadonlySet<string> | undefined;
        let projectTypes: ReturnType<ProjectIndex['visibleTypeNames']> | undefined;
        let projectClassMembers: ReturnType<ProjectIndex['projectMemberSurfaces']> | undefined;
        try {
            knownProcedures = project.visibleProcedureNames(mod.name);
            knownIdentifiers = project.visibleIdentifierNames(mod.name);
            knownNonTypeNames = project.visibleNonTypeNames(mod.name);
            projectTypes = project.visibleTypeNames(mod.name);
            projectClassMembers = project.projectMemberSurfaces(mod.name);
        } catch {
            knownProcedures = undefined;
            knownIdentifiers = undefined;
            knownNonTypeNames = undefined;
            projectTypes = undefined;
            projectClassMembers = undefined;
        }
        try {
            semantic = analyzeModule(mod.source, {
                moduleName: mod.name,
                moduleKind: moduleKindFromType(mod.type),
                documentType: mod.documentType,
                severities,
                knownProcedures,
                knownIdentifiers,
                knownNonTypeNames,
                projectProcedures,
                projectClassMembers,
                projectTypes,
            });
        } catch {
            semantic = [];
        }
        for (const d of semantic) {
            const start = offsetToLineColumn(starts, d.span.start);
            if (suppressions.isDiagnosticSuppressed(d.code, d.span)) {
                moduleSuppressedCount++;
                continue;
            }
            const end = offsetToLineColumn(starts, d.span.end);
            problems.push({
                moduleName: mod.name,
                moduleType: mod.type,
                line: start.line,
                column: start.column,
                endColumn: end.line === start.line ? end.column : start.column + 1,
                severity: severityFromRule(d.severity),
                code: d.code,
                ...metadataFieldsForCode(d.code),
                message: d.message,
            });
        }
        suppressedCount += moduleSuppressedCount;
    }

    problems.sort((a, b) => {
        if (a.moduleName !== b.moduleName) {
            return a.moduleName.localeCompare(b.moduleName);
        }
        if (a.line !== b.line) { return a.line - b.line; }
        return a.column - b.column;
    });

    const errorCount = problems.filter((p) => p.severity === 'error').length;
    const warningCount = problems.filter((p) => p.severity === 'warning').length;
    const summary = summarizeProblems(problems, suppressedCount);

    return {
        filePath,
        moduleCount: modules.length,
        problems,
        errorCount,
        warningCount,
        summary,
    };
}
