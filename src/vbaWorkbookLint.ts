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
import { PythonBridge } from './pythonBridge';
import {
    analyzeModule,
    DiagnosticSeverity as RuleSeverity,
    ModuleSymbolKind,
    ProjectIndex,
    SeverityOverrides,
    VbaDiagnostic,
} from './analyzer';
import { lintVbaSource } from './vbaLinter';

export type WorkbookLintSeverity = 'error' | 'warning' | 'information' | 'hint';

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
    /** Stable rule code (semantic rules only); undefined for structural ones. */
    code?: string;
    message: string;
}

/** Aggregate result of linting an entire workbook. */
export interface WorkbookLintResult {
    filePath: string;
    moduleCount: number;
    problems: WorkbookLintProblem[];
    errorCount: number;
    warningCount: number;
}

interface RawModule {
    name: string;
    type: string;
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

/** Precomputes the byte offset at which each line starts. */
function lineStartOffsets(source: string): number[] {
    const starts = [0];
    for (let i = 0; i < source.length; i++) {
        if (source[i] === '\n') { starts.push(i + 1); }
    }
    return starts;
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

/** Loads every module's source from the workbook (best-effort per module). */
async function loadWorkbookModules(
    bridge: PythonBridge,
    filePath: string,
): Promise<RawModule[]> {
    const list = await bridge.call<Array<{ name: string; type: string }>>(
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
            out.push({ name: entry.name, type: entry.type, source: result.source });
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

    // Project-wide procedure names enable the bare-call "Sub or Function not
    // defined" rule across modules.
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
    let knownProcedures: ReadonlySet<string> | undefined;
    try {
        knownProcedures = project.procedureNames();
    } catch {
        knownProcedures = undefined;
    }

    const config = vscode.workspace.getConfiguration('xlide');
    const optionExplicit = config.get<string>('diagnostics.optionExplicit', 'warning');
    const severities: SeverityOverrides = {
        optionExplicitMissing:
            optionExplicit === 'off' ? 'off' : (optionExplicit as RuleSeverity),
    };

    const problems: WorkbookLintProblem[] = [];

    for (const mod of modules) {
        const starts = lineStartOffsets(mod.source);

        // Structural block-balance pass (already line/column based, 0-based).
        try {
            for (const p of lintVbaSource(mod.source)) {
                problems.push({
                    moduleName: mod.name,
                    moduleType: mod.type,
                    line: p.line + 1,
                    column: p.startCol + 1,
                    endColumn: p.endCol + 1,
                    severity: p.severity,
                    message: p.message,
                });
            }
        } catch {
            // Structural linter is defensive; ignore any failure.
        }

        // Semantic rule pass (offset spans -> line/column).
        let semantic: VbaDiagnostic[];
        try {
            semantic = analyzeModule(mod.source, {
                moduleName: mod.name,
                moduleKind: moduleKindFromType(mod.type),
                severities,
                knownProcedures,
            });
        } catch {
            semantic = [];
        }
        for (const d of semantic) {
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
                message: d.message,
            });
        }
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

    return {
        filePath,
        moduleCount: modules.length,
        problems,
        errorCount,
        warningCount,
    };
}
