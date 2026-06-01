import {
    analyzeModule,
    scanLintSuppressions,
    type AnalyzeModuleOptions,
    type DiagnosticSeverity as RuleSeverity,
} from './analyzer';
import type { Span } from './analyzer/parser/nodes';
import { lineStartOffsets, lintVbaSource } from './vbaLinter';

export interface VbaModuleLintDiagnostic {
    code?: string;
    message: string;
    severity: RuleSeverity;
    span: Span;
}

export interface VbaModuleLintInput extends AnalyzeModuleOptions {
    source: string;
}

export interface VbaModuleLintResult {
    diagnostics: VbaModuleLintDiagnostic[];
    suppressedCount: number;
}

/**
 * Shared module-level lint core. Live diagnostics, current-module lint, and
 * workbook lint all flow through this function so structural checks, semantic
 * checks, and XLIDE suppression directives cannot drift by surface.
 */
export function lintVbaModuleSource(input: VbaModuleLintInput): VbaModuleLintResult {
    const { source, ...analyzeOptions } = input;
    const starts = lineStartOffsets(source);
    const suppressions = scanLintSuppressions(source);
    const diagnostics: VbaModuleLintDiagnostic[] = [...suppressions.diagnostics];
    let suppressedCount = 0;

    try {
        for (const problem of lintVbaSource(source)) {
            const span = {
                start: (starts[problem.line] ?? 0) + problem.startCol,
                end: (starts[problem.line] ?? 0) + problem.endCol,
            };
            if (suppressions.isDiagnosticSuppressed(problem.code, span)) {
                suppressedCount++;
                continue;
            }
            diagnostics.push({
                code: problem.code,
                message: problem.message,
                severity: problem.severity,
                span,
            });
        }
    } catch {
        // The structural pass is defensive; a failure should not break editing.
    }

    try {
        for (const diagnostic of analyzeModule(source, analyzeOptions)) {
            if (suppressions.isDiagnosticSuppressed(diagnostic.code, diagnostic.span)) {
                suppressedCount++;
                continue;
            }
            diagnostics.push(diagnostic);
        }
    } catch {
        // Keep lint non-throwing while the user is typing malformed VBA.
    }

    return { diagnostics, suppressedCount };
}
