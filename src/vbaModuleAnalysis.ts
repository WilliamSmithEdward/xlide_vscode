import {
    analyzeModule,
    incompleteExpressionEditSpan,
    normalizeDiagnosticSeverityOverride,
    scanAnalysisSuppressions,
    type AnalyzeModuleOptions,
    type DiagnosticSeverity as RuleSeverity,
    type VbaDiagnosticData,
} from './analyzer';
import type { Span } from './analyzer/parser/nodes';
import {
    analyzeVbaStructure,
    lineStartOffsets,
    type VbaStructuralDiagnostic,
} from './vbaStructuralAnalysis';

export interface VbaModuleAnalysisDiagnostic {
    code?: string;
    message: string;
    severity: RuleSeverity;
    span: Span;
    data?: VbaDiagnosticData;
    expectedClose?: VbaStructuralDiagnostic['expectedClose'];
    insertLine?: VbaStructuralDiagnostic['insertLine'];
}

export interface VbaModuleAnalysisInput extends AnalyzeModuleOptions {
    source: string;
    activeIncompleteExpressionOffset?: number;
    activeIncompleteMemberAccessOffset?: number;
}

export interface VbaModuleAnalysisResult {
    diagnostics: VbaModuleAnalysisDiagnostic[];
    suppressedDiagnostics: VbaModuleAnalysisDiagnostic[];
    suppressedCount: number;
}

/**
 * Shared module-level analysis core. Live diagnostics, current-module analysis,
 * and workbook analysis all flow through this function so structural checks, semantic
 * checks, and XLIDE suppression directives cannot drift by surface.
 */
export function analyzeVbaModuleSource(input: VbaModuleAnalysisInput): VbaModuleAnalysisResult {
    const {
        source,
        activeIncompleteExpressionOffset,
        activeIncompleteMemberAccessOffset,
        ...analyzeOptions
    } = input;
    const starts = lineStartOffsets(source);
    const suppressions = scanAnalysisSuppressions(source);
    const diagnostics: VbaModuleAnalysisDiagnostic[] = [...suppressions.diagnostics];
    const suppressedDiagnostics: VbaModuleAnalysisDiagnostic[] = [];
    const activeIncompleteOffset = activeIncompleteExpressionOffset ?? activeIncompleteMemberAccessOffset;
    const activeIncompleteExpressionSpan = activeIncompleteOffset === undefined
        ? undefined
        : incompleteExpressionEditSpan(source, activeIncompleteOffset);

    const isTransientIncompleteExpressionDiagnostic = (
        code: string | undefined,
        span: Span,
    ): boolean => {
        if (
            !activeIncompleteExpressionSpan ||
            (
                code !== 'invalid-expression-syntax' &&
                code !== 'scalar-member-access' &&
                code !== 'unbalanced-parens'
            )
        ) {
            return false;
        }
        return spansOverlap(span, activeIncompleteExpressionSpan);
    };

    try {
        for (const problem of analyzeVbaStructure(source)) {
            const span = {
                start: (starts[problem.line] ?? 0) + problem.startCol,
                end: (starts[problem.line] ?? 0) + problem.endCol,
            };
            if (isTransientIncompleteExpressionDiagnostic(problem.code, span)) {
                continue;
            }
            const override = normalizeDiagnosticSeverityOverride(
                problem.code,
                problem.code ? analyzeOptions.severityOverrides?.[problem.code] : undefined,
            );
            if (override === 'off') {
                continue;
            }
            const diagnostic: VbaModuleAnalysisDiagnostic = {
                code: problem.code,
                message: problem.message,
                severity: override ?? problem.severity,
                span,
                expectedClose: problem.expectedClose,
                insertLine: problem.insertLine,
            };
            if (suppressions.isDiagnosticSuppressed(problem.code, span)) {
                suppressedDiagnostics.push(diagnostic);
                continue;
            }
            diagnostics.push(diagnostic);
        }
    } catch {
        // The structural pass is defensive; a failure should not break editing.
    }

    try {
        for (const diagnostic of analyzeModule(source, analyzeOptions)) {
            if (isTransientIncompleteExpressionDiagnostic(diagnostic.code, diagnostic.span)) {
                continue;
            }
            if (suppressions.isDiagnosticSuppressed(diagnostic.code, diagnostic.span)) {
                suppressedDiagnostics.push(diagnostic);
                continue;
            }
            diagnostics.push(diagnostic);
        }
    } catch {
        // Keep analysis non-throwing while the user is typing malformed VBA.
    }

    return {
        diagnostics,
        suppressedDiagnostics,
        suppressedCount: suppressedDiagnostics.length,
    };
}

function spansOverlap(left: Span, right: Span): boolean {
    return left.start < right.end && left.end > right.start;
}
