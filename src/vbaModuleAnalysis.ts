import {
    analyzeModule,
    DIAGNOSTIC_RULES,
    diagnosticMetadataForCode,
    incompleteExpressionEditSpan,
    normalizeDiagnosticSeverityOverride,
    conditionalActivityAtOffset,
    parseModule,
    scanAnalysisSuppressions,
    tokenizeCached,
    type AnalyzeModuleOptions,
    type DiagnosticSeverity as RuleSeverity,
    type VbaDiagnosticData,
} from './analyzer';
import type { ModuleNode, ProcedureNode, Span } from './analyzer/parser/nodes';
import { lineStartOffsets } from './vbaSourceScan';
import {
    analyzeVbaStructure,
    type VbaStructuralDiagnostic,
} from './vbaStructuralDiagnostics';
import { discoverVbaTestsFromModule, validateVbaTestDirectivesFromModule } from './vbaTestRunner';

export interface VbaModuleAnalysisDiagnostic {
    code?: string;
    message: string;
    severity: RuleSeverity;
    span: Span;
    data?: VbaDiagnosticData;
    expectedClose?: VbaStructuralDiagnostic['expectedClose'];
    insertLine?: VbaStructuralDiagnostic['insertLine'];
    expectedCloseReplacementSpan?: Span;
    expectedCloseReplacementText?: string;
}

export interface VbaModuleAnalysisInput extends AnalyzeModuleOptions {
    source: string;
    moduleType?: string;
    activeIncompleteExpressionOffset?: number;
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
        moduleType,
        activeIncompleteExpressionOffset,
        ...analyzeOptions
    } = input;
    const starts = lineStartOffsets(source);
    // Lex and parse once per invocation; every pass below reuses these results.
    const module = analyzeOptions.parsedModule ?? parseModule(source);
    analyzeOptions.parsedModule = module;
    const suppressions = scanAnalysisSuppressions(source, {
        tokens: tokenizeCached(source),
        parsedModule: module,
    });
    const diagnostics: VbaModuleAnalysisDiagnostic[] = [...suppressions.diagnostics];
    const suppressedDiagnostics: VbaModuleAnalysisDiagnostic[] = [];
    const activeIncompleteExpressionSpan = activeIncompleteExpressionOffset === undefined
        ? undefined
        : incompleteExpressionEditSpan(source, activeIncompleteExpressionOffset);
    const expectedErrorRuntimeSuppressions = expectedErrorRuntimeSuppressionRanges(
        source,
        module,
        analyzeOptions.moduleName ?? 'Module',
        moduleType ?? analyzeOptions.moduleKind ?? 'standard',
    );

    try {
        const meta = DIAGNOSTIC_RULES.vbaTestDirective;
        const override = normalizeDiagnosticSeverityOverride(
            meta.code,
            analyzeOptions.severityOverrides?.[meta.code],
        );
        if (override !== 'off') {
            for (const issue of validateVbaTestDirectivesFromModule({
                name: analyzeOptions.moduleName ?? 'Module',
                type: moduleType ?? analyzeOptions.moduleKind ?? 'standard',
                source,
            }, module)) {
                const diagnostic: VbaModuleAnalysisDiagnostic = {
                    code: meta.code,
                    message: issue.message,
                    severity: override ?? meta.defaultSeverity,
                    span: issue.span,
                };
                if (suppressions.isDiagnosticSuppressed(meta.code, issue.span)) {
                    suppressedDiagnostics.push(diagnostic);
                    continue;
                }
                diagnostics.push(diagnostic);
            }
        }
    } catch {
        // Test directive validation should never interrupt live analysis.
    }

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
        const isInactiveLine = inactiveConditionalLinePredicate(source, module, starts, analyzeOptions);
        for (const problem of analyzeVbaStructure(source, { isInactiveLine })) {
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
                expectedCloseReplacementSpan: problem.expectedCloseReplacement
                    ? {
                        start: (starts[problem.expectedCloseReplacement.line] ?? 0) +
                            problem.expectedCloseReplacement.startCol,
                        end: (starts[problem.expectedCloseReplacement.line] ?? 0) +
                            problem.expectedCloseReplacement.endCol,
                    }
                    : undefined,
                expectedCloseReplacementText: problem.expectedCloseReplacement?.text,
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
            if (isExpectedErrorRuntimeDiagnosticSuppressed(diagnostic, expectedErrorRuntimeSuppressions)) {
                suppressedDiagnostics.push(diagnostic);
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

    const deduplicatedSuppressedDiagnostics = deduplicateDiagnostics(suppressedDiagnostics);
    return {
        diagnostics: deduplicateDiagnostics(diagnostics),
        suppressedDiagnostics: deduplicatedSuppressedDiagnostics,
        suppressedCount: deduplicatedSuppressedDiagnostics.length,
    };
}

function deduplicateDiagnostics(
    diagnostics: readonly VbaModuleAnalysisDiagnostic[],
): VbaModuleAnalysisDiagnostic[] {
    const result: VbaModuleAnalysisDiagnostic[] = [];
    const indexByKey = new Map<string, number>();
    for (const diagnostic of diagnostics) {
        const key = diagnosticIdentityKey(diagnostic);
        const existingIndex = indexByKey.get(key);
        if (existingIndex === undefined) {
            indexByKey.set(key, result.length);
            result.push(diagnostic);
            continue;
        }
        // Later passes carry richer analyzer-specific wording for the same rule/span.
        result[existingIndex] = diagnostic;
    }
    return result;
}

function diagnosticIdentityKey(diagnostic: VbaModuleAnalysisDiagnostic): string {
    return `${diagnostic.code ?? ''}:${diagnostic.span.start}:${diagnostic.span.end}`;
}

function inactiveConditionalLinePredicate(
    source: string,
    module: ModuleNode,
    starts: readonly number[],
    analyzeOptions: AnalyzeModuleOptions,
): ((line: number) => boolean) | undefined {
    if (!source.includes('#')) {
        return undefined;
    }
    return (line: number): boolean => {
        const offset = starts[line] ?? source.length;
        return conditionalActivityAtOffset(
            module,
            offset,
            analyzeOptions.conditionalCompilation,
        ) === 'inactive';
    };
}

interface ExpectedErrorRuntimeSuppression {
    span: Span;
    expectedError: number | 'any';
}

function expectedErrorRuntimeSuppressionRanges(
    source: string,
    module: ModuleNode,
    moduleName: string,
    moduleType: string,
): ExpectedErrorRuntimeSuppression[] {
    const tests = discoverVbaTestsFromModule({
        name: moduleName,
        type: moduleType,
        source,
    }, module).filter((test) => test.metadata.expectedError);
    if (tests.length === 0) {
        return [];
    }

    const byProcedureName = new Map<string, number | 'any'>();
    for (const test of tests) {
        if (test.metadata.expectedError) {
            byProcedureName.set(test.procedureName.toLowerCase(), test.metadata.expectedError);
        }
    }

    return module.members
        .filter((member): member is ProcedureNode => member.kind === 'Procedure')
        .flatMap((member) => {
            const expectedError = byProcedureName.get(member.name.toLowerCase());
            return expectedError ? [{ span: member.span, expectedError }] : [];
        });
}

function isExpectedErrorRuntimeDiagnosticSuppressed(
    diagnostic: VbaModuleAnalysisDiagnostic,
    suppressions: readonly ExpectedErrorRuntimeSuppression[],
): boolean {
    if (suppressions.length === 0 || !isDeterministicRuntimeDiagnostic(diagnostic.code)) {
        return false;
    }
    const range = suppressions.find((candidate) => spanStartsInside(diagnostic.span, candidate.span));
    if (!range) {
        return false;
    }
    if (range.expectedError === 'any') {
        return true;
    }
    return runtimeErrorNumberForDiagnostic(diagnostic) === range.expectedError;
}

function isDeterministicRuntimeDiagnostic(code: string | undefined): boolean {
    return diagnosticMetadataForCode(code)?.diagnosticKind === 'deterministic-runtime-error';
}

const RUNTIME_ERROR_NUMBER_BY_DIAGNOSTIC_CODE = new Map<string, number>([
    ['argument-type-mismatch', 13],
    ['assignment-type-mismatch', 13],
    ['object-variable-not-set', 91],
    ['unallocated-dynamic-array-access', 9],
    ['redim-impossible-bounds', 9],
    ['runtime-conversion-value', 13],
    ['string-arithmetic-coercion', 13],
    ['runtime-argument-value', 5],
]);

function runtimeErrorNumberForDiagnostic(diagnostic: VbaModuleAnalysisDiagnostic): number | undefined {
    const fromMessage = /(?:Run-time error|VBA error)\s*'?(\d+)'?/i.exec(diagnostic.message)?.[1];
    if (fromMessage) {
        const parsed = Number(fromMessage);
        if (Number.isSafeInteger(parsed) && parsed > 0) {
            return parsed;
        }
    }
    return diagnostic.code ? RUNTIME_ERROR_NUMBER_BY_DIAGNOSTIC_CODE.get(diagnostic.code) : undefined;
}

function spanStartsInside(inner: Span, outer: Span): boolean {
    return inner.start >= outer.start && inner.start < outer.end;
}

function spansOverlap(left: Span, right: Span): boolean {
    return left.start < right.end && left.end > right.start;
}
