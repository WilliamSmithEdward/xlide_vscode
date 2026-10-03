import {
    analyzeModule,
    withResolvedHostModel,
    analyzeModuleRulesIncremental,
    type ModuleRulesIncrementalState,
    DIAGNOSTIC_RULES,
    diagnosticMetadataForCode,
    incompleteExpressionEditSpan,
    normalizeDiagnosticSeverityOverride,
    createConditionalActivityTracker,
    parseModule,
    scanAnalysisSuppressions,
    tokenizeCached,
    type AnalysisFailure,
    type AnalyzeModuleOptions,
    type DiagnosticSeverity as RuleSeverity,
    type VbaDiagnosticData,
} from './analyzer';
import { errorHandlerExtents, onErrorMode } from './analyzer/diagnostics/rules/handlerFlow';
import { statementTokensAfterLeadingLabel } from './analyzer/diagnostics/walker';
import { unreachableStatementsIn } from './analyzer/diagnostics/typeInference';
import { buildModuleSymbols } from './analyzer/symbols/buildModuleSymbols';
import type { ConditionalActivityTracker } from './analyzer/conditional/conditionalCompilation';
import type { BodyNode, ModuleNode, ProcedureNode, Span } from './analyzer/parser/nodes';
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
    /** From projectAnalysisOptionsForModule: the index could not answer for this module. */
    projectContextFailure?: unknown;
    activeIncompleteExpressionOffset?: number;
    /**
     * Opt-in incremental rule re-analysis: pass the state returned by the
     * previous call (plus a fingerprint of every cross-module input) and the
     * expensive per-procedure rule walks re-run only for procedures whose body
     * changed. Any envelope change (declarations, signatures, directives) or
     * fingerprint mismatch falls back to a full pass automatically.
     */
    rulesIncremental?: {
        state?: ModuleRulesIncrementalState;
        fingerprint: readonly unknown[];
    };
}

export interface VbaModuleAnalysisResult {
    diagnostics: VbaModuleAnalysisDiagnostic[];
    suppressedDiagnostics: VbaModuleAnalysisDiagnostic[];
    suppressedCount: number;
    /** Present when rulesIncremental was requested: feed into the next call. */
    rulesIncrementalState?: ModuleRulesIncrementalState;
    rulesIncrementalMode?: 'full' | 'incremental';
    /**
     * Failures the analysis recovered from, when there were any: the checks
     * they stopped did not run, so the diagnostics above are not all there is
     * (issue #178). Plain data, so it crosses the worker boundary.
     */
    analysisFailures?: VbaModuleAnalysisFailure[];
}

/** A failure an analysis recovered from, as data. */
export interface VbaModuleAnalysisFailure {
    /** analyzeModule's own stages, and the passes this module adds around it. */
    stage: AnalysisFailure['stage'] | 'index' | 'project-context' | 'test-directives' | 'structural';
    rule?: string;
    message: string;
    stack?: string;
}

/**
 * Shared module-level analysis core. Live diagnostics, current-module analysis,
 * and project analysis all flow through this function so structural checks, semantic
 * checks, and XLIDE suppression directives cannot drift by surface.
 */
export function analyzeVbaModuleSource(input: VbaModuleAnalysisInput): VbaModuleAnalysisResult {
    const {
        source,
        moduleType,
        activeIncompleteExpressionOffset,
        rulesIncremental,
        projectContextFailure,
        ...restOptions
    } = input;
    // Resolve the caller's host token into a model ONCE, here, so the full
    // pass, the incremental pass and the structural checks all analyze under
    // the same host (issue #24). Absent keeps the Excel defaults.
    const analyzeOptions = withResolvedHostModel(restOptions);
    const analysisFailures: VbaModuleAnalysisFailure[] = [];
    const recordFailure = (error: unknown, where: Pick<VbaModuleAnalysisFailure, 'stage' | 'rule'>): void => {
        analysisFailures.push({
            ...where,
            message: error instanceof Error ? error.message : String(error),
            ...(error instanceof Error && error.stack ? { stack: error.stack } : {}),
        });
    };
    analyzeOptions.onInternalError = recordFailure;
    if (projectContextFailure !== undefined) {
        recordFailure(projectContextFailure, { stage: 'project-context' });
    }
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
    const expectedErrorRuntimeSuppressions = [
        ...expectedErrorRuntimeSuppressionRanges(
            source,
            module,
            analyzeOptions.moduleName ?? 'Module',
            moduleType ?? analyzeOptions.moduleKind ?? 'standard',
        ),
        ...onErrorResumeNextSuppressionRanges(source, module, () => ({
            symbols: buildModuleSymbols(analyzeOptions.moduleName ?? 'Module', analyzeOptions.moduleKind ?? 'standard', source, {
                conditionalCompilation: analyzeOptions.conditionalCompilation,
                parsedModule: module,
            }),
            activity: createConditionalActivityTracker(module, analyzeOptions.conditionalCompilation),
        })),
    ];

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
    } catch (err) {
        // Test directive validation should never interrupt live analysis.
        recordFailure(err, { stage: 'test-directives' });
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
    } catch (err) {
        // The structural pass is defensive; a failure should not break editing.
        recordFailure(err, { stage: 'structural' });
    }

    let rulesIncrementalState: ModuleRulesIncrementalState | undefined;
    let rulesIncrementalMode: 'full' | 'incremental' | undefined;
    try {
        let ruleDiagnostics: ReturnType<typeof analyzeModule>;
        if (rulesIncremental) {
            const inc = analyzeModuleRulesIncremental(
                source,
                analyzeOptions,
                rulesIncremental.state,
                rulesIncremental.fingerprint,
            );
            ruleDiagnostics = inc.diagnostics;
            rulesIncrementalState = inc.state;
            rulesIncrementalMode = inc.mode;
        } else {
            ruleDiagnostics = analyzeModule(source, analyzeOptions);
        }
        for (const diagnostic of ruleDiagnostics) {
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
    } catch (err) {
        // Keep analysis non-throwing while the user is typing malformed VBA.
        recordFailure(err, { stage: 'analysis' });
    }

    const deduplicatedSuppressedDiagnostics = deduplicateDiagnostics(suppressedDiagnostics);
    return {
        diagnostics: deduplicateDiagnostics(diagnostics),
        suppressedDiagnostics: deduplicatedSuppressedDiagnostics,
        suppressedCount: deduplicatedSuppressedDiagnostics.length,
        ...(rulesIncrementalState ? { rulesIncrementalState, rulesIncrementalMode } : {}),
        ...(analysisFailures.length > 0 ? { analysisFailures } : {}),
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
    // One tracker for the whole pass: constructing it replays the directive
    // stack once, and each per-line query is a binary search. Calling
    // conditionalActivityAtOffset per line instead re-walks the entire module
    // AST per query, which turns the structural pass quadratic (~10s on a
    // 24k-line module; a few hundred ms with the tracker).
    const tracker = createConditionalActivityTracker(module, analyzeOptions.conditionalCompilation);
    if (!tracker) {
        return undefined;
    }
    return (line: number): boolean => {
        const offset = starts[line] ?? source.length;
        return tracker.isInactive({ start: offset, end: offset });
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

/**
 * The stretches of each procedure under an active `On Error Resume Next`:
 * from that statement to the next `On Error` statement or the procedure's
 * end. A deterministic runtime error there is raised and handled, which is
 * usually the point of the code - `n = UBound(a)` under Resume Next is the
 * common test for an allocated array (issue #106) - so "This will raise" is
 * not the right report inside them. Branches are not modelled: a Resume Next
 * inside an If arm covers what follows it in source order, the way the VBE's
 * own handler state does once the arm runs. Inside a running error handler
 * the statement does not take effect, and it covers nothing (issue #199).
 */
function onErrorResumeNextSuppressionRanges(
    source: string,
    module: ModuleNode,
    analysisContext: () => { symbols: ReturnType<typeof buildModuleSymbols>; activity: ConditionalActivityTracker | undefined },
): ExpectedErrorRuntimeSuppression[] {
    const out: ExpectedErrorRuntimeSuppression[] = [];
    let context: ReturnType<typeof analysisContext> | undefined;
    for (const member of module.members) {
        if (member.kind !== 'Procedure') {
            continue;
        }
        const handlers: Array<{ start: number; end: number; resumeNext: boolean }> = [];
        const running = errorHandlerExtents(source, member);
        // An On Error statement a known guard keeps from running sets nothing:
        // `If False Then ... On Error Resume Next ... End If`, a For of no
        // pass, a Case that cannot match, the line after a GoTo (issue #486,
        // measured in Excel 16.0). The walk is run only where one could be dead.
        let dead: ReadonlySet<BodyNode> | undefined;
        const neverRuns = (node: BodyNode, nested: boolean, after: readonly BodyNode[]): boolean => {
            // `If True Then GoTo L` before it can leave it dead too (issue #673).
            const mayBeDead = nested || after.some((earlier) => earlier.kind === 'Statement'
                && statementTokensAfterLeadingLabel(source, earlier.span).some((tok) => ['goto', 'exit', 'end', 'resume', 'return'].includes(tok.rawText.toLowerCase())));
            if (!mayBeDead) {
                return false;
            }
            context ??= analysisContext();
            dead ??= unreachableStatementsIn(source, member, context.symbols, context.activity);
            return dead.has(node);
        };
        const visit = (body: readonly BodyNode[], nested: boolean): void => {
            for (const [index, node] of body.entries()) {
                if (node.kind === 'Statement') {
                    // Read after any line label: `10 On Error Resume Next`.
                    const mode = onErrorMode(statementTokensAfterLeadingLabel(source, node.span));
                    // GoTo 0 turns handling off and ends the stretch; GoTo -1
                    // only clears the error, and Resume Next stays (issue #313,
                    // measured in Excel 16.0).
                    if ((mode === 'resume-next' || mode === 'goto-label' || mode === 'goto-0') && !neverRuns(node, nested, body.slice(0, index))) {
                        const start = node.span.start;
                        handlers.push({
                            start,
                            end: node.span.end,
                            resumeNext: mode === 'resume-next'
                                && !running.some((extent) => start >= extent.start && start < extent.end),
                        });
                    }
                } else if ('body' in node && Array.isArray(node.body)) {
                    visit(node.body as BodyNode[], true);
                }
            }
        };
        visit(member.body, false);
        out.push(...unenteredHandlers(source, member, running));
        handlers.sort((a, b) => a.start - b.start);
        for (let i = 0; i < handlers.length; i++) {
            if (!handlers[i].resumeNext) {
                continue;
            }
            const until = handlers[i + 1]?.start ?? member.span.end;
            out.push({ span: { start: handlers[i].end, end: until }, expectedError: 'any' });
        }
    }
    return out;
}

/**
 * The handlers no error can reach: every `On Error GoTo H` naming the
 * handler's label is followed at once by Exit or End, so nothing raises
 * while it is set and what the handler would do never runs (issue #556,
 * measured in Excel 16.0).
 */
function unenteredHandlers(source: string, member: ProcedureNode, running: readonly Span[]): ExpectedErrorRuntimeSuppression[] {
    const safeAfter = new Map<string, boolean>();
    const visit = (body: readonly BodyNode[]): void => {
        for (const [index, node] of body.entries()) {
            if ('body' in node && Array.isArray(node.body)) {
                visit(node.body as BodyNode[]);
                continue;
            }
            if (node.kind !== 'Statement') {
                continue;
            }
            const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
            if (onErrorMode(toks) !== 'goto-label') {
                continue;
            }
            const label = toks[3]?.rawText.toLowerCase();
            const next = body[index + 1];
            const head = next?.kind === 'Statement' ? statementTokensAfterLeadingLabel(source, next.span).filter((tok) => tok.kind !== 'comment').map((tok) => tok.rawText.toLowerCase()) : [];
            const leaves = toks.length === 4 && (head[0] === 'exit' || (head[0] === 'end' && head.length === 1));
            if (label) {
                safeAfter.set(label, (safeAfter.get(label) ?? true) && leaves);
            }
        }
    };
    visit(member.body);
    return running.flatMap((extent) => {
        const label = /^\s*([A-Za-z_][A-Za-z0-9_]*|\d+)\s*:?/.exec(source.slice(extent.start, extent.end))?.[1]?.toLowerCase();
        return label && safeAfter.get(label) === true ? [{ span: extent, expectedError: 'any' as const }] : [];
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
    ['late-bound-friend-member', 438],
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
