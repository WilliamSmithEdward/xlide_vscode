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
import type { WorkbookEngine } from './workbookEngine';
import {
    diagnosticMetadataForCode,
    DiagnosticCategory,
    DiagnosticEvidenceKind,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    resolveDiagnosticCodeActions,
    type VbaDiagnosticData,
} from './analyzer';
import { lineStartOffsets } from './vbaSourceScan';
import { analyzeVbaModuleSource, type VbaModuleAnalysisDiagnostic } from './vbaModuleAnalysis';
import {
    buildVbaProjectIndexAsync,
    moduleKindFromType,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
} from './vbaProjectAnalysis';
import { compareVbaModulesForTreeOrder } from './moduleDisplay';
import { workbookIdentityKey } from './workbookIdentity';
import { openModuleSourceMapForWorkbook } from './vbaOpenDocuments';
import {
    analysisSuppressionScopeResolver,
    type AnalysisSuppressionScope,
} from './analysisSuppressionScopes';
import { effectiveWorkbookAnalysisSettings } from './workbookAnalysisSettings';
import { measurePerformance, measurePerformanceSync, startPerformanceTrace } from './performanceTrace';
import { mapWithConcurrency, yieldToExtensionHost } from './util/async';

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
    expectedCloseReplacementSpan?: VbaModuleAnalysisDiagnostic['expectedCloseReplacementSpan'];
    expectedCloseReplacementText?: string;
    quickFixAvailable?: boolean;
    quickFixTitles?: string[];
    suppressionScopes: AnalysisSuppressionScope[];
    /** True when the finding is hidden by an XLIDE analysis suppression directive. */
    suppressed?: boolean;
    message: string;
    /**
     * VS Code document version the finding's coordinates were computed against, so
     * mutating actions (quick-fix / ignore) can detect a drifted source and refuse
     * to apply a mislocated edit. Undefined for producers that do not set it.
     */
    documentVersion?: number;
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

export interface AnalyzeWorkbookOptions {
    progress?: (message: string) => void;
    token?: vscode.CancellationToken;
}

/**
 * The slice of AnalysisWorkerClient workbook analysis needs. Structural, so
 * tests can hand in a fake without touching worker_threads.
 */
export interface WorkbookAnalysisWorker {
    readonly available: boolean;
    ensureSeeded(
        workbookKey: string,
        generation: number,
        modules: () => Array<{ moduleName: string; source: string; type?: string; documentType?: string }>,
    ): void;
    analyze(request: {
        docKey: string;
        workbookKey: string;
        generation: number;
        source: string;
        moduleName: string;
        moduleType?: string;
        moduleKind?: string;
        documentType?: string;
        severityOverrides?: Record<string, string>;
    }): Promise<{
        diagnostics: VbaModuleAnalysisDiagnostic[];
        suppressedDiagnostics: VbaModuleAnalysisDiagnostic[];
    }>;
}

let workbookAnalysisWorker: WorkbookAnalysisWorker | undefined;

/**
 * Route per-module analysis through the analysis worker thread when it is
 * healthy, exactly as live diagnostics do - a ~700ms module otherwise blocks
 * the extension host mid-command. Wired once at activation; analysis never
 * depends on it (every worker failure falls back to the identical in-host
 * pass).
 */
export function setWorkbookAnalysisWorker(worker: WorkbookAnalysisWorker | undefined): void {
    workbookAnalysisWorker = worker;
}

/** Test hook: clears the per-workbook result cache between test cases. */
export function resetWorkbookAnalysisResultCacheForTests(): void {
    lastWorkbookAnalysisResults.clear();
}

/**
 * Content fingerprint standing in for a seed generation. The command seeds
 * under its own key namespace (never fighting live diagnostics over a
 * workbook's seed), so the only requirement is that unchanged sources map to
 * the same number - a re-run then skips the seed transfer entirely - and any
 * change maps elsewhere. FNV-1a over every module's name and full source.
 */
function workbookSeedFingerprint(
    modules: ReadonlyArray<{ name: string; source: string }>,
): number {
    let hash = 0x811c9dc5;
    const mix = (text: string): void => {
        for (let i = 0; i < text.length; i++) {
            hash ^= text.charCodeAt(i);
            hash = Math.imul(hash, 0x01000193);
        }
    };
    for (const mod of modules) {
        mix(mod.name);
        mix('\u0000');
        mix(mod.source);
        mix('\u0000');
    }
    return hash >>> 0;
}

const WORKBOOK_ANALYSIS_PROGRESS_MIN_INTERVAL_MS = 100;
const WORKBOOK_MODULE_ANALYSIS_CONCURRENCY = 4;

interface WorkbookAnalysisProgress {
    report(message: string, options?: { force?: boolean }): void;
}

function workbookAnalysisProgress(
    progress: AnalyzeWorkbookOptions['progress'],
): WorkbookAnalysisProgress {
    let lastReportAt = 0;

    return {
        report(message, options = {}) {
            if (!progress) {
                return;
            }
            const now = Date.now();
            if (options.force || now - lastReportAt >= WORKBOOK_ANALYSIS_PROGRESS_MIN_INTERVAL_MS) {
                lastReportAt = now;
                progress(message);
            }
        },
    };
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
    const suppressionScopesFor = analysisSuppressionScopeResolver(source);
    return diagnostics.map((diagnostic) => {
        const start = offsetToLineColumn(starts, diagnostic.span.start);
        const end = offsetToLineColumn(starts, diagnostic.span.end);
        const suppressionScopes = suppressionScopesFor(
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
                expectedCloseReplacementSpan: diagnostic.expectedCloseReplacementSpan,
                expectedCloseReplacementText: diagnostic.expectedCloseReplacementText,
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
            expectedCloseReplacementSpan: diagnostic.expectedCloseReplacementSpan,
            expectedCloseReplacementText: diagnostic.expectedCloseReplacementText,
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
    bridge: WorkbookEngine,
    filePath: string,
    progress: WorkbookAnalysisProgress,
    options: AnalyzeWorkbookOptions = {},
): Promise<RawModule[]> {
    progress.report('Reading VBA modules...', { force: true });
    const modules = await measurePerformance(
        'analyzeWorkbook.readModules',
        undefined,
        () => bridge.call<RawModule[]>(
            'readModules',
            { path: filePath },
            options.token,
        ),
    );
    throwIfAnalysisCancelled(options.token);
    return modules
        .filter((mod) => typeof mod.source === 'string')
        .map((mod) => ({
            name: mod.name,
            type: mod.type,
            documentType: mod.documentType,
            source: mod.source,
        }));
}

function throwIfAnalysisCancelled(token: vscode.CancellationToken | undefined): void {
    if (token?.isCancellationRequested) {
        throw new vscode.CancellationError();
    }
}

/**
 * Last completed result per workbook, keyed by the same content fingerprint
 * the worker seed uses (plus the analysis settings that shape diagnostics).
 * An unchanged workbook returns its previous result in milliseconds instead
 * of re-analyzing - which also empties the worker queue of the background
 * re-runs (results-panel refreshes) that used to stack up behind a user's
 * explicit run and make it appear hung.
 */
interface CachedWorkbookAnalysis {
    fingerprint: number;
    settingsKey: string;
    result: WorkbookAnalysisResult;
}
const WORKBOOK_ANALYSIS_RESULT_CACHE_MAX = 8;
const lastWorkbookAnalysisResults = new Map<string, CachedWorkbookAnalysis>();

// Single-flight: concurrent analyses of the SAME workbook share one run, so a
// double-trigger (the analysis command + the agent tool, or a re-run) neither
// repeats the expensive read+analyze nor renders out-of-order results. The shared
// run is driven by the FIRST caller's cancellation token and progress; a later
// concurrent caller reuses that run (and, in the rare case the first caller
// cancels, observes that cancellation). Cleared when the run settles.
const inFlightWorkbookAnalyses = new Map<string, Promise<WorkbookAnalysisResult>>();

/**
 * Analyzes every module in a workbook and returns the flattened, sorted problem
 * list. Never throws on a per-module analysis failure - those modules simply
 * contribute no problems. Concurrent calls for the same workbook are coalesced
 * into a single in-flight run.
 */
export function analyzeWorkbook(
    bridge: WorkbookEngine,
    filePath: string,
    options: AnalyzeWorkbookOptions = {},
): Promise<WorkbookAnalysisResult> {
    const key = workbookIdentityKey(filePath);
    const existing = inFlightWorkbookAnalyses.get(key);
    if (existing) {
        return existing;
    }
    const run = runWorkbookAnalysis(bridge, filePath, options);
    inFlightWorkbookAnalyses.set(key, run);
    return run.finally(() => {
        if (inFlightWorkbookAnalyses.get(key) === run) {
            inFlightWorkbookAnalyses.delete(key);
        }
    });
}

async function runWorkbookAnalysis(
    bridge: WorkbookEngine,
    filePath: string,
    options: AnalyzeWorkbookOptions = {},
): Promise<WorkbookAnalysisResult> {
    const totalTrace = startPerformanceTrace('analyzeWorkbook.total');
    const progress = workbookAnalysisProgress(options.progress);
    try {
        const modules = await loadWorkbookModules(bridge, filePath, progress, options);
        const openSources = openModuleSourceMapForWorkbook(filePath);
        for (const mod of modules) {
            mod.source = openSources.get(mod.name.toLowerCase()) ?? mod.source;
        }

        throwIfAnalysisCancelled(options.token);

        // Host-side project context is only needed by the in-host fallback, so
        // build it lazily (and once) instead of paying for it on the worker
        // path. mapWithConcurrency callbacks may race to it; the shared promise
        // makes the build single-flight.
        let hostContext: Promise<{
            project: Awaited<ReturnType<typeof buildVbaProjectIndexAsync>>;
            procedures: ReturnType<typeof projectProcedureSignatures>;
        }> | undefined;
        const ensureHostContext = () => hostContext ??= (async () => {
            progress.report('Building project context...', { force: true });
            const project = await measurePerformance('analyzeWorkbook.buildProjectContext', undefined, () =>
                buildVbaProjectIndexAsync(modules.map((mod) => ({
                    moduleName: mod.name,
                    source: mod.source,
                    type: mod.type,
                    documentType: mod.documentType,
                })), undefined, {
                    cancelIfRequested: () => throwIfAnalysisCancelled(options.token),
                }),
            );
            return { project, procedures: projectProcedureSignatures(project) };
        })();

        const analysisSettings = await measurePerformance(
            'analyzeWorkbook.settings',
            undefined,
            () => effectiveWorkbookAnalysisSettings(filePath),
        );
        throwIfAnalysisCancelled(options.token);

        // Content fingerprint over every module's effective source (open-editor
        // overlays included). It keys both the worker seed and the result
        // cache; the settings that shape diagnostics join the cache key so a
        // severity-override change re-analyzes.
        const contentFingerprint = workbookSeedFingerprint(modules);
        const settingsKey = JSON.stringify(analysisSettings.ruleSeverityOverrides ?? {});
        const resultCacheKey = workbookIdentityKey(filePath);
        const cached = lastWorkbookAnalysisResults.get(resultCacheKey);
        if (cached && cached.fingerprint === contentFingerprint && cached.settingsKey === settingsKey) {
            progress.report('Analysis up to date (no changes since the last run).', { force: true });
            totalTrace.end('ok');
            return cached.result;
        }

        // Seed the worker under the command's own key namespace so it never
        // fights live diagnostics over a workbook's editor-driven seed, keyed
        // by content so an unchanged re-run skips the transfer.
        const worker = workbookAnalysisWorker;
        const workerAvailable = worker?.available === true;
        const seedKey = `workbook-analysis:${workbookIdentityKey(filePath)}`;
        const seedGeneration = contentFingerprint;
        if (workerAvailable && worker) {
            worker.ensureSeeded(seedKey, seedGeneration, () => modules.map((mod) => ({
                moduleName: mod.name,
                source: mod.source,
                type: mod.type,
                documentType: mod.documentType,
            })));
        }

        // Progress must advance on COMPLETION, forced past the throttle. The
        // per-module start reports all fire within the first few milliseconds
        // now that analysis is async on the worker, so the 100ms throttle
        // drops every one of them and the toast sits on "Reading VBA
        // modules..." for the whole run - which reads as a hang, however fast
        // the run actually is.
        let completedModules = 0;
        const reportModuleDone = (name: string): void => {
            completedModules++;
            progress.report(
                `Analyzed ${name} (${completedModules}/${modules.length})`,
                { force: true },
            );
        };
        const analysisResults = await mapWithConcurrency(
            modules,
            WORKBOOK_MODULE_ANALYSIS_CONCURRENCY,
            async (mod, index) => {
                throwIfAnalysisCancelled(options.token);
                progress.report(`Analyzing ${mod.name} (${index + 1}/${modules.length})...`);
                await yieldToExtensionHost();
                throwIfAnalysisCancelled(options.token);
                if (workerAvailable && worker) {
                    try {
                        const workerResult = await measurePerformance(
                            'analyzeWorkbook.analyzeModule',
                            mod.name,
                            () => worker.analyze({
                                // Stable per (workbook, module): a re-run reuses
                                // the worker's incremental state and re-analyzes
                                // only what changed since the last run.
                                docKey: `${seedKey}:${mod.name.toLowerCase()}`,
                                workbookKey: seedKey,
                                generation: seedGeneration,
                                source: mod.source,
                                moduleName: mod.name,
                                moduleType: mod.type,
                                moduleKind: moduleKindFromType(mod.type),
                                documentType: mod.documentType,
                                severityOverrides: analysisSettings.ruleSeverityOverrides,
                            }),
                        );
                        throwIfAnalysisCancelled(options.token);
                        reportModuleDone(mod.name);
                        return {
                            problems: workbookProblemsForModule(
                                mod.name,
                                mod.type,
                                mod.source,
                                workerResult.diagnostics,
                            ),
                            suppressedProblems: workbookProblemsForModule(
                                mod.name,
                                mod.type,
                                mod.source,
                                workerResult.suppressedDiagnostics,
                                { suppressed: true },
                            ),
                        };
                    } catch (err) {
                        if (err instanceof vscode.CancellationError) {
                            throw err;
                        }
                        // Worker died or rejected: identical in-host pass below.
                    }
                }
                throwIfAnalysisCancelled(options.token);
                const { project, procedures } = await ensureHostContext();
                const projectOptions = projectAnalysisOptionsForModule(project, mod.name, procedures);
                const moduleAnalysis = measurePerformanceSync(
                    'analyzeWorkbook.analyzeModule',
                    mod.name,
                    () => analyzeVbaModuleSource({
                        source: mod.source,
                        moduleName: mod.name,
                        moduleType: mod.type,
                        moduleKind: moduleKindFromType(mod.type),
                        documentType: mod.documentType,
                        severityOverrides: analysisSettings.ruleSeverityOverrides,
                        ...projectOptions,
                    }),
                );
                reportModuleDone(mod.name);
                return {
                    problems: workbookProblemsForModule(
                        mod.name,
                        mod.type,
                        mod.source,
                        moduleAnalysis.diagnostics,
                    ),
                    suppressedProblems: workbookProblemsForModule(
                        mod.name,
                        mod.type,
                        mod.source,
                        moduleAnalysis.suppressedDiagnostics,
                        { suppressed: true },
                    ),
                };
            },
        );
        const problems = analysisResults.flatMap((result) => result.problems);
        const suppressedProblems = analysisResults.flatMap((result) => result.suppressedProblems);

        progress.report('Preparing results...', { force: true });
        sortWorkbookProblems(problems);
        sortWorkbookProblems(suppressedProblems);

        const errorCount = problems.filter((p) => p.severity === 'error').length;
        const warningCount = problems.filter((p) => p.severity === 'warning').length;
        const summary = summarizeWorkbookAnalysisProblems(problems, suppressedProblems.length);

        totalTrace.end('ok');
        const analysisResult: WorkbookAnalysisResult = {
            filePath,
            moduleCount: modules.length,
            problems,
            suppressedProblems,
            errorCount,
            warningCount,
            summary,
        };
        lastWorkbookAnalysisResults.delete(resultCacheKey);
        lastWorkbookAnalysisResults.set(resultCacheKey, {
            fingerprint: contentFingerprint,
            settingsKey,
            result: analysisResult,
        });
        while (lastWorkbookAnalysisResults.size > WORKBOOK_ANALYSIS_RESULT_CACHE_MAX) {
            const oldest = lastWorkbookAnalysisResults.keys().next().value;
            if (oldest === undefined) { break; }
            lastWorkbookAnalysisResults.delete(oldest);
        }
        return analysisResult;
    } catch (err) {
        totalTrace.end(err instanceof vscode.CancellationError ? 'canceled' : 'failed');
        throw err;
    }
}
