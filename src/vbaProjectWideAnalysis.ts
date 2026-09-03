// Workbook-wide VBA analysis. Reads every module's source from a project and
// runs the same two analysis passes the live editor uses - the structural
// block-balance analyzer (analyzeVbaStructure) and the high-confidence semantic rule
// engine (analyzeModule) - then flattens the findings into a single, sorted
// list of problems with 1-based line/column locations suitable for both the
// Output channel (with clickable file links) and the AI agent tool.
//
// This module owns no `vscode` UI surface beyond reading configuration, so the
// pure analysis stays reusable and testable.

import * as vscode from 'vscode';
import type { ProjectEngine } from './projectEngine';
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
import { hostTokenForFileName } from './analyzer/host/hostRegistry';
import {
    buildVbaProjectIndexAsync,
    moduleKindFromType,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
} from './vbaProjectAnalysis';
import { compareVbaModulesForTreeOrder } from './moduleDisplay';
import { projectIdentityKey } from './projectIdentity';
import { openModuleSourceMapForProject } from './vbaOpenDocuments';
import {
    analysisSuppressionScopeResolver,
    type AnalysisSuppressionScope,
} from './analysisSuppressionScopes';
import { effectiveProjectAnalysisSettings } from './projectAnalysisSettings';
import { measurePerformance, measurePerformanceSync, startPerformanceTrace } from './performanceTrace';
import { mapWithConcurrency, yieldToExtensionHost } from './util/async';

export type ProjectAnalysisSeverity = 'error' | 'warning' | 'information';
export type ProjectAnalysisSummaryCategory = DiagnosticCategory | 'uncategorized';
export type ProjectAnalysisSummaryKind = DiagnosticEvidenceKind | 'unknown';
export type { AnalysisSuppressionScope } from './analysisSuppressionScopes';

/** A single analysis finding located within one module of a project. */
export interface ProjectAnalysisProblem {
    moduleName: string;
    moduleType: string;
    /** 1-based line number of the finding. */
    line: number;
    /** 1-based start column of the finding. */
    column: number;
    /** 1-based end column (exclusive) of the finding. */
    endColumn: number;
    severity: ProjectAnalysisSeverity;
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

/** Aggregate metadata summary for a project analysis run. */
export interface ProjectAnalysisSummary {
    byCategory: Partial<Record<ProjectAnalysisSummaryCategory, number>>;
    byDiagnosticKind: Partial<Record<ProjectAnalysisSummaryKind, number>>;
    vbeCompileEquivalentCount: number;
    nonVbeCompileEquivalentCount: number;
    suppressedCount: number;
}

/** Aggregate result of analyzing an entire project. */
export interface ProjectAnalysisResult {
    filePath: string;
    moduleCount: number;
    problems: ProjectAnalysisProblem[];
    suppressedProblems: ProjectAnalysisProblem[];
    errorCount: number;
    warningCount: number;
    summary: ProjectAnalysisSummary;
}

interface RawModule {
    name: string;
    type: string;
    documentType?: EventHandlerDocumentType;
    source: string;
    /** A form's designer-declared controls, from the engine's designer read. */
    implicitMembers?: { name: string; type: string }[];
    /**
     * True when the module carries `Attribute VB_PredeclaredId = True`, giving
     * it a default instance so its own name is usable as a value. Absent means
     * the attribute header was not read, never "no".
     */
    predeclaredId?: boolean;
    /** The host type the module's designer makes it, when the engine could read one. */
    designerClass?: string;
}

export interface AnalyzeProjectOptions {
    progress?: (message: string) => void;
    token?: vscode.CancellationToken;
}

/**
 * The slice of AnalysisWorkerClient project analysis needs. Structural, so
 * tests can hand in a fake without touching worker_threads.
 */
export interface ProjectAnalysisWorker {
    readonly available: boolean;
    ensureSeeded(
        projectKey: string,
        generation: number,
        modules: () => Array<{
            moduleName: string;
            source: string;
            type?: string;
            documentType?: string;
            predeclaredId?: boolean;
        }>,
    ): void;
    analyze(request: {
        docKey: string;
        projectKey: string;
        generation: number;
        source: string;
        moduleName: string;
        moduleType?: string;
        moduleKind?: string;
        documentType?: string;
        severityOverrides?: Record<string, string>;
        /** Office host token for the container. Absent means Excel. */
        host?: string;
        /** The host type the module's designer makes it, when the engine read one. */
        designerClass?: string;
    }): Promise<{
        diagnostics: VbaModuleAnalysisDiagnostic[];
        suppressedDiagnostics: VbaModuleAnalysisDiagnostic[];
    }>;
}

let projectAnalysisWorker: ProjectAnalysisWorker | undefined;

/**
 * Route per-module analysis through the analysis worker thread when it is
 * healthy, exactly as live diagnostics do - a ~700ms module otherwise blocks
 * the extension host mid-command. Wired once at activation; analysis never
 * depends on it (every worker failure falls back to the identical in-host
 * pass).
 */
export function setProjectAnalysisWorker(worker: ProjectAnalysisWorker | undefined): void {
    projectAnalysisWorker = worker;
}

/** Test hook: clears the per-project result cache between test cases. */
export function resetProjectAnalysisResultCacheForTests(): void {
    lastProjectAnalysisResults.clear();
}

/**
 * Content fingerprint standing in for a seed generation. The command seeds
 * under its own key namespace (never fighting live diagnostics over a
 * project's seed), so the only requirement is that unchanged sources map to
 * the same number - a re-run then skips the seed transfer entirely - and any
 * change maps elsewhere. FNV-1a over every module's name and full source.
 */
function projectSeedFingerprint(
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

interface ProjectAnalysisProgress {
    report(message: string, options?: { force?: boolean }): void;
}

function projectAnalysisProgress(
    progress: AnalyzeProjectOptions['progress'],
): ProjectAnalysisProgress {
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

function severityFromRule(s: RuleSeverity): ProjectAnalysisSeverity {
    return s;
}

function metadataFieldsForCode(
    code: string | undefined,
): Pick<
    ProjectAnalysisProblem,
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

export function summarizeProjectAnalysisProblems(
    problems: readonly ProjectAnalysisProblem[],
    suppressedCount: number,
): ProjectAnalysisSummary {
    const byCategory: Partial<Record<ProjectAnalysisSummaryCategory, number>> = {};
    const byDiagnosticKind: Partial<Record<ProjectAnalysisSummaryKind, number>> = {};
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

export function projectProblemsForModule(
    moduleName: string,
    moduleType: string,
    source: string,
    diagnostics: readonly VbaModuleAnalysisDiagnostic[],
    options: { suppressed?: boolean } = {},
): ProjectAnalysisProblem[] {
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

function sortProjectProblems(problems: ProjectAnalysisProblem[]): void {
    problems.sort((a, b) => {
        const moduleOrder = compareVbaModulesForTreeOrder(a, b);
        if (moduleOrder !== 0) { return moduleOrder; }
        if (a.line !== b.line) { return a.line - b.line; }
        return a.column - b.column;
    });
}

/** Loads every module's source from the project (best-effort per module). */
async function loadProjectModules(
    bridge: ProjectEngine,
    filePath: string,
    progress: ProjectAnalysisProgress,
    options: AnalyzeProjectOptions = {},
): Promise<RawModule[]> {
    progress.report('Reading VBA modules...', { force: true });
    const modules = await measurePerformance(
        'analyzeProject.readModules',
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
            implicitMembers: mod.implicitMembers,
            predeclaredId: mod.predeclaredId,
            designerClass: mod.designerClass,
        }));
}

function throwIfAnalysisCancelled(token: vscode.CancellationToken | undefined): void {
    if (token?.isCancellationRequested) {
        throw new vscode.CancellationError();
    }
}

/**
 * Last completed result per project, keyed by the same content fingerprint
 * the worker seed uses (plus the analysis settings that shape diagnostics).
 * An unchanged project returns its previous result in milliseconds instead
 * of re-analyzing - which also empties the worker queue of the background
 * re-runs (results-panel refreshes) that used to stack up behind a user's
 * explicit run and make it appear hung.
 */
interface CachedProjectAnalysis {
    fingerprint: number;
    settingsKey: string;
    result: ProjectAnalysisResult;
}
const WORKBOOK_ANALYSIS_RESULT_CACHE_MAX = 8;
const lastProjectAnalysisResults = new Map<string, CachedProjectAnalysis>();

// Single-flight: concurrent analyses of the SAME project share one run, so a
// double-trigger (the analysis command + the agent tool, or a re-run) neither
// repeats the expensive read+analyze nor renders out-of-order results. The shared
// run is driven by the FIRST caller's cancellation token and progress; a later
// concurrent caller reuses that run (and, in the rare case the first caller
// cancels, observes that cancellation). Cleared when the run settles.
const inFlightProjectAnalyses = new Map<string, Promise<ProjectAnalysisResult>>();

/**
 * Analyzes every module in a project and returns the flattened, sorted problem
 * list. Never throws on a per-module analysis failure - those modules simply
 * contribute no problems. Concurrent calls for the same project are coalesced
 * into a single in-flight run.
 */
export function analyzeProject(
    bridge: ProjectEngine,
    filePath: string,
    options: AnalyzeProjectOptions = {},
): Promise<ProjectAnalysisResult> {
    const key = projectIdentityKey(filePath);
    const existing = inFlightProjectAnalyses.get(key);
    if (existing) {
        return existing;
    }
    const run = runProjectAnalysis(bridge, filePath, options);
    inFlightProjectAnalyses.set(key, run);
    return run.finally(() => {
        if (inFlightProjectAnalyses.get(key) === run) {
            inFlightProjectAnalyses.delete(key);
        }
    });
}

async function runProjectAnalysis(
    bridge: ProjectEngine,
    filePath: string,
    options: AnalyzeProjectOptions = {},
): Promise<ProjectAnalysisResult> {
    const totalTrace = startPerformanceTrace('analyzeProject.total');
    const progress = projectAnalysisProgress(options.progress);
    try {
        const modules = await loadProjectModules(bridge, filePath, progress, options);
        const openSources = openModuleSourceMapForProject(filePath);
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
            const project = await measurePerformance('analyzeProject.buildProjectContext', undefined, () =>
                buildVbaProjectIndexAsync(modules.map((mod) => ({
                    moduleName: mod.name,
                    source: mod.source,
                    type: mod.type,
                    documentType: mod.documentType,
                    implicitMembers: mod.implicitMembers,
                })), undefined, {
                    cancelIfRequested: () => throwIfAnalysisCancelled(options.token),
                }),
            );
            return { project, procedures: projectProcedureSignatures(project) };
        })();

        const analysisSettings = await measurePerformance(
            'analyzeProject.settings',
            undefined,
            () => effectiveProjectAnalysisSettings(filePath),
        );
        throwIfAnalysisCancelled(options.token);

        // Content fingerprint over every module's effective source (open-editor
        // overlays included). It keys both the worker seed and the result
        // cache; the settings that shape diagnostics join the cache key so a
        // severity-override change re-analyzes.
        const contentFingerprint = projectSeedFingerprint(modules);
        const settingsKey = JSON.stringify(analysisSettings.ruleSeverityOverrides ?? {});
        const resultCacheKey = projectIdentityKey(filePath);
        const cached = lastProjectAnalysisResults.get(resultCacheKey);
        if (cached && cached.fingerprint === contentFingerprint && cached.settingsKey === settingsKey) {
            progress.report('Analysis up to date (no changes since the last run).', { force: true });
            totalTrace.end('ok');
            return cached.result;
        }

        // Seed the worker under the command's own key namespace so it never
        // fights live diagnostics over a project's editor-driven seed, keyed
        // by content so an unchanged re-run skips the transfer.
        const worker = projectAnalysisWorker;
        const workerAvailable = worker?.available === true;
        const seedKey = `project-analysis:${projectIdentityKey(filePath)}`;
        const seedGeneration = contentFingerprint;
        if (workerAvailable && worker) {
            worker.ensureSeeded(seedKey, seedGeneration, () => modules.map((mod) => ({
                moduleName: mod.name,
                source: mod.source,
                type: mod.type,
                documentType: mod.documentType,
                implicitMembers: mod.implicitMembers,
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
                            'analyzeProject.analyzeModule',
                            mod.name,
                            () => worker.analyze({
                                // Stable per (project, module): a re-run reuses
                                // the worker's incremental state and re-analyzes
                                // only what changed since the last run.
                                docKey: `${seedKey}:${mod.name.toLowerCase()}`,
                                projectKey: seedKey,
                                generation: seedGeneration,
                                source: mod.source,
                                moduleName: mod.name,
                                moduleType: mod.type,
                                moduleKind: moduleKindFromType(mod.type),
                                documentType: mod.documentType,
                                severityOverrides: analysisSettings.ruleSeverityOverrides,
                                designerClass: mod.designerClass,
                                host: hostTokenForFileName(filePath),
                            }),
                        );
                        throwIfAnalysisCancelled(options.token);
                        reportModuleDone(mod.name);
                        return {
                            problems: projectProblemsForModule(
                                mod.name,
                                mod.type,
                                mod.source,
                                workerResult.diagnostics,
                            ),
                            suppressedProblems: projectProblemsForModule(
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
                    'analyzeProject.analyzeModule',
                    mod.name,
                    () => analyzeVbaModuleSource({
                        source: mod.source,
                        moduleName: mod.name,
                        moduleType: mod.type,
                        moduleKind: moduleKindFromType(mod.type),
                        documentType: mod.documentType,
                        severityOverrides: analysisSettings.ruleSeverityOverrides,
                        designerClass: mod.designerClass,
                        ...projectOptions,
                        host: hostTokenForFileName(filePath),
                    }),
                );
                reportModuleDone(mod.name);
                return {
                    problems: projectProblemsForModule(
                        mod.name,
                        mod.type,
                        mod.source,
                        moduleAnalysis.diagnostics,
                    ),
                    suppressedProblems: projectProblemsForModule(
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
        sortProjectProblems(problems);
        sortProjectProblems(suppressedProblems);

        const errorCount = problems.filter((p) => p.severity === 'error').length;
        const warningCount = problems.filter((p) => p.severity === 'warning').length;
        const summary = summarizeProjectAnalysisProblems(problems, suppressedProblems.length);

        totalTrace.end('ok');
        const analysisResult: ProjectAnalysisResult = {
            filePath,
            moduleCount: modules.length,
            problems,
            suppressedProblems,
            errorCount,
            warningCount,
            summary,
        };
        lastProjectAnalysisResults.delete(resultCacheKey);
        lastProjectAnalysisResults.set(resultCacheKey, {
            fingerprint: contentFingerprint,
            settingsKey,
            result: analysisResult,
        });
        while (lastProjectAnalysisResults.size > WORKBOOK_ANALYSIS_RESULT_CACHE_MAX) {
            const oldest = lastProjectAnalysisResults.keys().next().value;
            if (oldest === undefined) { break; }
            lastProjectAnalysisResults.delete(oldest);
        }
        return analysisResult;
    } catch (err) {
        totalTrace.end(err instanceof vscode.CancellationError ? 'canceled' : 'failed');
        throw err;
    }
}
