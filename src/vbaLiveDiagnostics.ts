// Live VBA diagnostics engine: runs the analyzer's module analysis on open
// and (debounced) on every edit. Local/full pass scheduling and generation
// tracking live in DiagnosticScheduler; the engine adds a TTL'd
// analysis-settings cache and per-workbook settings-sidecar
// FileSystemWatcher lifecycle.
//
// Extracted verbatim from vbaLanguageProviders.ts (audit #21).

import * as vscode from 'vscode';
import * as path from 'path';
import {
    XLIDE_SCHEME,
    isVbaDocument,
    moduleIdentityKey,
    workbookIdentityKey,
} from './xlideFileSystem';
import { analysisSourceForDocument, moduleLocationOfDocument, moduleNameFromDocument } from './vbaDocumentIdentity';
import { onDidChangeVb6Projects } from './vb6ProjectLocator';
import {
    diagnosticMetadataForCode,
    diagnosticSourceForCode,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    type ModuleSymbolKind,
} from './analyzer';
import { analyzeVbaModuleSource } from './vbaModuleAnalysis';
import { hostTokenForFileName } from './analyzer/host/hostRegistry';
import {
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
    type VbaProjectAnalysisOptions,
} from './vbaProjectAnalysis';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import type { AnalysisWorkerClient } from './analysisWorkerClient';
import {
    isAnalysisRuleTracked,
} from './analysisSettingsCore';
import {
    effectiveWorkbookAnalysisSettings,
    type EffectiveWorkbookAnalysisSettings,
} from './workbookAnalysisSettings';
import { startPerformanceTrace } from './performanceTrace';
import { isWorkbookSettingsError, settingsPathForWorkbook } from './workbookSettings';
import {
    validateXlideGlobalSettingsFromConfig,
    xlideDiagnosticsEnabledFromConfig,
    type XlideGlobalSettingsProblem,
} from './globalSettings';
import { errorMessage } from './util/errors';
import { visibleDiagnosticsForActiveLine } from './vbaActiveLineDiagnostics';
import {
    XLIDE_DIAGNOSTIC_DATA,
    type XlideDiagnosticWithData,
} from './xlideDiagnosticData';

function workbookContextKey(xlsmPath: string): string {
    return workbookIdentityKey(path.resolve(xlsmPath));
}

const DIAGNOSTIC_OPEN_LOCAL_DELAY_MS = 25;
const DIAGNOSTIC_OPEN_FULL_DELAY_MS = 150;
const DIAGNOSTIC_EDIT_LOCAL_DELAY_MS = 90;
const DIAGNOSTIC_EDIT_FULL_DELAY_MS = 450;
// Above this size the edit-time full pass backs off proportionally (see
// editScheduleDelaysFor), capped so diagnostics never lag more than 2s. Only
// applies when analysis runs in-host: the worker path has no thread contention
// to pace around, so it keeps the flat fast cadence.
const DIAGNOSTIC_LARGE_MODULE_LINES = 8000;
const DIAGNOSTIC_EDIT_FULL_DELAY_MAX_MS = 2000;
// While the workbook has not yet reported a module's kind (backend starting on
// first open), kind-sensitive analysis is deferred and retried instead of
// guessing 'standard' - which floods class modules with bogus Me/Friend errors.
const FULL_PASS_METADATA_RETRY_MS = 750;
const FULL_PASS_METADATA_RETRY_MAX = 40;
const DIAGNOSTIC_ANALYSIS_SETTINGS_CACHE_TTL_MS = 2_000;

type DiagnosticPassKind = 'local' | 'full';

interface DiagnosticScheduleDelays {
    /** Undefined skips the local pass entirely (a scheduled full pass covers it). */
    localDelayMs?: number;
    fullDelayMs?: number;
}

/**
 * Owns the per-document debounce timers and generation counters for the
 * local/full diagnostic passes. Scheduling bumps the document's generation
 * (invalidating in-flight runs) and re-arms both pass timers; a completed
 * full pass suppresses publishing a stale local pass of the same generation.
 */
/**
 * Edit-time pass delays, scaled to document size. The full analysis pass costs
 * roughly linear time in module size (about 1s at ~24k lines even after the
 * analyzer optimizations), so on very large modules it is paced further behind
 * the typing burst instead of contending with it on every 450ms pause. The
 * cheap local pass keeps its fast cadence regardless, so structural squiggles
 * stay responsive. document.lineCount is O(1).
 */
export function editScheduleDelaysFor(
    document: vscode.TextDocument,
    offThreadHealthy: boolean,
): DiagnosticScheduleDelays {
    if (offThreadHealthy) {
        // Passes run on the worker thread: no extension-host contention, so
        // even huge modules keep the flat fast cadence.
        return { localDelayMs: DIAGNOSTIC_EDIT_LOCAL_DELAY_MS, fullDelayMs: DIAGNOSTIC_EDIT_FULL_DELAY_MS };
    }
    const lines = document.lineCount;
    const fullDelayMs = lines > DIAGNOSTIC_LARGE_MODULE_LINES
        ? Math.min(
            DIAGNOSTIC_EDIT_FULL_DELAY_MAX_MS,
            DIAGNOSTIC_EDIT_FULL_DELAY_MS + Math.floor((lines - DIAGNOSTIC_LARGE_MODULE_LINES) / 8),
        )
        : DIAGNOSTIC_EDIT_FULL_DELAY_MS;
    // The in-host local pass is a full module analysis, not just structure, so
    // on a large module it blocks the extension host for hundreds of ms. With
    // no worker to absorb that: workbook documents drop the local pass (the
    // paced full pass above covers them - running both would block the host
    // twice for one result), and loose .bas documents - which never get a full
    // pass - keep a local pass paced to the same backoff.
    const localDelayMs = lines <= DIAGNOSTIC_LARGE_MODULE_LINES
        ? DIAGNOSTIC_EDIT_LOCAL_DELAY_MS
        : document.uri.scheme === XLIDE_SCHEME
            ? undefined
            : fullDelayMs;
    return { localDelayMs, fullDelayMs };
}

class DiagnosticScheduler {
    private readonly _localTimers = new Map<string, ReturnType<typeof setTimeout>>();
    private readonly _fullTimers = new Map<string, ReturnType<typeof setTimeout>>();
    private readonly _generations = new Map<string, number>();
    private readonly _completedFullGenerations = new Map<string, number>();

    constructor(
        private readonly _runPass: (
            document: vscode.TextDocument,
            generation: number,
            pass: DiagnosticPassKind,
        ) => void,
        private readonly _offThreadHealthy: () => boolean = () => false,
    ) {}

    schedule(
        document: vscode.TextDocument,
        delays?: DiagnosticScheduleDelays,
    ): void {
        if (!isVbaDocument(document)) { return; }
        delays ??= editScheduleDelaysFor(document, this._offThreadHealthy());
        const key = document.uri.toString();
        const generation = this._nextGeneration(key);
        this._clearTimer(this._localTimers, key);
        this._clearTimer(this._fullTimers, key);
        if (delays.localDelayMs !== undefined) {
            this._localTimers.set(key, setTimeout(() => {
                this._localTimers.delete(key);
                this._runPass(document, generation, 'local');
            }, delays.localDelayMs));
        }
        if (document.uri.scheme === XLIDE_SCHEME && delays.fullDelayMs !== undefined) {
            this._fullTimers.set(key, setTimeout(() => {
                this._fullTimers.delete(key);
                this._runPass(document, generation, 'full');
            }, delays.fullDelayMs));
        }
    }

    /** Cancels pending passes and invalidates in-flight runs (document closed). */
    cancel(key: string): void {
        this._clearTimer(this._localTimers, key);
        this._clearTimer(this._fullTimers, key);
        this._nextGeneration(key);
    }

    isCurrentRun(
        document: vscode.TextDocument,
        key: string,
        generation: number,
        documentVersion: number,
    ): boolean {
        return this._generations.get(key) === generation &&
            document.version === documentVersion &&
            vscode.workspace.textDocuments.includes(document);
    }

    /**
     * Whether a current run's diagnostics may be published: a local pass is
     * dropped once the same generation's full pass completed; a full pass
     * records its completion.
     */
    shouldPublish(key: string, generation: number, pass: DiagnosticPassKind): boolean {
        if (pass === 'local' && this._completedFullGenerations.get(key) === generation) {
            return false;
        }
        if (pass === 'full') {
            this._completedFullGenerations.set(key, generation);
        }
        return true;
    }

    private _nextGeneration(key: string): number {
        const next = (this._generations.get(key) ?? 0) + 1;
        this._generations.set(key, next);
        this._completedFullGenerations.delete(key);
        return next;
    }

    private _clearTimer(
        timers: Map<string, ReturnType<typeof setTimeout>>,
        key: string,
    ): void {
        const existing = timers.get(key);
        if (existing) {
            clearTimeout(existing);
            timers.delete(key);
        }
    }
}

/**
 * Lazily creates one FileSystemWatcher per workbook settings sidecar the
 * first time a module of that workbook is analyzed, notifies the engine when
 * the sidecar changes, and prunes watchers for workbooks that no longer have
 * open XLIDE documents.
 */
class WorkbookSettingsWatcherRegistry implements vscode.Disposable {
    private readonly _watchers = new Map<string, vscode.Disposable[]>();

    constructor(
        private readonly _onSettingsChanged: (workbookPath: string) => void,
    ) {}

    ensure(workbookPath: string): void {
        const key = workbookContextKey(workbookPath);
        if (this._watchers.has(key)) {
            return;
        }
        const settingsPath = settingsPathForWorkbook(workbookPath);
        const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(
            path.dirname(settingsPath),
            path.basename(settingsPath),
        ));
        const rerun = () => this._onSettingsChanged(workbookPath);
        this._watchers.set(key, [
            watcher.onDidCreate(rerun),
            watcher.onDidChange(rerun),
            watcher.onDidDelete(rerun),
            watcher,
        ]);
    }

    prune(): void {
        const openWorkbookKeys = new Set<string>();
        for (const document of vscode.workspace.textDocuments) {
            const location = moduleLocationOfDocument(document);
            if (location) {
                openWorkbookKeys.add(workbookContextKey(location.xlsmPath));
            }
        }
        for (const [key, disposables] of this._watchers) {
            if (openWorkbookKeys.has(key)) {
                continue;
            }
            disposables.forEach((disposable) => disposable.dispose());
            this._watchers.delete(key);
        }
    }

    dispose(): void {
        for (const disposables of this._watchers.values()) {
            disposables.forEach((disposable) => disposable.dispose());
        }
        this._watchers.clear();
    }
}

/**
 * Live diagnostics: structural block-balance (analyzeVbaStructure) plus the analyzer's
 * high-confidence semantic rules (analyzeModule) - unterminated strings,
 * duplicate procedures/declarations, assignment to a constant, and a
 * configurable Option Explicit reminder. Runs on open and (debounced) on every
 * edit so problems surface while typing, the way a real IDE does. No save and
 * no workbook round-trip required - everything is computed from the editor text.
 */
export function registerVbaDiagnostics(
    context: vscode.ExtensionContext,
    projectIndexService: VbaProjectIndexService,
    workerClient?: AnalysisWorkerClient,
): void {
    const collection = vscode.languages.createDiagnosticCollection('vba');
    const scheduler = new DiagnosticScheduler(
        (document, generation, pass) => runPass(document, generation, pass),
        () => workerClient?.available === true,
    );
    // Last-known workbook metadata per document, learned from successful full
    // passes. Local passes reuse it so they never analyze a class module under
    // a guessed 'standard' kind (the Me/Friend false-error flood); until the
    // kind is known, passes publish nothing and the deferred full pass leads.
    const moduleMetaByDoc = new Map<string, {
        moduleType?: string;
        moduleKind?: ModuleSymbolKind;
        documentType?: EventHandlerDocumentType;
    }>();
    const fullPassMetadataRetries = new Map<string, number>();
    const settingsWatchers = new WorkbookSettingsWatcherRegistry((workbookPath) => {
        invalidateAnalysisSettingsForWorkbook(workbookPath);
        rerunWorkbookDocuments(workbookPath);
    });
    const analysisSettingsCache = new Map<string, {
        loadedAt: number;
        promise: Promise<EffectiveWorkbookAnalysisSettings>;
    }>();

    const workbookKey = (workbookPath: string): string => workbookContextKey(workbookPath);

    const severityToVscode = (s: RuleSeverity): vscode.DiagnosticSeverity => {
        switch (s) {
            case 'error': return vscode.DiagnosticSeverity.Error;
            case 'warning': return vscode.DiagnosticSeverity.Warning;
            case 'information': return vscode.DiagnosticSeverity.Information;
        }
    };

    // Hold "still-typing" syntax errors (e.g. an `If` with no `Then` yet) on the
    // line the cursor is on, matching the VBE which only validates a line once you
    // leave it. Diagnostics are cached unfiltered so the suppression can be
    // re-applied on cursor moves / editor switches without re-running analysis.
    const heldWhileTyping = new WeakSet<vscode.Diagnostic>();
    const lastDiagnostics = new Map<string, {
        uri: vscode.Uri;
        version: number;
        diagnostics: vscode.Diagnostic[];
    }>();
    const lastSuppressedLine = new Map<string, number | undefined>();

    const activeSuppressedLine = (uri: vscode.Uri): number | undefined => {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.toString() !== uri.toString()) {
            return undefined;
        }
        return editor.selection.active.line;
    };

    const filterForActiveLine = (
        uri: vscode.Uri,
        diagnostics: vscode.Diagnostic[],
    ): vscode.Diagnostic[] =>
        visibleDiagnosticsForActiveLine(
            diagnostics,
            activeSuppressedLine(uri),
            (d) => heldWhileTyping.has(d),
        );

    const publish = (uri: vscode.Uri, version: number, diagnostics: vscode.Diagnostic[]): void => {
        lastDiagnostics.set(uri.toString(), { uri, version, diagnostics });
        lastSuppressedLine.set(uri.toString(), activeSuppressedLine(uri));
        collection.set(uri, filterForActiveLine(uri, diagnostics));
    };

    const run = (
        document: vscode.TextDocument,
        delays: DiagnosticScheduleDelays = { localDelayMs: 0, fullDelayMs: 0 },
    ): void => {
        scheduler.schedule(document, delays);
    };

    const runPass = (
        document: vscode.TextDocument,
        generation: number,
        pass: DiagnosticPassKind,
    ): void => {
        const trace = startPerformanceTrace(`liveDiagnostics.${pass}`, document.uri.scheme);
        void runPassAsync(document, generation, pass).then(() => {
            trace.end('ok', document.uri.scheme);
        }, (err) => {
            trace.end('failed', document.uri.scheme);
            if (!isVbaDocument(document)) {
                return;
            }
            const key = document.uri.toString();
            if (!scheduler.isCurrentRun(document, key, generation, document.version)) {
                return;
            }
            // Gate with shouldPublish too: a failed local pass must not overwrite a
            // successful full pass of the same generation that already published.
            if (!scheduler.shouldPublish(key, generation, pass)) {
                return;
            }
            publish(document.uri, document.version, [diagnosticForAnalysisRunError(document, err)]);
        });
    };

    const diagnosticForAnalysisRunError = (
        document: vscode.TextDocument,
        err: unknown,
    ): vscode.Diagnostic => {
        const firstLine = document.lineCount > 0 ? document.lineAt(0).text : '';
        const range = new vscode.Range(0, 0, 0, Math.min(firstLine.length, 1));
        const settingsError = isWorkbookSettingsError(err);
        const message = settingsError
            ? `${err.message} Fix or delete the workbook settings sidecar.`
            : `XLIDE diagnostics failed: ${errorMessage(err)}`;
        const diagnostic = new vscode.Diagnostic(
            range,
            message,
            vscode.DiagnosticSeverity.Error,
        );
        diagnostic.source = settingsError ? 'XLIDE/settings' : 'XLIDE';
        diagnostic.code = settingsError ? 'workbook-settings-invalid' : 'diagnostics-failed';
        return diagnostic;
    };

    const diagnosticsForGlobalSettingsProblems = (
        document: vscode.TextDocument,
        problems: readonly XlideGlobalSettingsProblem[],
    ): vscode.Diagnostic[] => {
        const firstLine = document.lineCount > 0 ? document.lineAt(0).text : '';
        const range = new vscode.Range(0, 0, 0, Math.min(firstLine.length, 1));
        return problems.map((problem) => {
            const diagnostic = new vscode.Diagnostic(
                range,
                `${problem.message} XLIDE is using a safe default; update it in VS Code settings.`,
                vscode.DiagnosticSeverity.Warning,
            );
            diagnostic.source = 'XLIDE/settings';
            diagnostic.code = 'global-setting-invalid';
            return diagnostic;
        });
    };

    const invalidateAnalysisSettingsForWorkbook = (workbookPath: string | undefined): void => {
        if (!workbookPath) {
            analysisSettingsCache.clear();
            return;
        }
        analysisSettingsCache.delete(workbookKey(workbookPath));
    };

    const analysisSettingsForDiagnostics = (
        workbookPath: string | undefined,
    ): Promise<EffectiveWorkbookAnalysisSettings> => {
        if (!workbookPath) {
            return effectiveWorkbookAnalysisSettings(undefined);
        }
        const key = workbookKey(workbookPath);
        const cached = analysisSettingsCache.get(key);
        if (cached && Date.now() - cached.loadedAt < DIAGNOSTIC_ANALYSIS_SETTINGS_CACHE_TTL_MS) {
            return cached.promise;
        }
        const promise = effectiveWorkbookAnalysisSettings(workbookPath).catch((err) => {
            if (analysisSettingsCache.get(key)?.promise === promise) {
                analysisSettingsCache.delete(key);
            }
            throw err;
        });
        analysisSettingsCache.set(key, { loadedAt: Date.now(), promise });
        return promise;
    };

    const runPassAsync = async (
        document: vscode.TextDocument,
        generation: number,
        pass: DiagnosticPassKind,
    ): Promise<void> => {
        if (!isVbaDocument(document)) { return; }
        const key = document.uri.toString();
        const documentVersion = document.version;
        const config = vscode.workspace.getConfiguration('xlide');
        const settingsDiagnostics = diagnosticsForGlobalSettingsProblems(
            document,
            validateXlideGlobalSettingsFromConfig(config),
        );
        if (!xlideDiagnosticsEnabledFromConfig(config).value) {
            if (!scheduler.isCurrentRun(document, key, generation, documentVersion)) {
                return;
            }
            // Route through publish()/cache cleanup so the held-diagnostics cache
            // stays in lockstep with the collection; otherwise a later cursor move
            // or editor switch would re-publish the now-disabled diagnostics.
            if (settingsDiagnostics.length > 0) {
                publish(document.uri, documentVersion, settingsDiagnostics);
            } else {
                collection.delete(document.uri);
                lastDiagnostics.delete(key);
                lastSuppressedLine.delete(key);
            }
            return;
        }
        const text = analysisSourceForDocument(document);

        const moduleName = moduleNameFromDocument(document);
        let workbookPath: string | undefined;
        let moduleType: string | undefined;
        let moduleKind: ModuleSymbolKind | undefined;
        let documentType: EventHandlerDocumentType | undefined;
        let projectOptions: VbaProjectAnalysisOptions = {};
        let workbookRecord: Awaited<ReturnType<VbaProjectIndexService['contextForWorkbook']>> | undefined;
        // A loose document (a .bas on disk no project claims) has no metadata
        // source and keeps the historical standard-kind analysis. A workbook
        // module and a VB6 project's file both have a project to ask.
        const location = moduleLocationOfDocument(document);
        let moduleMetadataKnown = location === undefined;
        if (location) {
            try {
                const xlsmPath = location.xlsmPath;
                workbookPath = xlsmPath;
                settingsWatchers.ensure(xlsmPath);
                if (pass === 'full' || workerClient?.available === true) {
                    const diagnosticProject = await projectIndexService.contextForWorkbook(xlsmPath);
                    workbookRecord = diagnosticProject;
                    const current = diagnosticProject.moduleMetadata.get(moduleIdentityKey(moduleName));
                    if (current) {
                        moduleMetadataKnown = true;
                        moduleType = current.moduleType;
                        moduleKind = current.moduleKind;
                        documentType = current.documentType;
                        moduleMetaByDoc.set(key, { moduleType, moduleKind, documentType });
                    }
                    const project = diagnosticProject.project;
                    diagnosticProject.projectProcedures ??= projectProcedureSignatures(project);
                    projectOptions = projectAnalysisOptionsForModule(
                        project,
                        moduleName,
                        diagnosticProject.projectProcedures,
                    );
                } else {
                    const cached = moduleMetaByDoc.get(key);
                    if (cached) {
                        moduleMetadataKnown = true;
                        moduleType = cached.moduleType;
                        moduleKind = cached.moduleKind;
                        documentType = cached.documentType;
                    }
                }
            } catch {
                projectOptions = {};
            }
        }
        if (!moduleMetadataKnown) {
            // The workbook has not reported this module's kind yet (typically
            // the workbook has not been indexed yet on first open). Analyzing
            // with a guessed 'standard' kind floods class modules with bogus
            // "'Me' is only valid in a class..." errors until a later pass
            // corrects them - publish nothing and retry the full pass instead.
            if (pass === 'full' && scheduler.isCurrentRun(document, key, generation, documentVersion)) {
                const retries = fullPassMetadataRetries.get(key) ?? 0;
                if (retries < FULL_PASS_METADATA_RETRY_MAX) {
                    fullPassMetadataRetries.set(key, retries + 1);
                    const timer = setTimeout(() => { run(document); }, FULL_PASS_METADATA_RETRY_MS);
                    (timer as unknown as { unref?: () => void }).unref?.();
                }
            }
            return;
        }
        fullPassMetadataRetries.delete(key);

        const analysisSettings = await analysisSettingsForDiagnostics(workbookPath);
        const activeEditor = vscode.window.activeTextEditor;
        const activeIncompleteExpressionOffset = activeEditor?.document === document
            ? document.offsetAt(activeEditor.selection.active)
            : undefined;

        // Both passes run on the analysis worker thread when it is healthy, so
        // a large module's pass never blocks the extension host. In-host, the
        // local pass on a ~24k-line class costs ~700ms on the extension-host
        // thread 90ms after every typing pause - the worker path is off-thread
        // and keeps per-document incremental state, so the follow-up full pass
        // of the same generation re-analyzes only what changed. Any failure
        // falls through to the identical in-host pass below.
        if (workerClient?.available && workbookPath && workbookRecord) {
            try {
                const record = workbookRecord;
                const wbKey = workbookKey(workbookPath);
                const crossGeneration = record.crossModuleGeneration(moduleName);
                workerClient.ensureSeeded(wbKey, crossGeneration, () => record.modules.map((m) => ({
                    moduleName: m.moduleName,
                    source: m.source,
                    type: m.type,
                    documentType: m.documentType,
                    implicitMembers: m.implicitMembers,
                    predeclaredId: m.predeclaredId,
                })));
                const workerResult = await workerClient.analyze({
                    docKey: key,
                    workbookKey: wbKey,
                    generation: crossGeneration,
                    source: text,
                    moduleName,
                    moduleType,
                    moduleKind,
                    documentType,
                    severityOverrides: analysisSettings.ruleSeverityOverrides,
                    activeIncompleteExpressionOffset,
                    host: workbookPath ? hostTokenForFileName(workbookPath) : undefined,
                });
                const diagnostics = diagnosticsFromModuleAnalysis(
                    document,
                    workerResult,
                    analysisSettings.untrackedRules,
                    settingsDiagnostics,
                );
                publishDiagnosticsIfCurrent(document, key, generation, documentVersion, pass, diagnostics);
                return;
            } catch {
                // Worker unavailable or died mid-request: fall through to the
                // in-host pass, which produces identical results.
            }
        }

        const moduleAnalysis = analyzeVbaModuleSource({
            source: text,
            moduleName,
            moduleType,
            moduleKind,
            documentType,
            severityOverrides: analysisSettings.ruleSeverityOverrides,
            ...projectOptions,
            activeIncompleteExpressionOffset,
            host: workbookPath ? hostTokenForFileName(workbookPath) : undefined,
        });
        const diagnostics = diagnosticsFromModuleAnalysis(
            document,
            moduleAnalysis,
            analysisSettings.untrackedRules,
            settingsDiagnostics,
        );
        publishDiagnosticsIfCurrent(document, key, generation, documentVersion, pass, diagnostics);
    };

    const diagnosticsFromModuleAnalysis = (
        document: vscode.TextDocument,
        moduleAnalysis: Pick<ReturnType<typeof analyzeVbaModuleSource>, 'diagnostics'>,
        untrackedRules: readonly string[],
        settingsDiagnostics: readonly vscode.Diagnostic[],
    ): vscode.Diagnostic[] => {
        const diagnostics: vscode.Diagnostic[] = [...settingsDiagnostics];
        for (const d of moduleAnalysis.diagnostics) {
            if (!isAnalysisRuleTracked(d.code, untrackedRules)) {
                continue;
            }
            const diag = new vscode.Diagnostic(
                new vscode.Range(
                    document.positionAt(d.span.start),
                    document.positionAt(d.span.end),
                ),
                d.message,
                severityToVscode(d.severity),
            );
            diag.source = diagnosticSourceForCode(d.code);
            if (d.code) {
                diag.code = d.code;
            }
            if (d.data) {
                (diag as XlideDiagnosticWithData)[XLIDE_DIAGNOSTIC_DATA] = d.data;
            }
            // Tag syntax-category findings so they can be held on the active line
            // (e.g. an `If` with no `Then` yet) until the cursor leaves the line.
            if (d.code && diagnosticMetadataForCode(d.code)?.category === 'syntax') {
                heldWhileTyping.add(diag);
            }
            diagnostics.push(diag);
        }
        return diagnostics;
    };

    const publishDiagnosticsIfCurrent = (
        document: vscode.TextDocument,
        key: string,
        generation: number,
        documentVersion: number,
        pass: DiagnosticPassKind,
        diagnostics: vscode.Diagnostic[],
    ): void => {
        if (!scheduler.isCurrentRun(document, key, generation, documentVersion)) {
            return;
        }
        if (!scheduler.shouldPublish(key, generation, pass)) {
            return;
        }
        publish(document.uri, documentVersion, diagnostics);
    };

    const rerunWorkbookDocuments = (workbookPath: string): void => {
        const key = workbookKey(workbookPath);
        for (const document of vscode.workspace.textDocuments) {
            const location = moduleLocationOfDocument(document);
            if (location && workbookKey(location.xlsmPath) === key) {
                run(document);
            }
        }
    };

    context.subscriptions.push(
        collection,
        // A VB6 project map that just loaded or changed can move a file from
        // "loose module" to "member of a project", which changes its
        // analysis: rerun every open file on disk.
        onDidChangeVb6Projects(() => {
            for (const document of vscode.workspace.textDocuments) {
                if (document.uri.scheme === 'file' && isVbaDocument(document)) {
                    scheduler.schedule(document);
                }
            }
        }),
        vscode.workspace.onDidOpenTextDocument((document) => scheduler.schedule(document, {
            localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
            fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
        })),
        vscode.workspace.onDidChangeTextDocument((e) => scheduler.schedule(e.document)),
        vscode.window.onDidChangeActiveTextEditor(() => {
            const editor = vscode.window.activeTextEditor;
            // Re-apply current-line suppression for every cached document: the doc
            // we left should reveal its held diagnostics, and the doc we entered
            // should hide them on its cursor line.
            for (const entry of lastDiagnostics.values()) {
                const key = entry.uri.toString();
                const line = activeSuppressedLine(entry.uri);
                // Skip docs whose suppressed line did not change - i.e. every doc
                // that is neither the editor we left nor the one we entered - so we
                // avoid a needless Array.filter + collection.set re-ingest per doc.
                if (lastSuppressedLine.get(key) === line) {
                    continue;
                }
                lastSuppressedLine.set(key, line);
                collection.set(entry.uri, filterForActiveLine(entry.uri, entry.diagnostics));
            }
            if (editor) {
                scheduler.schedule(editor.document, {
                    localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
                    fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
                });
            }
        }),
        vscode.window.onDidChangeTextEditorSelection((e) => {
            const doc = e.textEditor.document;
            const entry = lastDiagnostics.get(doc.uri.toString());
            // Re-apply suppression when the cursor line changes. Skip when an edit is
            // pending (version drift) - the debounced re-run will republish with
            // correct positions - or when the active line is unchanged.
            if (!entry || entry.version !== doc.version) {
                return;
            }
            const line = activeSuppressedLine(doc.uri);
            if (lastSuppressedLine.get(doc.uri.toString()) === line) {
                return;
            }
            lastSuppressedLine.set(doc.uri.toString(), line);
            collection.set(doc.uri, filterForActiveLine(doc.uri, entry.diagnostics));
        }),
        vscode.workspace.onDidCloseTextDocument((doc) => {
            scheduler.cancel(doc.uri.toString());
            collection.delete(doc.uri);
            lastDiagnostics.delete(doc.uri.toString());
            lastSuppressedLine.delete(doc.uri.toString());
            moduleMetaByDoc.delete(doc.uri.toString());
            fullPassMetadataRetries.delete(doc.uri.toString());
            workerClient?.forget(doc.uri.toString());
            settingsWatchers.prune();
        }),
        vscode.workspace.onDidChangeConfiguration((e) => {
            if (e.affectsConfiguration('xlide.diagnostics') ||
                e.affectsConfiguration('xlide.analysis')) {
                invalidateAnalysisSettingsForWorkbook(undefined);
                vscode.workspace.textDocuments.forEach((document) => run(document));
            }
        }),
        settingsWatchers,
    );
    vscode.workspace.textDocuments.forEach((document) => scheduler.schedule(document, {
        localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
        fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
    }));
}
