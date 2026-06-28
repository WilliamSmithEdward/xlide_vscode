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
    decodeModuleUri,
    isVbaDocument,
    moduleIdentityKey,
    workbookIdentityKey,
} from './xlideFileSystem';
import { moduleNameFromDocument } from './vbaDocumentIdentity';
import {
    diagnosticSourceForCode,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    type ModuleSymbolKind,
} from './analyzer';
import { analyzeVbaModuleSource } from './vbaModuleAnalysis';
import {
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
    type VbaProjectAnalysisOptions,
} from './vbaProjectAnalysis';
import { VbaProjectIndexService } from './vbaProjectIndexService';
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
const DIAGNOSTIC_ANALYSIS_SETTINGS_CACHE_TTL_MS = 2_000;

type DiagnosticPassKind = 'local' | 'full';

interface DiagnosticScheduleDelays {
    localDelayMs: number;
    fullDelayMs?: number;
}

/**
 * Owns the per-document debounce timers and generation counters for the
 * local/full diagnostic passes. Scheduling bumps the document's generation
 * (invalidating in-flight runs) and re-arms both pass timers; a completed
 * full pass suppresses publishing a stale local pass of the same generation.
 */
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
    ) {}

    schedule(
        document: vscode.TextDocument,
        delays: DiagnosticScheduleDelays = {
            localDelayMs: DIAGNOSTIC_EDIT_LOCAL_DELAY_MS,
            fullDelayMs: DIAGNOSTIC_EDIT_FULL_DELAY_MS,
        },
    ): void {
        if (!isVbaDocument(document)) { return; }
        const key = document.uri.toString();
        const generation = this._nextGeneration(key);
        this._clearTimer(this._localTimers, key);
        this._clearTimer(this._fullTimers, key);
        this._localTimers.set(key, setTimeout(() => {
            this._localTimers.delete(key);
            this._runPass(document, generation, 'local');
        }, delays.localDelayMs));
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
            if (document.uri.scheme !== XLIDE_SCHEME) {
                continue;
            }
            try {
                openWorkbookKeys.add(workbookContextKey(decodeModuleUri(document.uri).xlsmPath));
            } catch {
                // Ignore invalid XLIDE URIs.
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
 * no Python round-trip required - everything is computed from the editor text.
 */
export function registerVbaDiagnostics(
    context: vscode.ExtensionContext,
    projectIndexService: VbaProjectIndexService,
): void {
    const collection = vscode.languages.createDiagnosticCollection('vba');
    const scheduler = new DiagnosticScheduler(
        (document, generation, pass) => runPass(document, generation, pass),
    );
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
            collection.set(document.uri, [diagnosticForAnalysisRunError(document, err)]);
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
                `${problem.message} Fix the value in VS Code settings.`,
                vscode.DiagnosticSeverity.Error,
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
            if (settingsDiagnostics.length > 0) {
                collection.set(document.uri, settingsDiagnostics);
            } else {
                collection.delete(document.uri);
            }
            return;
        }
        const text = document.getText();

        const moduleName = moduleNameFromDocument(document);
        let workbookPath: string | undefined;
        let moduleType: string | undefined;
        let moduleKind: ModuleSymbolKind | undefined;
        let documentType: EventHandlerDocumentType | undefined;
        let projectOptions: VbaProjectAnalysisOptions = {};
        if (document.uri.scheme === XLIDE_SCHEME) {
            try {
                const { xlsmPath } = decodeModuleUri(document.uri);
                workbookPath = xlsmPath;
                settingsWatchers.ensure(xlsmPath);
                if (pass === 'full') {
                    const diagnosticProject = await projectIndexService.contextForWorkbook(xlsmPath);
                    const current = diagnosticProject.moduleMetadata.get(moduleIdentityKey(moduleName));
                    moduleType = current?.moduleType;
                    moduleKind = current?.moduleKind;
                    documentType = current?.documentType;
                    const project = diagnosticProject.project;
                    diagnosticProject.projectProcedures ??= projectProcedureSignatures(project);
                    projectOptions = projectAnalysisOptionsForModule(
                        project,
                        moduleName,
                        diagnosticProject.projectProcedures,
                    );
                }
            } catch {
                projectOptions = {};
            }
        }

        const analysisSettings = await analysisSettingsForDiagnostics(workbookPath);
        const activeEditor = vscode.window.activeTextEditor;
        const activeIncompleteExpressionOffset = activeEditor?.document === document
            ? document.offsetAt(activeEditor.selection.active)
            : undefined;
        const moduleAnalysis = analyzeVbaModuleSource({
            source: text,
            moduleName,
            moduleType,
            moduleKind,
            documentType,
            severityOverrides: analysisSettings.ruleSeverityOverrides,
            ...projectOptions,
            activeIncompleteExpressionOffset,
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
        moduleAnalysis: ReturnType<typeof analyzeVbaModuleSource>,
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
        collection.set(document.uri, diagnostics);
    };

    const rerunWorkbookDocuments = (workbookPath: string): void => {
        const key = workbookKey(workbookPath);
        for (const document of vscode.workspace.textDocuments) {
            if (document.uri.scheme !== XLIDE_SCHEME) {
                continue;
            }
            try {
                if (workbookKey(decodeModuleUri(document.uri).xlsmPath) === key) {
                    run(document);
                }
            } catch {
                // Ignore URIs that are no longer valid XLIDE module documents.
            }
        }
    };

    context.subscriptions.push(
        collection,
        vscode.workspace.onDidOpenTextDocument((document) => scheduler.schedule(document, {
            localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
            fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
        })),
        vscode.workspace.onDidChangeTextDocument((e) => scheduler.schedule(e.document)),
        vscode.window.onDidChangeActiveTextEditor(() => {
            const editor = vscode.window.activeTextEditor;
            if (editor) {
                scheduler.schedule(editor.document, {
                    localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
                    fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
                });
            }
        }),
        vscode.workspace.onDidCloseTextDocument((doc) => {
            scheduler.cancel(doc.uri.toString());
            collection.delete(doc.uri);
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
