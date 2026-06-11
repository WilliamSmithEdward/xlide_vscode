import * as vscode from 'vscode';
import * as path from 'path';
import { PythonBridge } from './pythonBridge';
import {
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    decodeModuleUri,
    encodeModuleUri,
    isVbaDocument,
    moduleIdentityKey,
    workbookIdentityKey,
} from './xlideFileSystem';
import {
    isStandaloneVbaDocument,
    liveProjectIndexForDocument,
    moduleKindFromDocument,
    moduleNameFromDocument,
} from './vbaDocumentIdentity';
import {
    VbaDefinitionProvider,
    VbaDocumentSymbolProvider,
    VbaReferenceProvider,
    VbaRenameProvider,
    VbaWorkspaceSymbolProvider,
} from './vbaNavigationProviders';
import {
    TYPE_TOKEN_LEGEND,
    VbaTypeSemanticTokensProvider,
} from './vbaSemanticTokensProvider';
import { VbaSymbolIndex, VbaModuleSymbols } from './vbaSymbolIndex';
import {
    findIdentifierOccurrences,
    lineStartOffsets,
    stripVba,
    VBA_IDENTIFIER_NAME_RE,
    VBA_IDENTIFIER_RE,
} from './vbaSourceScan';
import { analyzeVbaStructure } from './vbaStructuralDiagnostics';
import {
    detectSmartBlockOpener,
    isSmartBlockClosedAhead,
    procedureHeaderParensEdit,
    resolveLoopIteratorSyncEdit,
    smartBlockInsertion,
    withMemberContinuationText,
} from './vbaSmartEnter';
import {
    diagnosticSourceForCode,
    DiagnosticSeverity as RuleSeverity,
    EventHandlerDocumentType,
    eventHandlerDocumentTypeForContext,
    isXlideDiagnosticSource,
    normalizeDiagnosticCode,
    ProjectIndex,
    ReferenceScope,
    resolveDiagnosticCodeActions,
    resolveMemberDefinitionsAt,
    resolveProcedureLabelDefinitionAt,
    tokenizeCached,
    type MemberCompletionContext,
    resolveTypeReferenceAt,
    resolveTypeSemanticTokens,
    TypeSemanticTokenType,
    type VbaDiagnosticData,
    type VbaProjectClassMember,
    type VbaProjectClassMemberDefinition,
    type ModuleSymbolKind,
    type Span,
    VbaSymbol as AstSymbol,
    type VbaSymbolKind,
} from './analyzer';
import { analyzeVbaModuleSource } from './vbaModuleAnalysis';
import { registerVbaMemberCompletion } from './vbaMemberCompletion';
import { DocMetadataLoader } from './vbaDocMetadata';
import {
    createOffsetToPositionConverter,
    moduleKindFromType,
    offsetToPosition,
    projectTypeDefinitionToLocation,
    typeDefinitionsForReference,
    typeReferenceLocations,
} from './vbaNavigation';
import {
    buildLiveVbaProjectIndexAsync,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
    type VbaProjectAnalysisOptions,
} from './vbaProjectAnalysis';
import { VbaProjectIndexService } from './vbaProjectIndexService';
import {
    documentOutlineSymbolsForSource,
    workspaceSymbols as presentedWorkspaceSymbols,
    type VbaPresentedSymbol,
    type VbaPresentedWorkspaceSymbol,
} from './vbaSymbolPresentation';
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
    xlideEditorBlockLayoutFromConfig,
    type XlideGlobalSettingsProblem,
} from './globalSettings';
import { errorMessage } from './util/errors';
import {
    XLIDE_DIAGNOSTIC_DATA,
    type XlideDiagnosticWithData,
} from './xlideDiagnosticData';

const VBA_SELECTOR: vscode.DocumentSelector = [
    { scheme: XLIDE_SCHEME, language: 'vba' },
    { scheme: XLIDE_SCHEME, language: XLIDE_VBA_LANGUAGE_ID },
    { scheme: XLIDE_SCHEME },
    { language: 'vba' },
    { language: XLIDE_VBA_LANGUAGE_ID },
];
const XLIDE_SOURCE_ACTION_KIND = vscode.CodeActionKind.Source.append('xlide');
const XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND = XLIDE_SOURCE_ACTION_KIND.append('analyzeCurrentModule');
const XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND = XLIDE_SOURCE_ACTION_KIND.append('exportCurrentModule');

function workbookContextKey(xlsmPath: string): string {
    return workbookIdentityKey(path.resolve(xlsmPath));
}

const DIAGNOSTIC_OPEN_LOCAL_DELAY_MS = 25;
const DIAGNOSTIC_OPEN_FULL_DELAY_MS = 150;
const DIAGNOSTIC_EDIT_LOCAL_DELAY_MS = 90;
const DIAGNOSTIC_EDIT_FULL_DELAY_MS = 450;
const DIAGNOSTIC_ANALYSIS_SETTINGS_CACHE_TTL_MS = 2_000;

/**
 * Live diagnostics: structural block-balance (analyzeVbaStructure) plus the analyzer's
 * high-confidence semantic rules (analyzeModule) - unterminated strings,
 * duplicate procedures/declarations, assignment to a constant, and a
 * configurable Option Explicit reminder. Runs on open and (debounced) on every
 * edit so problems surface while typing, the way a real IDE does. No save and
 * no Python round-trip required - everything is computed from the editor text.
 */
function registerVbaDiagnostics(
    context: vscode.ExtensionContext,
    projectIndexService: VbaProjectIndexService,
): void {
    const collection = vscode.languages.createDiagnosticCollection('vba');
    const localTimers = new Map<string, ReturnType<typeof setTimeout>>();
    const fullTimers = new Map<string, ReturnType<typeof setTimeout>>();
    const diagnosticGenerations = new Map<string, number>();
    const completedFullGenerations = new Map<string, number>();
    const workbookSettingsWatchers = new Map<string, vscode.Disposable[]>();
    const analysisSettingsCache = new Map<string, {
        loadedAt: number;
        promise: Promise<EffectiveWorkbookAnalysisSettings>;
    }>();

    type DiagnosticPassKind = 'local' | 'full';

    interface DiagnosticScheduleDelays {
        localDelayMs: number;
        fullDelayMs?: number;
    }

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
        schedule(document, delays);
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
            if (!isCurrentDiagnosticRun(document, key, generation, document.version)) {
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
            if (!isCurrentDiagnosticRun(document, key, generation, documentVersion)) {
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
                ensureWorkbookSettingsWatcher(xlsmPath);
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
        if (!isCurrentDiagnosticRun(document, key, generation, documentVersion)) {
            return;
        }
        if (pass === 'local' && completedFullGenerations.get(key) === generation) {
            return;
        }
        if (pass === 'full') {
            completedFullGenerations.set(key, generation);
        }
        collection.set(document.uri, diagnostics);
    };

    const isCurrentDiagnosticRun = (
        document: vscode.TextDocument,
        key: string,
        generation: number,
        documentVersion: number,
    ): boolean => {
        return diagnosticGenerations.get(key) === generation &&
            document.version === documentVersion &&
            vscode.workspace.textDocuments.includes(document);
    };

    const nextDiagnosticGeneration = (key: string): number => {
        const next = (diagnosticGenerations.get(key) ?? 0) + 1;
        diagnosticGenerations.set(key, next);
        completedFullGenerations.delete(key);
        return next;
    };

    const clearTimer = (
        timers: Map<string, ReturnType<typeof setTimeout>>,
        key: string,
    ): void => {
        const existing = timers.get(key);
        if (existing) {
            clearTimeout(existing);
            timers.delete(key);
        }
    };

    const schedule = (
        document: vscode.TextDocument,
        delays: DiagnosticScheduleDelays = {
            localDelayMs: DIAGNOSTIC_EDIT_LOCAL_DELAY_MS,
            fullDelayMs: DIAGNOSTIC_EDIT_FULL_DELAY_MS,
        },
    ): void => {
        if (!isVbaDocument(document)) { return; }
        const key = document.uri.toString();
        const generation = nextDiagnosticGeneration(key);
        clearTimer(localTimers, key);
        clearTimer(fullTimers, key);
        localTimers.set(key, setTimeout(() => {
            localTimers.delete(key);
            runPass(document, generation, 'local');
        }, delays.localDelayMs));
        if (document.uri.scheme === XLIDE_SCHEME && delays.fullDelayMs !== undefined) {
            fullTimers.set(key, setTimeout(() => {
                fullTimers.delete(key);
                runPass(document, generation, 'full');
            }, delays.fullDelayMs));
        }
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
    const ensureWorkbookSettingsWatcher = (workbookPath: string): void => {
        const key = workbookKey(workbookPath);
        if (workbookSettingsWatchers.has(key)) {
            return;
        }
        const settingsPath = settingsPathForWorkbook(workbookPath);
        const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(
            path.dirname(settingsPath),
            path.basename(settingsPath),
        ));
        const rerun = () => {
            invalidateAnalysisSettingsForWorkbook(workbookPath);
            rerunWorkbookDocuments(workbookPath);
        };
        workbookSettingsWatchers.set(key, [
            watcher.onDidCreate(rerun),
            watcher.onDidChange(rerun),
            watcher.onDidDelete(rerun),
            watcher,
        ]);
    };
    const pruneWorkbookSettingsWatchers = (): void => {
        const openWorkbookKeys = new Set<string>();
        for (const document of vscode.workspace.textDocuments) {
            if (document.uri.scheme !== XLIDE_SCHEME) {
                continue;
            }
            try {
                openWorkbookKeys.add(workbookKey(decodeModuleUri(document.uri).xlsmPath));
            } catch {
                // Ignore invalid XLIDE URIs.
            }
        }
        for (const [key, disposables] of workbookSettingsWatchers) {
            if (openWorkbookKeys.has(key)) {
                continue;
            }
            disposables.forEach((disposable) => disposable.dispose());
            workbookSettingsWatchers.delete(key);
        }
    };
    const disposeWorkbookSettingsWatchers = (): void => {
        for (const disposables of workbookSettingsWatchers.values()) {
            disposables.forEach((disposable) => disposable.dispose());
        }
        workbookSettingsWatchers.clear();
    };

    context.subscriptions.push(
        collection,
        vscode.workspace.onDidOpenTextDocument((document) => schedule(document, {
            localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
            fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
        })),
        vscode.workspace.onDidChangeTextDocument((e) => schedule(e.document)),
        vscode.window.onDidChangeActiveTextEditor(() => {
            const editor = vscode.window.activeTextEditor;
            if (editor) {
                schedule(editor.document, {
                    localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
                    fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
                });
            }
        }),
        vscode.workspace.onDidCloseTextDocument((doc) => {
            const key = doc.uri.toString();
            clearTimer(localTimers, key);
            clearTimer(fullTimers, key);
            nextDiagnosticGeneration(key);
            collection.delete(doc.uri);
            pruneWorkbookSettingsWatchers();
        }),
        vscode.workspace.onDidChangeConfiguration((e) => {
            if (e.affectsConfiguration('xlide.diagnostics') ||
                e.affectsConfiguration('xlide.analysis')) {
                invalidateAnalysisSettingsForWorkbook(undefined);
                vscode.workspace.textDocuments.forEach((document) => run(document));
            }
        }),
        { dispose: disposeWorkbookSettingsWatchers },
    );
    vscode.workspace.textDocuments.forEach((document) => schedule(document, {
        localDelayMs: DIAGNOSTIC_OPEN_LOCAL_DELAY_MS,
        fullDelayMs: DIAGNOSTIC_OPEN_FULL_DELAY_MS,
    }));
}

class VbaCodeActionProvider implements vscode.CodeActionProvider {
    public provideCodeActions(
        document: vscode.TextDocument,
        _range: vscode.Range,
        context: vscode.CodeActionContext,
        _token: vscode.CancellationToken,
    ): vscode.ProviderResult<vscode.CodeAction[]> {
        if (!isVbaDocument(document)) { return []; }
        const wantsQuickFix = codeActionKindRequested(context.only, vscode.CodeActionKind.QuickFix);
        const wantsAnalyzeCurrentModule = codeActionKindRequested(context.only, XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND);
        const wantsExportCurrentModule = codeActionKindRequested(context.only, XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND);
        if (!wantsQuickFix && !wantsAnalyzeCurrentModule && !wantsExportCurrentModule) {
            return [];
        }

        const source = document.getText();
        const actions: vscode.CodeAction[] = [];
        if (wantsAnalyzeCurrentModule && document.uri.scheme === XLIDE_SCHEME) {
            const action = new vscode.CodeAction(
                'XLIDE: Analyze Current Module',
                XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
            );
            action.command = {
                command: 'xlide.analyzeCurrentModule',
                title: 'Analyze Current Module',
            };
            actions.push(action);
        }
        if (wantsExportCurrentModule && document.uri.scheme === XLIDE_SCHEME && !document.uri.authority) {
            const action = new vscode.CodeAction(
                'XLIDE: Export/Sync Current Module',
                XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
            );
            action.command = {
                command: 'xlide.exportCurrentModuleToFolder',
                title: 'Export/Sync Current Module',
            };
            actions.push(action);
        }
        if (!wantsQuickFix) {
            return actions;
        }
        let structuralDiagnostics: ReturnType<typeof analyzeVbaStructure> | undefined;
        for (const diagnostic of context.diagnostics) {
            if (!isXlideDiagnosticSource(diagnostic.source)) { continue; }
            const code = normalizeDiagnosticCode(diagnostic.code);
            if (!code) { continue; }
            const structuralDiagnostic = code === 'missing-block-closer'
                ? matchingStructuralDiagnostic(
                    diagnostic,
                    structuralDiagnostics ??= analyzeVbaStructure(source),
                )
                : undefined;
            const fixes = resolveDiagnosticCodeActions(source, {
                code,
                message: diagnostic.message,
                expectedClose: structuralDiagnostic?.expectedClose,
                insertLine: structuralDiagnostic?.insertLine,
                expectedCloseReplacementSpan: structuralDiagnostic?.expectedCloseReplacement
                    ? {
                        start: document.offsetAt(new vscode.Position(
                            structuralDiagnostic.expectedCloseReplacement.line,
                            structuralDiagnostic.expectedCloseReplacement.startCol,
                        )),
                        end: document.offsetAt(new vscode.Position(
                            structuralDiagnostic.expectedCloseReplacement.line,
                            structuralDiagnostic.expectedCloseReplacement.endCol,
                        )),
                    }
                    : undefined,
                expectedCloseReplacementText: structuralDiagnostic?.expectedCloseReplacement?.text,
                span: {
                    start: document.offsetAt(diagnostic.range.start),
                    end: document.offsetAt(diagnostic.range.end),
                },
                includeSuppressionAction: true,
                data: (diagnostic as XlideDiagnosticWithData)[XLIDE_DIAGNOSTIC_DATA],
            });
            for (const fix of fixes) {
                const action = new vscode.CodeAction(fix.title, vscode.CodeActionKind.QuickFix);
                action.diagnostics = [diagnostic];
                action.isPreferred = fix.isPreferred;
                const edit = new vscode.WorkspaceEdit();
                for (const textEdit of fix.edits) {
                    edit.replace(
                        document.uri,
                        new vscode.Range(
                            document.positionAt(textEdit.span.start),
                            document.positionAt(textEdit.span.end),
                        ),
                        textEdit.newText,
                    );
                }
                action.edit = edit;
                actions.push(action);
            }
        }
        return actions;
    }
}

function codeActionKindRequested(
    only: vscode.CodeActionKind | undefined,
    kind: vscode.CodeActionKind,
): boolean {
    return !only || only.contains(kind) || kind.contains(only);
}

function matchingStructuralDiagnostic(
    diagnostic: vscode.Diagnostic,
    problems: ReturnType<typeof analyzeVbaStructure>,
): ReturnType<typeof analyzeVbaStructure>[number] | undefined {
    return problems.find((problem) =>
        problem.code === normalizeDiagnosticCode(diagnostic.code) &&
        problem.message === diagnostic.message &&
        diagnostic.range.start.isEqual(new vscode.Position(problem.line, problem.startCol)) &&
        diagnostic.range.end.isEqual(new vscode.Position(problem.line, problem.endCol)),
    );
}

/**
 * VBA-IDE-style smart Enter: typing a block opener and pressing Enter
 * auto-inserts the matching closer below, leaving the cursor on the indented
 * body line. `With` also seeds the body line with `.` so member completion can
 * start immediately.
 */
function registerVbaAutoBlock(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        // React only to a plain Enter (newline plus optional auto-indent),
        // never to pastes or multi-character insertions.
        if (!/^\r?\n[ \t]*$/.test(change.text)) { return; }

        const openerLineIndex = change.range.start.line;
        const openerLine = doc.lineAt(openerLineIndex).text;
        const headerParensEdit = procedureHeaderParensEdit(openerLine);
        const normalizedOpenerLine = headerParensEdit
            ? `${openerLine.slice(0, headerParensEdit.startCol)}${headerParensEdit.newText}${openerLine.slice(headerParensEdit.endCol)}`
            : openerLine;
        const opener = detectSmartBlockOpener(stripVba(normalizedOpenerLine));
        if (!opener) {
            await maybeContinueWithMemberLine(doc, openerLineIndex);
            return;
        }

        const bodyLineIndex = openerLineIndex + 1;
        if (bodyLineIndex >= doc.lineCount) { return; }

        const strippedLines = doc.getText().split(/\r\n|\r|\n/).map(stripVba);
        strippedLines[openerLineIndex] = stripVba(normalizedOpenerLine);
        const closedAhead = isSmartBlockClosedAhead(strippedLines, openerLineIndex, opener);

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
        const bodyLine = doc.lineAt(bodyLineIndex).text;
        if (!/^[ \t]*$/.test(bodyLine)) { return; }
        const smartBlock = smartBlockInsertion(normalizedOpenerLine, bodyLine, opener, {
            eol,
            insertCloser: !closedAhead,
            layout: xlideEditorBlockLayoutFromConfig(vscode.workspace.getConfiguration('xlide')).value,
        });
        const bodyRange = new vscode.Range(
            new vscode.Position(bodyLineIndex, 0),
            new vscode.Position(bodyLineIndex, bodyLine.length),
        );

        applying = true;
        try {
            await editor.edit(
                (eb) => {
                    if (headerParensEdit) {
                        eb.insert(
                            new vscode.Position(openerLineIndex, headerParensEdit.startCol),
                            headerParensEdit.newText,
                        );
                    }
                    eb.replace(
                        bodyRange,
                        smartBlock.replacementText,
                    );
                },
                { undoStopBefore: false, undoStopAfter: true },
            );
        } finally {
            applying = false;
        }

        // Keep the caret on the indented body line, above the inserted End. The
        // delayed pass wins same-Enter listener races such as canonical casing.
        const placeCaret = (): void => {
            if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
                return;
            }
            const caretLineIndex = bodyLineIndex + smartBlock.bodyLineOffset;
            if (caretLineIndex >= doc.lineCount || doc.lineAt(caretLineIndex).text !== smartBlock.bodyText) {
                return;
            }
            const caret = new vscode.Position(
                caretLineIndex,
                smartBlock.bodyText.length,
            );
            editor.selection = new vscode.Selection(caret, caret);
        };
        placeCaret();
        setTimeout(placeCaret, 0);
    });

    context.subscriptions.push(sub);
}

async function maybeContinueWithMemberLine(
    doc: vscode.TextDocument,
    previousLineIndex: number,
): Promise<void> {
    const bodyLineIndex = previousLineIndex + 1;
    if (bodyLineIndex >= doc.lineCount) { return; }

    const bodyLine = doc.lineAt(bodyLineIndex).text;
    if (!/^[ \t]*$/.test(bodyLine)) { return; }

    const lineText = withMemberContinuationText(doc.getText(), previousLineIndex);
    if (!lineText) { return; }

    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document !== doc) { return; }

    const bodyRange = new vscode.Range(
        new vscode.Position(bodyLineIndex, 0),
        new vscode.Position(bodyLineIndex, bodyLine.length),
    );
    const applied = await editor.edit(
        (eb) => eb.replace(bodyRange, lineText),
        { undoStopBefore: false, undoStopAfter: true },
    );
    if (!applied) { return; }

    const placeCaret = (): void => {
        if (vscode.window.activeTextEditor !== editor || editor.document !== doc) {
            return;
        }
        if (bodyLineIndex >= doc.lineCount || doc.lineAt(bodyLineIndex).text !== lineText) {
            return;
        }
        const caret = new vscode.Position(bodyLineIndex, lineText.length);
        editor.selection = new vscode.Selection(caret, caret);
    };
    placeCaret();
    setTimeout(placeCaret, 0);
}

/**
 * Keeps simple loop iterator names paired across `For`/`For Each` and `Next`.
 * This intentionally lives outside snippets so hand-written loops get the same
 * behavior as completed loops.
 */
function registerVbaLoopIteratorSync(context: vscode.ExtensionContext): void {
    let applying = false;

    const sub = vscode.workspace.onDidChangeTextDocument(async (e) => {
        if (applying) { return; }
        const doc = e.document;
        if (!isVbaDocument(doc)) { return; }
        if (e.contentChanges.length !== 1) { return; }

        const change = e.contentChanges[0];
        if (/[\r\n]/.test(change.text)) { return; }

        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document !== doc) { return; }

        const lineIndex = Math.min(change.range.start.line, doc.lineCount - 1);
        const lineLength = doc.lineAt(lineIndex).text.length;
        const character = Math.min(lineLength, change.range.start.character + change.text.length);
        const offset = doc.offsetAt(new vscode.Position(lineIndex, character));
        const syncEdit = resolveLoopIteratorSyncEdit(doc.getText(), offset);
        if (!syncEdit) { return; }

        applying = true;
        try {
            await editor.edit(
                (eb) => eb.replace(
                    new vscode.Range(
                        doc.positionAt(syncEdit.span.start),
                        doc.positionAt(syncEdit.span.end),
                    ),
                    syncEdit.newText,
                ),
                { undoStopBefore: false, undoStopAfter: false },
            );
        } finally {
            applying = false;
        }
    });

    context.subscriptions.push(sub);
}

export function registerVbaLanguageProviders(
    context: vscode.ExtensionContext,
    bridge: PythonBridge,
): VbaSymbolIndex {
    const index = new VbaSymbolIndex(bridge);
    const projectIndexService = new VbaProjectIndexService(index);

    registerVbaDiagnostics(context, projectIndexService);
    registerVbaAutoBlock(context);
    registerVbaLoopIteratorSync(context);
    const docMetadata = new DocMetadataLoader();
    void docMetadata.start(context);
    registerVbaMemberCompletion(context, projectIndexService, VBA_SELECTOR, docMetadata.registry);

    context.subscriptions.push(
        index,
        projectIndexService,
        vscode.languages.registerDocumentSymbolProvider(
            VBA_SELECTOR,
            new VbaDocumentSymbolProvider(),
            { label: 'XLIDE VBA' },
        ),
        vscode.languages.registerWorkspaceSymbolProvider(
            new VbaWorkspaceSymbolProvider(projectIndexService),
        ),
        vscode.languages.registerDefinitionProvider(
            VBA_SELECTOR,
            new VbaDefinitionProvider(projectIndexService),
        ),
        vscode.languages.registerReferenceProvider(
            VBA_SELECTOR,
            new VbaReferenceProvider(projectIndexService),
        ),
        vscode.languages.registerRenameProvider(
            VBA_SELECTOR,
            new VbaRenameProvider(projectIndexService),
        ),
        vscode.languages.registerCodeActionsProvider(
            VBA_SELECTOR,
            new VbaCodeActionProvider(),
            {
                providedCodeActionKinds: [
                    vscode.CodeActionKind.QuickFix,
                    XLIDE_ANALYZE_CURRENT_MODULE_ACTION_KIND,
                    XLIDE_EXPORT_CURRENT_MODULE_ACTION_KIND,
                ],
            },
        ),
        vscode.languages.registerDocumentSemanticTokensProvider(
            VBA_SELECTOR,
            new VbaTypeSemanticTokensProvider(projectIndexService),
            TYPE_TOKEN_LEGEND,
        ),
        // Keep the index consistent with saves to virtual VBA documents.
        vscode.workspace.onDidSaveTextDocument((doc) => {
            if (doc.uri.scheme !== XLIDE_SCHEME) { return; }
            try {
                const { xlsmPath, moduleName } = decodeModuleUri(doc.uri);
                index.updateModuleSource(xlsmPath, moduleName, doc.getText());
            } catch {
                // Ignore URIs we cannot decode.
            }
        }),
    );

    return index;
}
