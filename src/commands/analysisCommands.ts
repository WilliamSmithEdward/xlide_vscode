import * as vscode from 'vscode';
import * as path from 'path';
import {
    encodeModuleUri,
    decodeModuleUri,
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
} from '../xlideFileSystem';
import { applyOpenDocumentSources } from '../vbaOpenDocuments';
import {
    analyzeProject,
    summarizeProjectAnalysisProblems,
    projectProblemsForModule,
    type ProjectAnalysisProblem,
    type ProjectAnalysisResult,
} from '../vbaProjectWideAnalysis';
import {
    openProjectAnalysisResults,
    type ProjectAnalysisSuppressScope,
} from '../projectAnalysisWebview';
import { analyzeVbaModuleSource } from '../vbaModuleAnalysis';
import { hostTokenForFileName } from '../analyzer/host/hostRegistry';
import { effectiveProjectAnalysisSettings } from '../projectAnalysisSettings';
import { lineStartOffsets } from '../vbaSourceScan';
import type { VbaSymbolIndex } from '../vbaSymbolIndex';
import { moduleKindFromType } from '../vbaNavigation';
import {
    buildLiveVbaProjectIndexAsync,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
} from '../vbaProjectAnalysis';
import { resolveDiagnosticCodeActions } from '../analyzer';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../projectExplorer';
import { suppressionTargetForProblem } from '../vbaAnalysisSuppression';
import { errorMessage } from '../util/errors';
import {
    resolveProjectPath,
    showAnalysisSourceDocument,
    statusMessage,
    type CommandDeps,
} from './shared';

function copilotAnalysisPrompt(
    filePath: string,
    problem: ProjectAnalysisProblem,
    source: string,
): string {
    const lines = source.split(/\r\n|\r|\n/);
    const zeroBasedLine = Math.max(0, problem.line - 1);
    const start = Math.max(0, zeroBasedLine - 6);
    const end = Math.min(lines.length, zeroBasedLine + 7);
    const excerpt = lines
        .slice(start, end)
        .map((line, index) => `${String(start + index + 1).padStart(4, ' ')}: ${line}`)
        .join('\n');
    const rule = problem.code
        ? `${problem.code}${problem.ruleTitle ? ` (${problem.ruleTitle})` : ''}`
        : problem.ruleTitle ?? 'unknown rule';
    return [
        'Please help me understand and fix this Excel VBA analysis finding from XLIDE.',
        '',
        `Workbook: ${path.basename(filePath)}`,
        `Module: ${problem.moduleName} (${problem.moduleType})`,
        `Location: ${problem.line}:${problem.column}`,
        `Severity: ${problem.severity}`,
        `Rule: ${rule}`,
        `Evidence: ${problem.diagnosticKind ?? 'unknown'}`,
        `Message: ${problem.message}`,
        '',
        'Relevant VBA source:',
        '```vba',
        excerpt,
        '```',
        '',
        'Please explain whether the finding is valid, what VBA rule or behavior is involved, and the smallest code change that would fix it.',
    ].join('\n');
}

// Single live-module analysis pipeline shared by the analyze-current-module
// command and the support bundle so the analyzeVbaModuleSource inputs
// cannot drift between the two surfaces.
export async function analyzeOpenModule(
    vbaIndex: VbaSymbolIndex,
    projectPath: string,
    moduleName: string,
    source: string,
) {
    const modules = applyOpenDocumentSources(
        await vbaIndex.getAllModules(projectPath),
        projectPath,
    );
    const current = modules.find(
        (mod) => mod.moduleName.toLowerCase() === moduleName.toLowerCase(),
    );
    const moduleKind = moduleKindFromType(current?.type);
    const moduleType = current?.type ?? 'standard';
    const project = await buildLiveVbaProjectIndexAsync(modules, {
        moduleName,
        moduleKind,
        source,
    });
    const projectOptions = projectAnalysisOptionsForModule(
        project,
        moduleName,
        projectProcedureSignatures(project),
    );
    const analysisSettings = await effectiveProjectAnalysisSettings(projectPath);
    const result = analyzeVbaModuleSource({
        source,
        moduleName,
        moduleType,
        moduleKind,
        documentType: current?.documentType,
        severityOverrides: analysisSettings.ruleSeverityOverrides,
        ...projectOptions,
        host: hostTokenForFileName(projectPath),
    });
    return { modules, current, moduleType, result };
}

export function registerAnalysisCommands(deps: CommandDeps): vscode.Disposable[] {
    const { context, bridge, explorer, out, vbaIndex } = deps;

    function log(msg: string): void {
        out.appendLine(msg);
    }

    let analysisSourceOpenQueue: Promise<void> = Promise.resolve();
    let analysisSourceOpenSequence = 0;

    async function openProjectAnalysisProblem(
        filePath: string,
        problem: ProjectAnalysisProblem,
        _analysisPanelColumn?: vscode.ViewColumn,
    ): Promise<void> {
        await queueAnalysisSourceOpen(() => openProjectAnalysisProblemNow(
            filePath,
            problem,
        ));
    }

    async function openProjectAnalysisProblemNow(
        filePath: string,
        problem: ProjectAnalysisProblem,
    ): Promise<void> {
        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
        const editor = await showAnalysisSourceDocument(doc);
        const line = Math.max(0, problem.line - 1);
        const startColumn = Math.max(0, problem.column - 1);
        const endColumn = Math.max(startColumn + 1, problem.endColumn - 1);
        const start = new vscode.Position(line, startColumn);
        const end = new vscode.Position(line, endColumn);
        const range = new vscode.Range(start, end);
        editor.selection = new vscode.Selection(start, end);
        editor.revealRange(range, vscode.TextEditorRevealType.InCenter);
    }

    function analysisSourceDrifted(doc: vscode.TextDocument, problem: ProjectAnalysisProblem): boolean {
        if (problem.documentVersion !== undefined && doc.version !== problem.documentVersion) {
            vscode.window.showWarningMessage(
                'XLIDE: The module changed since it was analyzed, so this action could land on the wrong location. Re-run analysis, then try again.',
            );
            return true;
        }
        return false;
    }

    async function suppressProjectAnalysisProblem(
        filePath: string,
        problem: ProjectAnalysisProblem,
        scope: ProjectAnalysisSuppressScope,
        _analysisPanelColumn?: vscode.ViewColumn,
    ): Promise<void> {
        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
        if (analysisSourceDrifted(doc, problem)) {
            return;
        }
        const source = doc.getText();
        const starts = lineStartOffsets(source);
        const problemOffset = Math.max(
            0,
            (starts[Math.max(0, problem.line - 1)] ?? 0) + Math.max(0, problem.column - 1),
        );
        if (!problem.suppressionScopes.includes(scope)) {
            throw new Error(`Ignore ${scope} is not valid for '${problem.code ?? 'this analysis finding'}'.`);
        }
        const target = suppressionTargetForProblem(source, starts, problemOffset, scope);
        const code = (problem.code ?? 'all').trim() || 'all';
        const eol = doc.eol === vscode.EndOfLine.CRLF ? '\r\n' : '\n';
        const edit = new vscode.WorkspaceEdit();

        if (target.kind === 'block') {
            edit.insert(uri, new vscode.Position(target.endLine + 1, 0), `' @xlide-analysis-enable-block ${code}${eol}`);
            edit.insert(uri, new vscode.Position(target.startLine, 0), `' @xlide-analysis-disable-block ${code}${eol}`);
        } else if (target.kind === 'member') {
            edit.insert(uri, new vscode.Position(target.startLine, 0), `' @xlide-analysis-disable-next-member ${code}${eol}`);
        } else {
            edit.insert(uri, new vscode.Position(target.startLine, 0), `' @xlide-analysis-disable-file ${code}${eol}`);
        }

        const applied = await vscode.workspace.applyEdit(edit);
        if (!applied) {
            throw new Error('VS Code rejected the analysis ignore edit.');
        }
        const editor = await showAnalysisSourceDocument(doc);
        const position = new vscode.Position(target.startLine, 0);
        editor.selection = new vscode.Selection(position, position);
        editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
        // The directive is visible in the editor at the revealed cursor position.
        statusMessage(`XLIDE: Added ${scope} analysis ignore directive for '${code}'.`);
    }

    async function askCopilotAboutProjectAnalysisProblem(
        filePath: string,
        problem: ProjectAnalysisProblem,
        analysisPanelColumn?: vscode.ViewColumn,
    ): Promise<void> {
        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
        await openProjectAnalysisProblem(filePath, problem, analysisPanelColumn);
        const prompt = copilotAnalysisPrompt(filePath, problem, doc.getText());
        try {
            await vscode.commands.executeCommand('workbench.action.chat.open', { query: prompt });
        } catch {
            await vscode.env.clipboard.writeText(prompt);
            vscode.window.showWarningMessage(
                'XLIDE: Could not open VS Code Chat. The Copilot prompt was copied to the clipboard.',
            );
        }
    }

    async function quickFixProjectAnalysisProblem(
        filePath: string,
        problem: ProjectAnalysisProblem,
        _analysisPanelColumn?: vscode.ViewColumn,
        fixIndex = 0,
    ): Promise<boolean> {
        if (!problem.code) {
            return false;
        }

        const uri = encodeModuleUri(filePath, problem.moduleName);
        const doc = await vscode.workspace.openTextDocument(uri);
        await vscode.languages.setTextDocumentLanguage(doc, XLIDE_VBA_LANGUAGE_ID);
        if (analysisSourceDrifted(doc, problem)) {
            return false;
        }
        const source = doc.getText();
        const starts = lineStartOffsets(source);
        const lineStart = starts[Math.max(0, problem.line - 1)] ?? 0;
        const span = {
            start: lineStart + Math.max(0, problem.column - 1),
            end: lineStart + Math.max(0, problem.endColumn - 1),
        };
        const fixes = resolveDiagnosticCodeActions(source, {
            code: problem.code,
            message: problem.message,
            span,
            expectedClose: problem.expectedClose,
            insertLine: problem.insertLine,
            expectedCloseReplacementSpan: problem.expectedCloseReplacementSpan,
            expectedCloseReplacementText: problem.expectedCloseReplacementText,
            data: problem.data,
            includeSuppressionAction: false,
        });
        const fix = fixes[fixIndex] ?? fixes[0];
        if (!fix || fix.edits.length === 0) {
            return false;
        }

        const edit = new vscode.WorkspaceEdit();
        for (const textEdit of fix.edits) {
            edit.replace(
                uri,
                new vscode.Range(
                    doc.positionAt(textEdit.span.start),
                    doc.positionAt(textEdit.span.end),
                ),
                textEdit.newText,
            );
        }
        const applied = await vscode.workspace.applyEdit(edit);
        if (!applied) {
            return false;
        }

        const editor = await showAnalysisSourceDocument(doc);
        const firstEdit = fix.edits[0];
        const position = doc.positionAt(firstEdit.span.start);
        editor.selection = new vscode.Selection(position, position);
        editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenterIfOutsideViewport);
        return true;
    }

    function showProjectAnalysisResults(
        result: ProjectAnalysisResult,
        onRefreshResult?: () => Promise<ProjectAnalysisResult>,
    ): void {
        const filePath = result.filePath;
        openProjectAnalysisResults(context, result, {
            onOpenProblem: (problem, analysisPanelColumn) =>
                openProjectAnalysisProblem(filePath, problem, analysisPanelColumn),
            onQuickFixProblem: (problem, analysisPanelColumn, fixIndex) =>
                quickFixProjectAnalysisProblem(filePath, problem, analysisPanelColumn, fixIndex),
            onSuppressProblem: (problem, scope, analysisPanelColumn) =>
                suppressProjectAnalysisProblem(filePath, problem, scope, analysisPanelColumn),
            onAskCopilot: (problem, analysisPanelColumn) =>
                askCopilotAboutProjectAnalysisProblem(filePath, problem, analysisPanelColumn),
            onRefreshResult,
        });
    }

    async function queueAnalysisSourceOpen(operation: () => Promise<void>): Promise<void> {
        const sequence = ++analysisSourceOpenSequence;
        const run = analysisSourceOpenQueue
            .catch(() => undefined)
            .then(async () => {
                if (sequence !== analysisSourceOpenSequence) {
                    return;
                }
                await operation();
            });
        analysisSourceOpenQueue = run;
        await run;
    }

    async function currentModuleAnalysisResult(document: vscode.TextDocument): Promise<ProjectAnalysisResult> {
        const { projectPath, moduleName } = decodeModuleUri(document.uri);
        const source = document.getText();
        // Capture the version alongside the source snapshot the problems are
        // computed from, so mutating actions can detect later edits (drift).
        const documentVersion = document.version;
        const { moduleType, result } = await analyzeOpenModule(vbaIndex, projectPath, moduleName, source);
        const byPosition = (a: ProjectAnalysisProblem, b: ProjectAnalysisProblem) => {
            if (a.line !== b.line) { return a.line - b.line; }
            return a.column - b.column;
        };
        const stampVersion = (p: ProjectAnalysisProblem): ProjectAnalysisProblem => ({ ...p, documentVersion });
        const problems = projectProblemsForModule(
            moduleName,
            moduleType,
            source,
            result.diagnostics,
        ).sort(byPosition).map(stampVersion);
        const suppressedProblems = projectProblemsForModule(
            moduleName,
            moduleType,
            source,
            result.suppressedDiagnostics,
            { suppressed: true },
        ).sort(byPosition).map(stampVersion);
        const errorCount = problems.filter((p) => p.severity === 'error').length;
        const warningCount = problems.filter((p) => p.severity === 'warning').length;
        const summary = summarizeProjectAnalysisProblems(problems, suppressedProblems.length);
        return {
            filePath: projectPath,
            moduleCount: 1,
            problems,
            suppressedProblems,
            errorCount,
            warningCount,
            summary,
        };
    }

    async function analyzeActiveModule(): Promise<void> {
        const editor = vscode.window.activeTextEditor;
        if (!editor || editor.document.uri.scheme !== XLIDE_SCHEME) {
            vscode.window.showWarningMessage('XLIDE: Open a VBA module from a macro-enabled file to analyze the current module.');
            return;
        }

        const documentUri = editor.document.uri;
        const { moduleName } = decodeModuleUri(documentUri);
        const analysisResult = await currentModuleAnalysisResult(editor.document);
        const { problems, errorCount, warningCount } = analysisResult;

        // Re-resolve the document fresh on each refresh (rather than capturing the
        // possibly-disposed editor) so the long-lived analysis panel re-analyzes the
        // current content and survives the original editor being closed.
        showProjectAnalysisResults(analysisResult, async () =>
            currentModuleAnalysisResult(await vscode.workspace.openTextDocument(documentUri)));
        // The results panel just opened with these exact counts; a popup toast
        // on top of it is pure noise. A transient status-bar line suffices.
        statusMessage(problems.length === 0
            ? `XLIDE: "${moduleName}" passed analysis (no unsuppressed problems).`
            : `XLIDE: "${moduleName}" has ${errorCount} error(s) and ${warningCount} warning(s).`);
    }

    return [
        // Validate the project's VBA project structure
        registerXlideCommand('xlide.validateProject', async (node: XlideNode) => {
            const filePath = resolveProjectPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No file selected to validate.');
                return;
            }
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Validating "${name}"...`, cancellable: false },
                async () => {
                    const res = await bridge.call<{ issues: string[] }>('validateProject', { path: filePath });
                    const issues = res.issues ?? [];
                    if (issues.length === 0) {
                        log(`[validate] "${name}": no issues`);
                        statusMessage(`XLIDE: "${name}" passed validation (no issues).`);
                        return;
                    }
                    log(`[validate] "${name}": ${issues.length} issue(s):`);
                    for (const issue of issues) {
                        log(`[validate]   - ${issue}`);
                    }
                    void vscode.window.showWarningMessage(
                        `XLIDE: "${name}" has ${issues.length} validation issue(s). See XLIDE Output for details.`,
                    );
                },
            );
        }, { errorPrefix: 'Validation failed', logTag: 'validate', log }),

        // Analyze the active VBA module using the same source text the editor shows.
        registerXlideCommand('xlide.analyzeCurrentModule', () => analyzeActiveModule(), {
            errorPrefix: 'Failed to analyze current module',
            logTag: 'analyzeCurrentModule',
            log,
        }),

        // Analyze every VBA module in the project and show a navigable results panel.
        registerXlideCommand('xlide.analyzeProject', async (node: XlideNode) => {
            const filePath = resolveProjectPath(node);
            if (!filePath) {
                vscode.window.showWarningMessage('XLIDE: No file selected to analyze.');
                return;
            }
            const name = path.basename(filePath);
            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `XLIDE: Analyzing "${name}"...`, cancellable: true },
                async (progress, token) => {
                    try {
                        const result = await analyzeProject(bridge, filePath, {
                            token,
                            progress: (message) => progress.report({ message }),
                        });
                        showProjectAnalysisResults(result, () => analyzeProject(bridge, filePath, {
                            progress: (message) => progress.report({ message }),
                        }));
                        // The results panel just opened with these exact counts;
                        // no popup toast on top of it.
                        statusMessage(result.problems.length === 0
                            ? `XLIDE: "${name}" passed analysis (no problems across ${result.moduleCount} module(s)).`
                            : `XLIDE: "${name}" has ${result.errorCount} error(s) and ${result.warningCount} warning(s).`);
                    } catch (err) {
                        const msg = errorMessage(err);
                        if (err instanceof vscode.CancellationError) {
                            log(`[analyzeProject] Canceled: ${name}`);
                            return;
                        }
                        log(`[analyzeProject] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE: Analysis failed: ${msg}`);
                    }
                },
            );
        }),
    ];
}
