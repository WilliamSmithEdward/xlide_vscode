import * as vscode from 'vscode';
import { checkModuleContentToken, moduleContentToken } from './moduleContentToken';
import type { ProjectAnalysisResult } from './vbaProjectWideAnalysis';
import * as fs from 'fs';
import * as path from 'path';
import { ProjectEngine } from './projectEngine';
import { ProjectExplorer } from './projectExplorer';
import { XlideFileSystemProvider } from './xlideFileSystem';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { containerAppNameForPath } from './macroContainerUi';
import {
    agentWriteDiffsEnabled,
    keepAgentChange,
    onDidChangePendingAgentReviews,
    openAgentReviewDiff,
    presentAgentModuleWrite,
    registerAgentDiffProvider,
    revertAgentChange,
    type AgentWriteReviewDeps,
} from './xlideAgentDiff';
import {
    deleteProjectModule,
    renameProjectModule,
    writeProjectModule,
    type ProjectModuleOperationDeps,
} from './projectModuleOperations';
import {
    exportProjectModules,
} from './moduleExport';
import { type ExportMode } from './projectSettings';
import { setProjectModuleSyncExportMode } from './projectModuleSyncSettings';
import { analyzeProject } from './vbaProjectWideAnalysis';
import { executeVbaTestRun } from './vbaTestRunPipeline';
import { agentVbaTestArtifactPayloadFromPipeline } from './agentVbaTestArtifacts';
import {
    describeVbaTestSelection,
    summarizeVbaTestRun,
    type VbaTestSelectionOptions,
} from './vbaTestRunner';
import { formatChangeSummary, withWriteAudit } from './xlideWriteAudit';
import { runWriteWithHostCoordination } from './officeWriteCoordinator';
import { registerXlideCommand } from './xlideCommandRegistration';
import { gitChangesReport, gitModuleCompareDeps, type GitModuleCompareDeps } from './gitModuleCompare';

// --------------------------------------------------------------------------
// Input types matching the inputSchema in package.json
// --------------------------------------------------------------------------

interface ListModulesInput { filePath: string; }
interface ListSubsInput    { filePath: string; moduleName: string; }
interface SearchModulesInput { filePath: string; query: string; isRegex?: boolean; matchCase?: boolean; maxResults?: number; }
interface ReadModuleInput  { filePath: string; moduleName: string; startLine?: number; endLine?: number; }
interface WriteModuleInput { filePath: string; moduleName: string; source: string; expectedContentToken?: string; kind?: string; }
interface RenameModuleInput { filePath: string; moduleName: string; newName: string; }
interface DeleteModuleInput { filePath: string; moduleName: string; }
interface ListSheetsInput  { filePath: string; }
interface GetProjectInfoInput { filePath: string; }
interface ValidateProjectInput { filePath: string; }
interface AnalyzeProjectInput { filePath: string; moduleName?: string; }
interface RunVbaTestsInput {
    filePath: string;
    moduleName?: string;
    procedureName?: string;
    testIds?: string[];
    includeTags?: string[];
    excludeTags?: string[];
    failFast?: boolean;
    includeHostEvents?: boolean;
}
interface CreateProjectInput { filePath: string; }
interface ReadCellsInput   { filePath: string; sheet: string; range: string; }
interface ReadFormulasInput { filePath: string; sheet: string; range: string; }
interface WriteCellsInput  { filePath: string; sheet: string; startCell: string; data: unknown[][]; }
/** `sheet` is the old name for `surface`, still accepted. */
interface ListShapesInput  { filePath: string; surface?: string; sheet?: string; }
interface EditShapeInput {
    filePath: string;
    surface?: string;
    sheet?: string;
    action: 'add' | 'update' | 'delete';
    name?: string;
    type?: string;
    range?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
    text?: string;
    macro?: string;
    linkedCell?: string;
    inputRange?: string;
    altText?: string;
    newName?: string;
}
interface ExportModulesInput { filePath: string; exportFolder?: string; exportMode?: ExportMode; }
interface ConfigureExportModeInput { filePath: string; exportMode: ExportMode; }
interface GitChangesInput { filePath: string; revision?: string; moduleName?: string; }

function shapeActionTitle(action: string): string {
    return action === 'add' ? 'Add shape' : action === 'delete' ? 'Delete shape' : 'Change shape';
}

function textResult(value: string): vscode.LanguageModelToolResult {
    return new vscode.LanguageModelToolResult([
        new vscode.LanguageModelTextPart(value),
    ]);
}

function vbaTestSelectionFromInput(input: RunVbaTestsInput): VbaTestSelectionOptions | undefined {
    const selection: VbaTestSelectionOptions = {
        moduleName: input.moduleName,
        procedureName: input.procedureName,
        testIds: input.testIds,
        includeTags: input.includeTags,
        excludeTags: input.excludeTags,
    };
    if (
        !selection.moduleName &&
        !selection.procedureName &&
        !selection.testIds?.length &&
        !selection.includeTags?.length &&
        !selection.excludeTags?.length
    ) {
        return undefined;
    }
    return selection;
}

export function registerAgentTools(
    _context: vscode.ExtensionContext,
    bridge: ProjectEngine,
    explorer: ProjectExplorer,
    fsProvider: XlideFileSystemProvider,
    vbaIndex: VbaSymbolIndex,
    gitDeps: GitModuleCompareDeps = gitModuleCompareDeps(bridge),
): vscode.Disposable[] {
    const ops: ProjectModuleOperationDeps = { bridge, explorer, fsProvider, vbaIndex };
    // Revert runs through the same audited write path as the tools, so the
    // audit log shows the user's revert next to the agent write it undoes.
    const agentDiffDeps: AgentWriteReviewDeps = {
        readModuleSource: async (filePath: string, moduleName: string): Promise<string> => {
            const current = await bridge.call<{ source: string }>(
                'readModule', { path: filePath, module: moduleName },
            );
            return current.source;
        },
        writeModuleSource: async (filePath: string, moduleName: string, source: string): Promise<void> => {
            await withWriteAudit({
                command: 'xlide.revertAgentChange',
                operation: 'write-module',
                projectPath: filePath,
                moduleName,
                failedSummary: 'Revert agent change: 0 changed, 1 failed',
            }, async () => ({
                // agentReviewHandled: the revert resolves its own review.
                result: await writeProjectModule(
                    ops,
                    { filePath, moduleName, source },
                    { agentReviewHandled: true },
                ),
                summary: 'Revert agent change: 1 changed',
            }));
        },
        deleteModule: async (filePath: string, moduleName: string): Promise<void> => {
            await withWriteAudit({
                command: 'xlide.revertAgentChange',
                operation: 'delete-module',
                projectPath: filePath,
                moduleName,
                failedSummary: 'Revert agent change: 0 changed, 1 failed',
            }, async () => ({
                result: await deleteProjectModule(ops, { filePath, moduleName }),
                summary: 'Revert agent change: 1 removed',
            }));
        },
    };
    return [
        registerAgentDiffProvider(),
        // The tree marks follow pending reviews: redraw the module that gained
        // or lost one, and the folders it sits under.
        onDidChangePendingAgentReviews((change) => {
            explorer.refreshAgentReviewMarks(change.filePath, change.moduleName);
        }),
        registerXlideCommand('xlide.reviewAgentChange', async (node?: { filePath?: string; moduleName?: string }) => {
            if (!node?.filePath || !node.moduleName) {
                return;
            }
            await openAgentReviewDiff(node.filePath, node.moduleName);
        }),
        registerXlideCommand('xlide.keepAgentChange', (node?: { filePath?: string; moduleName?: string }) => {
            if (!node?.filePath || !node.moduleName) {
                return;
            }
            keepAgentChange(node.filePath, node.moduleName);
        }),
        registerXlideCommand('xlide.revertAgentChange', async (node?: { filePath?: string; moduleName?: string }) => {
            if (!node?.filePath || !node.moduleName) {
                return;
            }
            await revertAgentChange(agentDiffDeps, node.filePath, node.moduleName);
        }),
        // ----------------------------------------------------------------
        // xlide_listProjects
        // ----------------------------------------------------------------
        vscode.lm.registerTool<Record<string, never>>('xlide_listProjects', {
            async invoke(_options, _token) {
                const uris = await findMacroContainerFiles();
                const files = uris.map((u) => u.fsPath);
                return textResult(JSON.stringify(files, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listModules
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListModulesInput>('xlide_listModules', {
            async invoke(options, token) {
                const modules = await bridge.call<Array<{ name: string; type: string }>>(
                    'listModules',
                    { path: options.input.filePath },
                    token,
                );
                return textResult(JSON.stringify(modules, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listSubs
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListSubsInput>('xlide_listSubs', {
            async invoke(options, token) {
                const subs = await bridge.call<Array<{ name: string; kind: string; line: number }>>(
                    'listSubs',
                    { path: options.input.filePath, module: options.input.moduleName },
                    token,
                );
                return textResult(JSON.stringify(subs, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_searchModules
        // ----------------------------------------------------------------
        vscode.lm.registerTool<SearchModulesInput>('xlide_searchModules', {
            async invoke(options, token) {
                const { filePath, query, isRegex, matchCase, maxResults } = options.input;
                const cap = Math.max(1, maxResults ?? 200);
                let matcher: (line: string) => boolean;
                if (isRegex) {
                    let re: RegExp;
                    try {
                        re = new RegExp(query, matchCase ? 'u' : 'iu');
                    } catch (err) {
                        return textResult(`Invalid regular expression: ${(err as Error).message}`);
                    }
                    matcher = (line) => re.test(line);
                } else {
                    const needle = matchCase ? query : query.toLowerCase();
                    matcher = (line) => (matchCase ? line : line.toLowerCase()).includes(needle);
                }

                const modules = await bridge.call<Array<{ name: string }>>(
                    'listModules', { path: filePath }, token,
                );
                const hits: Array<{ moduleName: string; line: number; text: string }> = [];
                let capped = false;
                for (const module of modules) {
                    if (capped) { break; }
                    const read = await bridge.call<{ source: string }>(
                        'readModule', { path: filePath, module: module.name }, token,
                    );
                    const lines = read.source.split(/\r?\n/);
                    for (let i = 0; i < lines.length; i += 1) {
                        if (!matcher(lines[i])) { continue; }
                        if (hits.length >= cap) { capped = true; break; }
                        hits.push({ moduleName: module.name, line: i + 1, text: lines[i].trim() });
                    }
                }
                // Say so when results were dropped: a truncated list that looks
                // complete is worse than no list.
                const note = capped ? `\n(stopped at ${cap} matches; narrow the query or raise maxResults)` : '';
                return textResult(`${JSON.stringify(hits, null, 2)}${note}`);
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readModule
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadModuleInput>('xlide_readModule', {
            async invoke(options, token) {
                const { filePath, moduleName, startLine, endLine } = options.input;
                const result = await bridge.call<{ source: string }>(
                    'readModule',
                    { path: filePath, module: moduleName },
                    token,
                );
                const contentToken = moduleContentToken(result.source);
                const lines = result.source.split(/\r?\n/);
                // A window is over the WHOLE module, so the token still
                // describes what a later conditional write is checked against.
                const from = Math.max(1, startLine ?? 1);
                const to = Math.min(lines.length, endLine ?? lines.length);
                const windowed = startLine === undefined && endLine === undefined;
                const body = windowed ? result.source : lines.slice(from - 1, to).join('\n');
                const header = windowed
                    ? `contentToken: ${contentToken} (${lines.length} lines)`
                    : `contentToken: ${contentToken} (lines ${from}-${to} of ${lines.length})`;
                return textResult(`${header}\n${body}`);
            },
        }),

        // ----------------------------------------------------------------
        // xlide_writeModule  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<WriteModuleInput>('xlide_writeModule', {
            async invoke(options, _token) {
                const { filePath, moduleName, source, expectedContentToken } = options.input;
                // What a create makes. The description offered kind='class'
                // and the tool never read it, so every new module was standard.
                const kind = options.input.kind?.toLowerCase();
                if (kind !== undefined && kind !== 'standard' && kind !== 'class') {
                    return textResult(`kind must be 'standard' or 'class', not '${options.input.kind}'.`);
                }
                if (kind !== undefined) {
                    // Given for a module of the other kind, the code went into that
                    // module as it was: class code in a standard module, where `Me`
                    // does not compile.
                    const existing = (await bridge.call<Array<{ name: string; type: string }>>(
                        'listModules', { path: filePath },
                    )).find((module) => module.name.toLowerCase() === moduleName.toLowerCase());
                    if (existing && (existing.type === 'standard') !== (kind === 'standard')) {
                        return textResult(
                            `"${existing.name}" is already a ${existing.type} module, and kind only chooses what a new `
                            + 'module is. Write it without kind, or use a name no module has.',
                        );
                    }
                }
                // Only chat-driven invocations carry a toolInvocationToken;
                // programmatic calls get no review surface.
                const wantsReview = options.toolInvocationToken !== undefined
                    && agentWriteDiffsEnabled();
                // The before-image feeds both the stale-token check and the
                // keep/revert review; a missing module (a create) reads as ''.
                let beforeSource = '';
                let beforeExisted = false;
                if (expectedContentToken || wantsReview) {
                    try {
                        const current = await bridge.call<{ source: string }>(
                            'readModule', { path: filePath, module: moduleName },
                        );
                        beforeSource = current.source;
                        beforeExisted = true;
                    } catch {
                        // A create: the module does not exist yet.
                    }
                }
                if (expectedContentToken) {
                    const stale = checkModuleContentToken(beforeSource, expectedContentToken, moduleName);
                    if (stale) {
                        return textResult(stale.message);
                    }
                }
                const { summary } = await withWriteAudit({
                    command: 'xlide_writeModule',
                    operation: 'write-module',
                    projectPath: filePath,
                    moduleName,
                    failedSummary: 'Write module: 0 changed, 1 failed',
                }, async () => {
                    // agentReviewHandled only when a review will actually be
                    // presented below; a token-less programmatic write is
                    // tracked like any other out-of-band write.
                    const result = await writeProjectModule(
                        ops,
                        { filePath, moduleName, source, ...(kind !== undefined ? { kind } : {}) },
                        { agentReviewHandled: wantsReview },
                    );
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Write module',
                            changed: [moduleName],
                        }),
                    };
                });
                if (wantsReview) {
                    let afterSource = source;
                    try {
                        afterSource = (await agentDiffDeps.readModuleSource(filePath, moduleName));
                    } catch {
                        // The write succeeded; the review still opens with the
                        // requested source as the after-image.
                    }
                    void presentAgentModuleWrite(filePath, moduleName, {
                        before: beforeSource,
                        beforeExisted,
                        after: afterSource,
                    });
                }
                return textResult(`${summary}\nModule "${moduleName}" written successfully.`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, moduleName } = options.input;
                return {
                    invocationMessage: `Writing VBA module "${moduleName}"`,
                    confirmationMessages: {
                        title: 'Write VBA Module',
                        message: new vscode.MarkdownString(
                            `Write changes to **${moduleName}** in \`${filePath}\`?\n\n` +
                            `This will overwrite the module source and save the project.`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_renameModule  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<RenameModuleInput>('xlide_renameModule', {
            async invoke(options, _token) {
                const { filePath, moduleName, newName } = options.input;
                let renamedTo = newName;
                const { summary } = await withWriteAudit({
                    command: 'xlide_renameModule',
                    operation: 'rename-module',
                    projectPath: filePath,
                    moduleName,
                    failedSummary: 'Rename module: 0 changed, 1 failed',
                }, async () => {
                    const result = await renameProjectModule(ops, { filePath, moduleName, newName });
                    // An Access form's module keeps its prefix: `Form_Customers`.
                    renamedTo = result.moduleName ?? newName;
                    return {
                        result,
                        moduleName: renamedTo,
                        summary: formatChangeSummary({
                            operation: 'Rename module',
                            changed: [`${moduleName} -> ${renamedTo}`],
                        }),
                    };
                });
                return textResult(`${summary}\nModule "${moduleName}" renamed to "${renamedTo}".`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, moduleName, newName } = options.input;
                return {
                    invocationMessage: `Renaming module "${moduleName}" to "${newName}"`,
                    confirmationMessages: {
                        title: 'Rename VBA Module',
                        message: new vscode.MarkdownString(
                            `Rename module **${moduleName}** to **${newName}** in \`${filePath}\`?`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_deleteModule  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<DeleteModuleInput>('xlide_deleteModule', {
            async invoke(options, _token) {
                const { filePath, moduleName } = options.input;
                const { summary } = await withWriteAudit({
                    command: 'xlide_deleteModule',
                    operation: 'delete-module',
                    projectPath: filePath,
                    moduleName,
                    failedSummary: 'Delete module: 0 changed, 1 failed',
                }, async () => {
                    const result = await deleteProjectModule(ops, { filePath, moduleName });
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Delete module',
                            changed: [moduleName],
                        }),
                    };
                });
                return textResult(`${summary}\nModule "${moduleName}" deleted.`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, moduleName } = options.input;
                return {
                    invocationMessage: `Deleting module "${moduleName}"`,
                    confirmationMessages: {
                        title: 'Delete VBA Module',
                        message: new vscode.MarkdownString(
                            `Permanently delete module **${moduleName}** from \`${filePath}\`?\n\n` +
                            `This cannot be undone.`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listSheets
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListSheetsInput>('xlide_listSheets', {
            async invoke(options, token) {
                const result = await bridge.call<{ sheets: Array<{ name: string; dimensions: string }> }>(
                    'listSheets',
                    { path: options.input.filePath },
                    token,
                );
                return textResult(JSON.stringify(result.sheets, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_getProjectInfo
        // ----------------------------------------------------------------
        vscode.lm.registerTool<GetProjectInfoInput>('xlide_getProjectInfo', {
            async invoke(options, token) {
                const result = await bridge.call<{
                    modules: Array<{ name: string; type: string }>;
                    sheets: Array<{ name: string; dimensions: string }>;
                    namedRanges: Array<{ name: string; ref: string }>;
                }>('getProjectInfo', { path: options.input.filePath }, token);
                return textResult(JSON.stringify(result, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_validateProject
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ValidateProjectInput>('xlide_validateProject', {
            async invoke(options, token) {
                const result = await bridge.call<{ issues: string[] }>(
                    'validateProject',
                    { path: options.input.filePath },
                    token,
                );
                return textResult(JSON.stringify(result, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_analyzeProject
        // ----------------------------------------------------------------
        vscode.lm.registerTool<AnalyzeProjectInput>('xlide_analyzeProject', {
            async invoke(options, token) {
                const { filePath, moduleName } = options.input;
                const result = await analyzeProject(bridge, filePath, { token });
                if (!moduleName) {
                    return textResult(JSON.stringify(result, null, 2));
                }
                // Checking one module you just edited should not mean reading
                // back every finding in the project.
                const wanted = moduleName.toLowerCase();
                const forModule = (p: { moduleName: string }) => p.moduleName.toLowerCase() === wanted;
                const problems = result.problems.filter(forModule);
                const scoped: ProjectAnalysisResult = {
                    ...result,
                    moduleCount: 1,
                    problems,
                    suppressedProblems: result.suppressedProblems.filter(forModule),
                    errorCount: problems.filter((p) => p.severity === 'error').length,
                    warningCount: problems.filter((p) => p.severity === 'warning').length,
                };
                return textResult(JSON.stringify(scoped, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_runVbaTests
        // ----------------------------------------------------------------
        vscode.lm.registerTool<RunVbaTestsInput>('xlide_runVbaTests', {
            async invoke(options, _token) {
                const { filePath, failFast, includeHostEvents } = options.input;
                const selection = vbaTestSelectionFromInput(options.input);
                const result = await executeVbaTestRun(bridge, filePath, { selection, failFast });
                if (result.kind === 'blocked-support') {
                    return textResult(JSON.stringify({
                        ok: false,
                        blocked: true,
                        reason: 'test-support',
                        filePath,
                        support: result.support,
                    }, null, 2));
                }
                if (result.kind === 'blocked-com') {
                    return textResult(JSON.stringify({
                        ok: false,
                        blocked: true,
                        reason: 'office-com',
                        filePath,
                        runtime: result.runtime,
                    }, null, 2));
                }
                if (result.kind === 'blocked-busy') {
                    return textResult(JSON.stringify({
                        ok: false,
                        blocked: true,
                        reason: 'run-in-progress',
                        filePath,
                        activeRun: result.activeRunDescription,
                        advice: 'A VBA test run is already executing. Wait for it to finish, then retry.',
                    }, null, 2));
                }

                const { execution } = result;
                const summary = summarizeVbaTestRun(execution.report);
                const ok = summary.failed === 0 &&
                    summary.timeout === 0 &&
                    summary.hostError === 0 &&
                    summary.xpass === 0;
                const artifacts = agentVbaTestArtifactPayloadFromPipeline(result.artifacts);
                return textResult(JSON.stringify({
                    ok,
                    summary,
                    artifacts,
                    report: execution.report,
                    ...(includeHostEvents ? { hostEvents: execution.hostEvents } : {}),
                }, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath, failFast } = options.input;
                const selection = vbaTestSelectionFromInput(options.input);
                const scope = describeVbaTestSelection(selection) || 'all tests';
                return {
                    invocationMessage: `Running XLIDE VBA tests for "${filePath}"`,
                    confirmationMessages: {
                        title: 'Run XLIDE VBA Tests',
                        message: new vscode.MarkdownString(
                            `Run **${scope}** in \`${filePath}\` through the XLIDE owned read-only ${containerAppNameForPath(filePath)} test host?` +
                            `${failFast ? '\n\nFail-fast is enabled.' : ''}`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_createProject
        // ----------------------------------------------------------------
        vscode.lm.registerTool<CreateProjectInput>('xlide_createProject', {
            async invoke(options, _token) {
                const { filePath } = options.input;
                const { result } = await withWriteAudit({
                    command: 'xlide_createProject',
                    operation: 'create-project',
                    projectPath: filePath,
                    failedSummary: 'Create project: 0 changed, 1 failed',
                }, async () => {
                    if (!path.isAbsolute(filePath)) {
                        // A relative path resolves against this process's cwd, which
                        // is not the user's workspace, so the existsSync overwrite guard
                        // below could check a different directory than the write.
                        throw new Error(
                            `xlide_createProject requires an absolute filePath (got "${filePath}").`,
                        );
                    }
                    if (fs.existsSync(filePath)) {
                        throw new Error(
                            `File already exists: "${filePath}". ` +
                            `xlide_createProject does not overwrite existing projects - choose a different filePath.`,
                        );
                    }
                    const result = await bridge.call<{ ok: boolean; path: string }>(
                        'createProject',
                        { path: filePath },
                    );
                    explorer.refresh();
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Create project',
                            changed: [filePath],
                        }),
                    };
                });
                return textResult(JSON.stringify(result, null, 2));
            },
            async prepareInvocation(options, _token) {
                return {
                    invocationMessage: `Creating project "${options.input.filePath}"`,
                    confirmationMessages: {
                        title: 'Create New Macro-Enabled File',
                        message: new vscode.MarkdownString(
                            `Create a new macro-enabled ${containerAppNameForPath(options.input.filePath)} file at \`${options.input.filePath}\`?`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readCells
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadCellsInput>('xlide_readCells', {
            async invoke(options, token) {
                const { filePath, sheet, range } = options.input;
                const result = await bridge.call<{ data: unknown[][] }>(
                    'readCells',
                    { path: filePath, sheet, range },
                    token,
                );
                return textResult(JSON.stringify(result.data, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_readFormulas
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ReadFormulasInput>('xlide_readFormulas', {
            async invoke(options, token) {
                const { filePath, sheet, range } = options.input;
                const result = await bridge.call<{ data: unknown[][] }>(
                    'readFormulas',
                    { path: filePath, sheet, range },
                    token,
                );
                return textResult(JSON.stringify(result.data, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_writeCells  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<WriteCellsInput>('xlide_writeCells', {
            async invoke(options, _token) {
                const { filePath, sheet, startCell, data } = options.input;
                const { summary } = await withWriteAudit({
                    command: 'xlide_writeCells',
                    operation: 'write-cells',
                    projectPath: filePath,
                    failedSummary: 'Write cells: 0 changed, 1 failed',
                }, async () => {
                    // Coordinated like a module write: with the workbook open in
                    // Excel, the lock is handled the way the user's setting says.
                    const result = await runWriteWithHostCoordination(filePath, () => bridge.call('writeCells', {
                        path: filePath,
                        sheet,
                        startCell,
                        data,
                    }));
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: 'Write cells',
                            changed: [`${sheet}!${startCell}`],
                        }),
                    };
                });
                return textResult(`${summary}\nCells written to sheet "${sheet}" starting at "${startCell}".`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, sheet, startCell } = options.input;
                return {
                    invocationMessage: `Writing cells to "${sheet}" in "${filePath}"`,
                    confirmationMessages: {
                        title: 'Write Excel Cells',
                        message: new vscode.MarkdownString(
                            `Write data to sheet **${sheet}** starting at \`${startCell}\` in \`${filePath}\`?`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_listShapes
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListShapesInput>('xlide_listShapes', {
            async invoke(options, token) {
                const { filePath } = options.input;
                const surface = options.input.surface ?? options.input.sheet;
                const result = await bridge.call<{ surfaces: unknown[] }>(
                    'listShapes',
                    { path: filePath, ...(surface ? { surface } : {}) },
                    token,
                );
                return textResult(JSON.stringify(result.surfaces, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_editShape  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<EditShapeInput>('xlide_editShape', {
            async invoke(options, _token) {
                const { filePath, surface: named, sheet, ...edit } = options.input;
                // Word's body is the surface a caller who names none means.
                const surface = named ?? sheet ?? '';
                const { result, summary } = await withWriteAudit({
                    command: 'xlide_editShape',
                    operation: 'edit-shape',
                    projectPath: filePath,
                    failedSummary: `${shapeActionTitle(edit.action)}: 0 changed, 1 failed`,
                }, async () => {
                    const result = await runWriteWithHostCoordination(filePath, () => bridge.call<{ ok: true; name: string }>('editShape', {
                        path: filePath,
                        surface,
                        ...edit,
                    }));
                    const shape = surface ? `${surface}!${result.name}` : result.name;
                    return {
                        result,
                        summary: formatChangeSummary({
                            operation: shapeActionTitle(edit.action),
                            ...(edit.action === 'delete' ? { removed: [shape] } : { changed: [shape] }),
                        }),
                    };
                });
                const done = edit.action === 'add' ? 'added' : edit.action === 'delete' ? 'deleted' : 'changed';
                const where = surface ? ` on "${surface}"` : '';
                return textResult(`${summary}\nShape "${result.name}" ${done}${where} in "${filePath}".`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, action, name, type, range, macro, left, top } = options.input;
                const surface = options.input.surface ?? options.input.sheet;
                const at = range ? `at \`${range}\`` : left !== undefined || top !== undefined ? `at ${left ?? 0}, ${top ?? 0}` : '';
                const what = action === 'add'
                    ? `a ${type ?? 'shape'}${name ? ` named **${name}**` : ''}${at ? ` ${at}` : ''}`
                    : `**${name ?? '?'}**`;
                const link = macro === undefined ? '' : macro ? `, running \`${macro}\` on a click` : ', with no macro';
                const where = surface ? ` on **${surface}**` : '';
                return {
                    invocationMessage: `${shapeActionTitle(action)}${surface ? ` on "${surface}"` : ''} in "${filePath}"`,
                    confirmationMessages: {
                        title: shapeActionTitle(action),
                        message: new vscode.MarkdownString(
                            `${action === 'add' ? 'Add' : action === 'update' ? 'Change' : 'Delete'} ${what}${link}${where} in \`${filePath}\`?`,
                        ),
                    },
                };
            },
        }),
        // ----------------------------------------------------------------
        // xlide_exportModules  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ExportModulesInput>('xlide_exportModules', {
            async invoke(options, _token) {
                const { filePath, exportFolder, exportMode } = options.input;
                const { result, summary } = await withWriteAudit({
                    command: 'xlide_exportModules',
                    operation: 'export-modules',
                    projectPath: filePath,
                    targetPath: exportFolder,
                    failedSummary: 'Export modules: 0 changed, 1 failed',
                }, async () => {
                    const result = await exportProjectModules(bridge, { filePath, exportFolder, exportMode });
                    return {
                        result,
                        targetPath: result.exportFolder,
                        summary: formatChangeSummary({
                            operation: 'Export modules',
                            changed: result.writtenFiles,
                            removed: result.removedFiles,
                        }),
                    };
                });
                return textResult(JSON.stringify({ ...result, changeSummary: summary }, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath, exportFolder, exportMode } = options.input;
                return {
                    invocationMessage: `Exporting VBA modules for "${filePath}"`,
                    confirmationMessages: {
                        title: 'Export VBA Modules',
                        message: new vscode.MarkdownString(
                            `Export all modules for \`${filePath}\` using mode **${exportMode ?? 'exportAll'}**` +
                            `${exportFolder ? ` to folder \`${exportFolder}\`` : ' using configured folder'}` +
                            `?\n\nThis writes files and updates <project>.xlide_settings.json.`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_configureExportMode  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ConfigureExportModeInput>('xlide_configureExportMode', {
            async invoke(options, _token) {
                const { filePath, exportMode } = options.input;
                const updated = await setProjectModuleSyncExportMode(filePath, exportMode);
                return textResult(JSON.stringify({ filePath, ...updated }, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath, exportMode } = options.input;
                return {
                    invocationMessage: `Configuring export mode for "${filePath}"`,
                    confirmationMessages: {
                        title: 'Configure Export Mode',
                        message: new vscode.MarkdownString(
                            `Set export mode for \`${filePath}\` to **${exportMode}**?\n\n` +
                            `This updates <project>.xlide_settings.json beside the project.`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_gitChanges
        // ----------------------------------------------------------------
        // `git diff` on a workbook says "binary files differ"; this reads the
        // committed workbook with the engine and diffs it module by module,
        // so an agent can review a change or write a commit message from it.
        // Refusals (no repository, untracked) come back as a report, so the
        // agent can read the reason instead of catching an error.
        vscode.lm.registerTool<GitChangesInput>('xlide_gitChanges', {
            async invoke(options, _token) {
                const { filePath, revision, moduleName } = options.input;
                const report = await gitChangesReport(gitDeps, filePath, revision || 'HEAD', moduleName);
                return textResult(JSON.stringify(report, null, 2));
            },
        }),
    ];
}
