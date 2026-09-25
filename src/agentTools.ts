import * as vscode from 'vscode';
import { checkModuleContentToken, moduleContentToken } from './moduleContentToken';
import type { ProjectAnalysisResult } from './vbaProjectWideAnalysis';
import * as fs from 'fs';
import * as path from 'path';
import { applyModuleEdits, modulePartHeading, readModuleParts, splitModuleLines, type ModuleEdit, type ModuleLineRange } from './moduleParts';
import { applyImportModuleSyncPlan, selectedModuleSyncItems } from './moduleImport';
import { buildImportModuleSyncPlan, type ImportMode, type ModuleSyncPlan } from './moduleSyncPlan';
import { effectiveProjectModuleSyncSettings } from './projectModuleSyncSettings';
import { errorMessage } from './util/errors';
import { ProjectEngine } from './projectEngine';
import { ProjectExplorer } from './projectExplorer';
import { XlideFileSystemProvider } from './xlideFileSystem';
import { VbaSymbolIndex } from './vbaSymbolIndex';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { canAddVbaProjectTo, containerAppNameForPath } from './macroContainerUi';
import {
    agentWriteDiffsEnabled,
    keepAgentChange,
    onDidChangePendingAgentReviews,
    openAgentReviewDiff,
    pendingAgentReviews,
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
interface ReadModuleInput  { filePath: string; moduleName: string; startLine?: number; endLine?: number; ranges?: ModuleLineRange[]; procedures?: string[]; }
interface WriteModuleInput { filePath: string; moduleName: string; source: string; expectedContentToken?: string; kind?: string; }
interface EditModuleInput  { filePath: string; moduleName: string; expectedContentToken?: string; edits?: ModuleEdit[]; }
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
interface ListReferencesInput { filePath: string; }
interface AddReferenceInput { filePath: string; library: string; }
interface RemoveReferenceInput { filePath: string; library: string; }
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
    hidden?: boolean;
    rotation?: number;
    zOrder?: string;
    fill?: Record<string, unknown>;
    line?: Record<string, unknown>;
    font?: Record<string, unknown>;
}
interface ExportModulesInput { filePath: string; exportFolder?: string; exportMode?: ExportMode; }
interface ImportModulesInput { filePath: string; importFolder?: string; importMode?: ImportMode; modules?: string[]; }
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

/**
 * The line a listing tool adds when it found no code, said plainly enough
 * that an agent does not read it as a failure.
 *
 * A bare `[]` is ambiguous in a way that matters: a project with nothing in
 * it will take a new module, and a file with no project at all will not.
 * Neither is an error - a workbook saved as .xlsm before the first macro is
 * written has no VBA project in it, and XLIDE reads it perfectly well - so
 * the answer says which state it is rather than leaving an agent to find out
 * by having a write refused.
 */
async function noCodeNote(bridge: ProjectEngine, filePath: string): Promise<string> {
    const name = path.basename(filePath);
    const app = containerAppNameForPath(filePath);
    let hasVbaProject: boolean;
    try {
        ({ hasVbaProject } = await bridge.call<{ hasVbaProject: boolean }>(
            'hasVbaProject', { path: filePath },
        ));
    } catch {
        return '';
    }
    return hasVbaProject
        ? `\n\n${name} has a VBA project with no modules in it yet. This is not an error; `
            + 'xlide_writeModule adds the first one.'
        : `\n\n${name} has no VBA project in it at all, which is how ${app} saves a macro-enabled `
            + 'file that has never held a macro. This is not an error and there is nothing to '
            + 'retry: XLIDE read the file fine. xlide_writeModule cannot put the first module in. '
            + (canAddVbaProjectTo(filePath)
                ? "Ask the user to add a project: the file's \"No VBA in this file yet\" row in the XLIDE tree "
                    + 'offers Add VBA Project. '
                : `This is a legacy file, so the first macro has to be written in ${app} and saved. `)
            + 'Everything else about the file (sheets, cells, shapes) is readable and writable now.';
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
    /**
     * Writes a module the way the write tools do, and presents the write for
     * review when the call is chat-driven: the before-image the caller read,
     * against what the module holds now. The summary line for the result.
     */
    async function writeModuleForAgent(request: {
        command: 'xlide_writeModule' | 'xlide_editModule';
        operationLabel: string;
        filePath: string;
        moduleName: string;
        source: string;
        kind?: string;
        wantsReview: boolean;
        before: { source: string; existed: boolean };
    }): Promise<string> {
        const { filePath, moduleName, kind } = request;
        const { summary } = await withWriteAudit({
            command: request.command,
            operation: 'write-module',
            projectPath: filePath,
            moduleName,
            failedSummary: `${request.operationLabel}: 0 changed, 1 failed`,
        }, async () => {
            // agentReviewHandled only when a review will actually be
            // presented below; a token-less programmatic write is tracked
            // like any other out-of-band write.
            const result = await writeProjectModule(
                ops,
                { filePath, moduleName, source: request.source, ...(kind !== undefined ? { kind } : {}) },
                { agentReviewHandled: request.wantsReview },
            );
            return {
                result,
                summary: formatChangeSummary({ operation: request.operationLabel, changed: [moduleName] }),
            };
        });
        if (request.wantsReview) {
            let afterSource = request.source;
            try {
                afterSource = await agentDiffDeps.readModuleSource(filePath, moduleName);
            } catch {
                // The write succeeded; the review still opens with the
                // requested source as the after-image.
            }
            void presentAgentModuleWrite(filePath, moduleName, {
                before: request.before.source,
                beforeExisted: request.before.existed,
                after: afterSource,
            });
        }
        return summary;
    }

    /**
     * The plan an import call describes and the items it applies: what the
     * confirmation shows, and what invoke runs. The folder is the call's, or
     * the one the project's settings record from its last export.
     */
    async function planImport(
        input: ImportModulesInput,
    ): Promise<{ plan: ModuleSyncPlan; selectedIds: string[] } | { refused: string }> {
        const { filePath, importMode, modules } = input;
        if (importMode !== undefined && importMode !== 'updateOnly' && importMode !== 'trueUpStandardClass') {
            return { refused: `importMode must be 'updateOnly' or 'trueUpStandardClass', not '${String(importMode)}'.` };
        }
        if (modules !== undefined && (!Array.isArray(modules) || modules.some((name) => typeof name !== 'string'))) {
            return { refused: 'modules must be a list of module names.' };
        }
        const settings = await effectiveProjectModuleSyncSettings(filePath);
        const folder = input.importFolder ?? settings.folderPath;
        if (!folder) {
            return {
                refused: 'No folder to import from: pass importFolder, or export first, which records the folder in '
                    + `${path.basename(settings.settingsPath)}.`,
            };
        }
        if (!path.isAbsolute(folder)) {
            return { refused: 'importFolder must be an absolute path.' };
        }
        let isFolder = false;
        try {
            isFolder = (await fs.promises.stat(folder)).isDirectory();
        } catch {
            // Not there: refused below.
        }
        if (!isFolder) {
            return { refused: `"${folder}" is not a folder.` };
        }
        const plan = await buildImportModuleSyncPlan(bridge, {
            projectPath: filePath,
            importFolder: folder,
            importMode: importMode ?? settings.importMode,
            folderPathSource: input.importFolder ? 'session' : settings.folderPathSource,
            importModeSource: importMode ? 'session' : settings.importModeSource,
            settingsPath: settings.settingsPath,
            // The confirmation counts the items and the apply writes them;
            // nobody looks at a diff.
            withDiffs: false,
        });
        // What the preview would check, plus the files it would list as
        // unchanged or skipped, so the result says what became of every file.
        let items = plan.items.filter((item) =>
            item.checked || item.status === 'unchanged' || item.status === 'skipping-import');
        if (modules?.length) {
            const wanted = new Set(modules.map((name) => name.toLowerCase()));
            const known = new Set(plan.items.map((item) => item.moduleName.toLowerCase()));
            const missing = modules.filter((name) => !known.has(name.toLowerCase()));
            if (missing.length > 0) {
                const files = plan.items.map((item) => item.relativeName).sort().join(', ') || 'no module files';
                return {
                    refused: `No file in ${folder} is for ${missing.map((name) => `"${name}"`).join(', ')}. `
                        + `The folder has: ${files}.`,
                };
            }
            items = plan.items.filter((item) => wanted.has(item.moduleName.toLowerCase()));
        }
        return { plan, selectedIds: items.map((item) => item.id) };
    }

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
        // The buttons above the tree (issue #94): every pending review at
        // once. Revert asks first, since it writes every module back.
        registerXlideCommand('xlide.keepAllAgentChanges', () => {
            for (const { filePath, moduleName } of pendingAgentReviews()) {
                keepAgentChange(filePath, moduleName);
            }
        }),
        registerXlideCommand('xlide.revertAllAgentChanges', async () => {
            const pending = pendingAgentReviews();
            if (pending.length === 0) {
                return;
            }
            const choice = await vscode.window.showWarningMessage(
                `Revert ${pending.length === 1 ? '1 agent change' : `${pending.length} agent changes`}?`,
                {
                    modal: true,
                    detail: 'Each module goes back to what it held before the agent wrote it, and a module the agent '
                        + 'created is removed. A module changed again since the agent wrote it is left as it is.',
                },
                'Revert All',
            );
            if (choice !== 'Revert All') {
                return;
            }
            for (const { filePath, moduleName } of pending) {
                await revertAgentChange(agentDiffDeps, filePath, moduleName);
            }
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
                const note = modules.length === 0 ? await noCodeNote(bridge, options.input.filePath) : '';
                return textResult(`${JSON.stringify(modules, null, 2)}${note}`);
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
                const { filePath, moduleName, startLine, endLine, ranges, procedures } = options.input;
                if (ranges !== undefined && !Array.isArray(ranges)) {
                    return textResult('ranges must be a list of {startLine, endLine}.');
                }
                if (procedures !== undefined && !Array.isArray(procedures)) {
                    return textResult('procedures must be a list of procedure names.');
                }
                const result = await bridge.call<{ source: string }>(
                    'readModule',
                    { path: filePath, module: moduleName },
                    token,
                );
                const contentToken = moduleContentToken(result.source);
                const lines = splitModuleLines(result.source);
                // A window or a part is over the WHOLE module, so the token
                // still describes what a later conditional write is checked
                // against.
                if (ranges?.length || procedures?.length) {
                    // Several parts in one call, each under a line that says
                    // which lines it is, so an edit can name them back.
                    const parts = readModuleParts(result.source, { ranges, procedures });
                    if (!parts.ok) {
                        return textResult(parts.message);
                    }
                    return textResult([
                        `contentToken: ${contentToken} (${lines.length} lines)`,
                        ...parts.parts.map((part) => `${modulePartHeading(part)}\n${part.text}`),
                    ].join('\n'));
                }
                if (startLine === undefined && endLine === undefined) {
                    return textResult(`contentToken: ${contentToken} (${lines.length} lines)\n${result.source}`);
                }
                const from = Math.max(1, startLine ?? 1);
                const to = Math.min(lines.length, endLine ?? lines.length);
                return textResult(
                    `contentToken: ${contentToken} (lines ${from}-${to} of ${lines.length})\n${lines.slice(from - 1, to).join('\n')}`,
                );
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
                const summary = await writeModuleForAgent({
                    command: 'xlide_writeModule',
                    operationLabel: 'Write module',
                    filePath,
                    moduleName,
                    source,
                    kind,
                    wantsReview,
                    before: { source: beforeSource, existed: beforeExisted },
                });
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
        // xlide_editModule  (requires user confirmation)
        // ----------------------------------------------------------------
        // Parts of a module changed in one call (issue #95): line ranges and
        // procedures, named as the read showed them. The edits only mean
        // anything against that read, so the token is required here where
        // the whole-module write leaves it optional.
        vscode.lm.registerTool<EditModuleInput>('xlide_editModule', {
            async invoke(options, _token) {
                const { filePath, moduleName, expectedContentToken, edits } = options.input;
                if (!expectedContentToken) {
                    return textResult(
                        'expectedContentToken is required: the edits name lines of the module as xlide_readModule '
                        + 'last showed it. Read the module and pass the contentToken from that read.',
                    );
                }
                if (!Array.isArray(edits) || edits.length === 0) {
                    return textResult('edits must list at least one edit.');
                }
                let before: string;
                try {
                    before = (await bridge.call<{ source: string }>('readModule', { path: filePath, module: moduleName })).source;
                } catch (err) {
                    return textResult(
                        `Module "${moduleName}" could not be read: ${errorMessage(err)}. `
                        + 'xlide_editModule changes a module the file has; xlide_writeModule creates one.',
                    );
                }
                const stale = checkModuleContentToken(before, expectedContentToken, moduleName);
                if (stale) {
                    return textResult(stale.message);
                }
                const edited = applyModuleEdits(before, edits);
                if (!edited.ok) {
                    return textResult(edited.message);
                }
                const wantsReview = options.toolInvocationToken !== undefined && agentWriteDiffsEnabled();
                const summary = await writeModuleForAgent({
                    command: 'xlide_editModule',
                    operationLabel: 'Edit module',
                    filePath,
                    moduleName,
                    source: edited.source,
                    wantsReview,
                    before: { source: before, existed: true },
                });
                const now = await agentDiffDeps.readModuleSource(filePath, moduleName).catch(() => edited.source);
                const lines = [
                    summary,
                    `Module "${moduleName}" edited. contentToken: ${moduleContentToken(now)} (${splitModuleLines(now).length} lines)`,
                ];
                // The engine stores a module as it keeps one: blank lines
                // above its code are dropped. Where that moved the lines, the
                // arithmetic below would name the wrong ones.
                if (now.replace(/\r\n?/g, '\n') === edited.source.replace(/\r\n?/g, '\n')) {
                    lines.push(...edited.applied.map((edit) => edit.newStartLine === undefined
                        ? `- ${edit.label}: removed`
                        : `- ${edit.label}: now lines ${edit.newStartLine}-${edit.newEndLine}`));
                } else {
                    lines.push('The module was stored in a different layout than the edits made (blank lines above its code are dropped); read it again for line numbers.');
                }
                return textResult(lines.join('\n'));
            },
            async prepareInvocation(options, _token) {
                const { filePath, moduleName, edits } = options.input;
                const count = Array.isArray(edits) ? edits.length : 0;
                return {
                    invocationMessage: `Editing VBA module "${moduleName}"`,
                    confirmationMessages: {
                        title: 'Edit VBA Module',
                        message: new vscode.MarkdownString(
                            `Apply ${count === 1 ? '1 edit' : `${count} edits`} to **${moduleName}** in \`${filePath}\`?\n\n` +
                            'This changes the named lines of the module and saves the project.',
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
                const note = result.modules.length === 0
                    ? await noCodeNote(bridge, options.input.filePath)
                    : '';
                return textResult(`${JSON.stringify(result, null, 2)}${note}`);
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
        // xlide_listReferences
        // ----------------------------------------------------------------
        vscode.lm.registerTool<ListReferencesInput>('xlide_listReferences', {
            async invoke(options, token) {
                const result = await bridge.call<{ references: unknown[] }>(
                    'listReferences',
                    { path: options.input.filePath },
                    token,
                );
                return textResult(JSON.stringify(result.references, null, 2));
            },
        }),

        // ----------------------------------------------------------------
        // xlide_addReference  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<AddReferenceInput>('xlide_addReference', {
            async invoke(options, _token) {
                const { filePath, library } = options.input;
                const { result, summary } = await withWriteAudit({
                    command: 'xlide_addReference',
                    operation: 'add-reference',
                    projectPath: filePath,
                    failedSummary: 'Add reference: 0 changed, 1 failed',
                }, async () => {
                    // Writing the project, so it is coordinated with the host
                    // the same way a module write is.
                    const added = await runWriteWithHostCoordination(filePath, () =>
                        bridge.call<{ added: boolean; name: string }>('addReference', {
                            path: filePath,
                            library,
                        }));
                    return {
                        result: added,
                        summary: formatChangeSummary({
                            operation: 'Add reference',
                            changed: added.added ? [added.name] : [],
                        }),
                    };
                });
                return textResult(result.added
                    ? `${summary}\nThe project now references ${result.name}, so its VBA can name ${result.name} types early bound.`
                    : `The project already references ${result.name}; nothing was changed.`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, library } = options.input;
                return {
                    invocationMessage: `Adding a ${library} reference to "${filePath}"`,
                    confirmationMessages: {
                        title: 'Add Project Reference',
                        message: new vscode.MarkdownString(
                            `Add a reference to the **${library}** object library to \`${filePath}\`?`,
                        ),
                    },
                };
            },
        }),

        // ----------------------------------------------------------------
        // xlide_removeReference  (requires user confirmation)
        // ----------------------------------------------------------------
        vscode.lm.registerTool<RemoveReferenceInput>('xlide_removeReference', {
            async invoke(options, _token) {
                const { filePath, library } = options.input;
                const { result, summary } = await withWriteAudit({
                    command: 'xlide_removeReference',
                    operation: 'remove-reference',
                    projectPath: filePath,
                    failedSummary: 'Remove reference: 0 changed, 1 failed',
                }, async () => {
                    const gone = await runWriteWithHostCoordination(filePath, () =>
                        bridge.call<{ removed: boolean; name: string }>('removeReference', {
                            path: filePath,
                            library,
                        }));
                    return {
                        result: gone,
                        summary: formatChangeSummary({
                            operation: 'Remove reference',
                            changed: gone.removed ? [gone.name] : [],
                        }),
                    };
                });
                return textResult(result.removed
                    ? `${summary}\nThe project no longer references ${result.name}. Code that named ${result.name} types early bound no longer compiles.`
                    : `The project does not reference ${result.name}; nothing was changed.`);
            },
            async prepareInvocation(options, _token) {
                const { filePath, library } = options.input;
                return {
                    invocationMessage: `Removing the ${library} reference from "${filePath}"`,
                    confirmationMessages: {
                        title: 'Remove Project Reference',
                        message: new vscode.MarkdownString(
                            `Remove the reference to **${library}** from \`${filePath}\`? `
                            + 'Code that names it early bound will stop compiling.',
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
                // The tree's own edit path says the same; the watcher alone
                // would not move a sheet that has its first shape now.
                explorer.refreshShapes(filePath, { shapesChanged: true });
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
                const { hidden, rotation, zOrder, fill, line, font } = options.input;
                const look = [
                    hidden === undefined ? '' : hidden ? 'hidden' : 'shown',
                    fill || line || font ? 'restyled' : '',
                    rotation === undefined ? '' : `turned to ${rotation} degrees`,
                    zOrder ? `moved ${zOrder === 'front' || zOrder === 'back' ? `to the ${zOrder}` : zOrder}` : '',
                ].filter(Boolean);
                const styled = look.length > 0 ? `, ${look.join(', ')}` : '';
                const where = surface ? ` on **${surface}**` : '';
                return {
                    invocationMessage: `${shapeActionTitle(action)}${surface ? ` on "${surface}"` : ''} in "${filePath}"`,
                    confirmationMessages: {
                        title: shapeActionTitle(action),
                        message: new vscode.MarkdownString(
                            `${action === 'add' ? 'Add' : action === 'update' ? 'Change' : 'Delete'} ${what}${link}${styled}${where} in \`${filePath}\`?`,
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
        // xlide_importModules  (requires user confirmation)
        // ----------------------------------------------------------------
        // The Import Modules from Folder command, without its preview: the
        // confirmation says what the folder would change, and the apply is
        // the command's own (issue #92).
        vscode.lm.registerTool<ImportModulesInput>('xlide_importModules', {
            async invoke(options, _token) {
                const { filePath } = options.input;
                const planned = await planImport(options.input);
                if ('refused' in planned) {
                    return textResult(planned.refused);
                }
                const { plan, selectedIds } = planned;
                const wantsReview = options.toolInvocationToken !== undefined && agentWriteDiffsEnabled();
                const result = await applyImportModuleSyncPlan(ops, plan, selectedIds, {
                    command: 'xlide_importModules',
                    log: () => undefined,
                    agentReviewHandled: wantsReview,
                });
                if (wantsReview) {
                    // Each module the import wrote is an agent write: the
                    // same diff and Keep / Revert as xlide_writeModule gives.
                    for (const module of result.written) {
                        try {
                            const after = await agentDiffDeps.readModuleSource(filePath, module.moduleName);
                            void presentAgentModuleWrite(filePath, module.moduleName, {
                                before: module.before,
                                beforeExisted: module.beforeExisted,
                                after,
                            });
                        } catch {
                            // Written, but not readable back: nothing to show.
                        }
                    }
                }
                return textResult(JSON.stringify({
                    filePath,
                    importFolder: plan.folderPath,
                    importMode: plan.importMode,
                    updated: result.written.filter((module) => module.beforeExisted).map((module) => module.moduleName),
                    created: result.written.filter((module) => !module.beforeExisted).map((module) => module.moduleName),
                    removed: result.removed,
                    skipped: result.skipped.map((skip) => ({ module: skip.moduleName, file: skip.relativeName, reason: skip.reason })),
                    failed: result.failed,
                    changeSummary: result.summary,
                }, null, 2));
            },
            async prepareInvocation(options, _token) {
                const { filePath } = options.input;
                let detail = 'This writes modules into the file and updates <project>.xlide_settings.json.';
                try {
                    const planned = await planImport(options.input);
                    if ('plan' in planned) {
                        const chosen = selectedModuleSyncItems(planned.plan, planned.selectedIds);
                        const count = (status: string): number => chosen.filter((item) => item.status === status).length;
                        const acted = count('will-update') + count('will-create') + count('will-remove');
                        detail = `From \`${planned.plan.folderPath}\` (${planned.plan.importMode}): `
                            + `${count('will-update')} to update, ${count('will-create')} to create, `
                            + `${count('will-remove')} to delete, ${chosen.length - acted} skipped.\n\n${detail}`;
                    }
                } catch {
                    // The confirmation says what it can; invoke reports the failure.
                }
                return {
                    invocationMessage: `Importing VBA modules into "${filePath}"`,
                    confirmationMessages: {
                        title: 'Import VBA Modules',
                        message: new vscode.MarkdownString(`Import modules into \`${filePath}\`?\n\n${detail}`),
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
