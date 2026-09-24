// Mirrors the edits the XLIDE MCP server (xlide_mcp) makes into this window,
// the way XLIDE shows its own agent tools' writes. The server writes a file in
// its own process, which reached XLIDE only as a file that changed on disk:
// no before, so no diff and no Revert, and the first write to a project the
// tree merely lists was not reported at all (see projectFileChanges.ts). The
// server now says what each tool call changed, through the loopback API in
// xlideApiServer.ts, and this is what a report does here.
//
// A window takes a report's new state, and presents a review, only for a file
// it shows: in its tree, or with a module of it open. The server reports to
// every window, and one showing another folder must not open a diff for a
// file it does not list. A review this window already holds follows the
// module whatever it shows, as it follows every other write.

import * as vscode from 'vscode';
import { xlideAgentMirrorMcpEditsFromConfig } from './globalSettings';
import type { ProjectEngine } from './projectEngine';
import { takeReportedProjectFileChange } from './projectFileChanges';
import { projectIdentityKey } from './projectIdentity';
import { errorMessage } from './util/errors';
import { createKeyedAsyncLock } from './util/keyedAsyncLock';
import { splitVbaSource } from './vba/moduleSource';
import {
    agentWriteDiffsEnabled,
    discardPendingAgentReview,
    hasPendingAgentReview,
    presentAgentModuleWrite,
    renamePendingAgentReview,
    trackModuleWriteForAgentReview,
} from './xlideAgentDiff';
import {
    XlideApiServer,
    type AgentEditReport,
    type ReportAnswer,
    type XlideApiHandlers,
} from './xlideApiServer';
import { decodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';

export interface McpEditMirrorDeps {
    bridge: Pick<ProjectEngine, 'call'>;
    /** Whether the XLIDE tree lists the file as a project. */
    treeListsProject(filePath: string): Promise<boolean>;
    log(line: string): void;
}

export interface McpEditMirrorOptions extends McpEditMirrorDeps {
    /** The extension's version, which the record carries. */
    version: string;
    /** Where the record goes, for tests. */
    stateDir?: string;
}

/** Whether mirroring is on: xlide.agent.mirrorMcpEdits. */
export function mcpEditMirrorEnabled(): boolean {
    return xlideAgentMirrorMcpEditsFromConfig(vscode.workspace.getConfiguration('xlide')).value === true;
}

/**
 * Serves the API while xlide.agent.mirrorMcpEdits is on, starting and
 * stopping it as the setting changes, and keeps the record's workspace
 * folders current.
 */
export function mirrorMcpEdits(options: McpEditMirrorOptions): vscode.Disposable {
    const handlers = mcpEditMirrorHandlers(options);
    let live: XlideApiServer | undefined;
    let disposed = false;
    let turn = Promise.resolve();
    const wanted = (): boolean => !disposed && mcpEditMirrorEnabled();
    // In the order the setting changed, so a quick off and on cannot leave two
    // servers running, or none.
    const apply = (): void => {
        turn = turn.then(async () => {
            if (wanted() === (live !== undefined)) {
                return;
            }
            if (live) {
                live.dispose();
                live = undefined;
                options.log('Stopped mirroring MCP server edits.');
                return;
            }
            try {
                const started = await XlideApiServer.start({
                    handlers,
                    version: options.version,
                    workspaceFolders: workspaceFolderPaths,
                    log: options.log,
                    stateDir: options.stateDir,
                });
                if (!wanted()) {
                    // Turned off, or the window closed, while it started.
                    started.dispose();
                    return;
                }
                live = started;
                options.log(`Mirroring MCP server edits: listening on 127.0.0.1:${started.port}.`);
            } catch (err) {
                options.log(`Could not start mirroring MCP server edits: ${errorMessage(err)}`);
            }
        });
    };
    apply();
    const subscriptions = [
        vscode.workspace.onDidChangeConfiguration((event) => {
            if (event.affectsConfiguration('xlide.agent.mirrorMcpEdits')) {
                apply();
            }
        }),
        vscode.workspace.onDidChangeWorkspaceFolders(() => {
            try {
                live?.writeRecord();
            } catch (err) {
                options.log(`Could not update the MCP server's record of this window: ${errorMessage(err)}`);
            }
        }),
    ];
    return new vscode.Disposable(() => {
        disposed = true;
        subscriptions.forEach((subscription) => subscription.dispose());
        live?.dispose();
        live = undefined;
    });
}

function workspaceFolderPaths(): string[] {
    return (vscode.workspace.workspaceFolders ?? []).map((folder) => folder.uri.fsPath);
}

/** What each report does to this window. */
export function mcpEditMirrorHandlers(deps: McpEditMirrorDeps): XlideApiHandlers {
    // One report at a time per file, in the order they came: a review keeps
    // the before of the agent's first unreviewed write, so two reports for one
    // module must not cross.
    const inTurn = createKeyedAsyncLock();
    const shows = async (filePath: string): Promise<boolean> =>
        hasModuleOpen(filePath) || await deps.treeListsProject(filePath);
    const answered = (what: string, filePath: string, answer: ReportAnswer): ReportAnswer => {
        deps.log(`MCP server: ${what} in ${filePath} (${answer.shown ? 'shown' : 'not shown'} here, review ${answer.review}).`);
        return answer;
    };
    return {
        agentEdit: (report) => inTurn(projectIdentityKey(report.file), async () => {
            const shown = await shows(report.file);
            if (shown) {
                takeReportedProjectFileChange(report.file);
            }
            if (!report.afterExists) {
                // A deleted module has nothing left to review.
                discardPendingAgentReview(report.file, report.module);
            } else if (shown && agentWriteDiffsEnabled()) {
                // The after is XLIDE's own read, not the server's copy: Revert
                // compares the module with it before writing the before back.
                void presentAgentModuleWrite(report.file, report.module, {
                    before: inEngineForm(report.before),
                    beforeExisted: report.beforeExisted,
                    after: await engineSource(deps, report),
                });
            } else if (hasPendingAgentReview(report.file, report.module)) {
                trackModuleWriteForAgentReview(report.file, report.module, await engineSource(deps, report));
            }
            const what = `${report.afterExists ? (report.beforeExisted ? 'wrote' : 'created') : 'deleted'} ${report.module}`;
            return answered(what, report.file, answer(shown, report.file, report.module));
        }),
        moduleRenamed: (report) => inTurn(projectIdentityKey(report.file), async () => {
            const shown = await shows(report.file);
            if (shown) {
                // The tree and the editors on the module follow from here.
                takeReportedProjectFileChange(report.file);
            }
            // An unreviewed agent change goes with the module to its new name.
            renamePendingAgentReview(report.file, report.from, report.to);
            return answered(`renamed ${report.from} to ${report.to}`, report.file, answer(shown, report.file, report.to));
        }),
        fileChanged: (report) => inTurn(projectIdentityKey(report.file), async () => {
            const shown = await shows(report.file);
            if (shown) {
                takeReportedProjectFileChange(report.file);
            }
            return answered(`changed ${report.what ?? 'the file'}`, report.file, answer(shown, report.file));
        }),
    };
}

function answer(shown: boolean, filePath: string, moduleName?: string): ReportAnswer {
    const pending = moduleName !== undefined && hasPendingAgentReview(filePath, moduleName);
    return { shown, review: pending ? 'pending' : 'none' };
}

/**
 * The module's code as XLIDE reads it now. When that fails, the server's own
 * after is the best record left.
 */
async function engineSource(deps: Pick<McpEditMirrorDeps, 'bridge'>, report: AgentEditReport): Promise<string> {
    try {
        const read = await deps.bridge.call<{ source: string }>('readModule', {
            path: report.file,
            module: report.module,
        });
        return read.source;
    } catch {
        return inEngineForm(report.after);
    }
}

/**
 * Code from the server in the form the engine reads it: without the line
 * breaks above the first line. The server's header-stripped body can keep
 * them, and compared with XLIDE's own read, a write that changed nothing
 * would show a diff of blank lines.
 */
function inEngineForm(code: string): string {
    return splitVbaSource(code).body;
}

/** Whether a module of the file is open: as a document, or in a tab, diffs included. */
function hasModuleOpen(filePath: string): boolean {
    const project = projectIdentityKey(filePath);
    const ofProject = (uri: vscode.Uri): boolean => {
        if (uri.scheme !== XLIDE_SCHEME) {
            return false;
        }
        try {
            return projectIdentityKey(decodeModuleUri(uri).projectPath) === project;
        } catch {
            return false;
        }
    };
    if (vscode.workspace.textDocuments.some((document) => !document.isClosed && ofProject(document.uri))) {
        return true;
    }
    return vscode.window.tabGroups.all.some((group) => group.tabs.some((tab) => {
        const input = tab.input;
        if (input instanceof vscode.TabInputTextDiff) {
            return ofProject(input.modified);
        }
        return (input instanceof vscode.TabInputText || input instanceof vscode.TabInputCustom) && ofProject(input.uri);
    }));
}
