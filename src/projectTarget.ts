// Which project a command acts on when nothing named one.
//
// A command reached from a tree row carries that row, and a command reached
// from an open module can read the editor. A keybinding carries neither, and
// a palette entry pressed with the focus somewhere else carries neither, so
// something has to decide what "the project" means - and picking wrong opens
// the wrong file in Excel, which is worse than doing nothing.
//
// The ladder, strongest signal first:
//
//   1. sidebarSelection - the file picked in the XLIDE sidebar, while it is
//      still one of the workspace's. An explicit choice outranks everything.
//   2. activeEditor - an XLIDE module is in front right now.
//   3. lastOpened - a module of it was the last one open in this window. This
//      is what carries the answer once the editor moves to something else.
//   4. singleProject - the workspace has exactly one, so there is nothing to
//      be wrong about.
//
// Below that the answer is ambiguous and there is none: several projects, and
// nothing said which. Callers do nothing rather than guess.
//
// The sidebar resolves its own displayed project through here too, so the
// file its Open buttons act on and the file a keybinding acts on are always
// the same one.

import * as vscode from 'vscode';
import { activeLocalVbaEditor, decodeModuleUri, XLIDE_SCHEME } from './xlideFileSystem';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { sameProjectPath } from './projectIdentity';

/** Where the answer came from, in the ladder's own words. */
export type ProjectTargetSource =
    | 'sidebarSelection'
    | 'activeEditor'
    | 'lastOpened'
    | 'singleProject';

export interface ProjectTarget {
    filePath: string;
    source: ProjectTargetSource;
}

/** The sidebar's picked file, in workspace state so it survives a reload. */
export const SELECTED_PROJECT_STATE_KEY = 'xlide.sidebar.selectedProjectPath';

/**
 * The project a module was last open from in this window. Deliberately NOT
 * persisted: it is a fact about this session's work, and a stale one from
 * last week is exactly the wrong file to open.
 */
let lastOpenedProjectPath: string | undefined;

const targetInputsEmitter = new vscode.EventEmitter<void>();

/**
 * Fires when a rung of the ladder the window itself keeps may have changed:
 * the sidebar's pick, or the project last worked in. The active editor and
 * the workspace's list of projects have events of their own.
 */
export const onDidChangeProjectTargetInputs = targetInputsEmitter.event;

/** Records a project as the one most recently worked in. */
export function noteProjectOpened(projectPath: string): void {
    if (lastOpenedProjectPath !== projectPath) {
        lastOpenedProjectPath = projectPath;
        targetInputsEmitter.fire();
    }
}

/** Forgets it. For tests, and for a window with nothing open again. */
export function forgetLastOpenedProject(): void {
    if (lastOpenedProjectPath !== undefined) {
        lastOpenedProjectPath = undefined;
        targetInputsEmitter.fire();
    }
}

/** Stores the sidebar's pick, or clears it. */
export async function storeSelectedProject(
    workspaceState: vscode.Memento | undefined,
    filePath: string | undefined,
): Promise<void> {
    await workspaceState?.update(SELECTED_PROJECT_STATE_KEY, filePath);
    targetInputsEmitter.fire();
}

/** The project of an XLIDE module document, or undefined for anything else. */
function projectOfDocument(document: vscode.TextDocument): string | undefined {
    if (document.uri.scheme !== XLIDE_SCHEME) {
        return undefined;
    }
    try {
        return decodeModuleUri(document.uri).projectPath;
    } catch {
        // Not a module address; nothing to remember.
        return undefined;
    }
}

/**
 * Keeps {@link noteProjectOpened} fed, for the life of the returned
 * disposable. Both events matter: opening a module names its project, and so
 * does coming back to one already open.
 */
export function watchOpenedProjects(): vscode.Disposable {
    const note = (document: vscode.TextDocument | undefined): void => {
        const projectPath = document && projectOfDocument(document);
        if (projectPath) {
            noteProjectOpened(projectPath);
        }
    };
    for (const document of vscode.workspace.textDocuments) {
        note(document);
    }
    return vscode.Disposable.from(
        vscode.workspace.onDidOpenTextDocument(note),
        vscode.window.onDidChangeActiveTextEditor((editor) => note(editor?.document)),
    );
}

function findProject(
    projects: readonly vscode.Uri[],
    filePath: string,
): vscode.Uri | undefined {
    return projects.find((uri) => sameProjectPath(uri.fsPath, filePath));
}

export interface ProjectTargetOptions {
    /** The workspace's projects, when the caller already has them. */
    projects?: readonly vscode.Uri[];
    /** The sidebar's pick, already validated. Omit to read it from state. */
    selectedProjectPath?: string;
    /** Workspace state holding the sidebar's pick. */
    workspaceState?: vscode.Memento;
}

/**
 * The project to act on, or undefined when the window does not say.
 *
 * Every rung is checked against the workspace's own list, so a file that was
 * deleted, renamed or moved out answers nothing rather than a path that is no
 * longer there.
 */
export async function resolveProjectTarget(
    options: ProjectTargetOptions = {},
): Promise<ProjectTarget | undefined> {
    const projects = options.projects ?? await findMacroContainerFiles();
    const selected = options.selectedProjectPath
        ?? options.workspaceState?.get<string>(SELECTED_PROJECT_STATE_KEY);

    const ladder: Array<[ProjectTargetSource, string | undefined]> = [
        ['sidebarSelection', selected],
        ['activeEditor', activeEditorProjectPath()],
        ['lastOpened', lastOpenedProjectPath],
    ];
    for (const [source, candidate] of ladder) {
        const match = candidate ? findProject(projects, candidate) : undefined;
        if (match) {
            return { filePath: match.fsPath, source };
        }
    }
    return projects.length === 1
        ? { filePath: projects[0].fsPath, source: 'singleProject' }
        : undefined;
}

function activeEditorProjectPath(): string | undefined {
    const editor = activeLocalVbaEditor();
    return editor ? decodeModuleUri(editor.document.uri).projectPath : undefined;
}
