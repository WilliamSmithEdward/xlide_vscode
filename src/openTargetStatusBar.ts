// Where Ctrl+Alt+O will go, shown before it is pressed.
//
// The Open in Office Application keybindings act on the project the window
// points at (projectTarget.ts), and do nothing when it points at none - the
// right answer, since opening the wrong workbook is worse than opening none,
// but a key that silently does nothing looks broken. This item names the file
// the key would open, or says it would open nothing and why.
//
// Desktop only, like the keybindings: there is no Office to open in a browser.

import * as path from 'path';
import * as vscode from 'vscode';
import { findMacroContainerFiles } from './macroContainerDiscovery';
import { containerAppNameForPath, MACRO_CONTAINER_GLOB } from './macroContainerUi';
import { osPlatform, type OsPlatform } from './util/osPlatform';
import {
    onDidChangeProjectTargetInputs,
    resolveProjectTarget,
    type ProjectTarget,
    type ProjectTargetSource,
} from './projectTarget';

/** Why the ladder chose the file, as the end of a sentence. */
const REASONS: Record<ProjectTargetSource, string> = {
    sidebarSelection: 'it is the file selected in the XLIDE sidebar',
    activeEditor: 'a module of it is the editor in front',
    lastOpened: 'it is the file you last had a module open from',
    singleProject: 'it is the only one in this window',
};

export interface OpenTargetHint {
    text: string;
    tooltip: string;
    command: string;
}

/** The keys, as the platform's keybinding declares them. */
function keys(platform: OsPlatform): { open: string; readOnly: string } {
    return platform === 'darwin'
        ? { open: 'Cmd+Alt+O', readOnly: 'Cmd+Alt+Shift+O' }
        : { open: 'Ctrl+Alt+O', readOnly: 'Ctrl+Alt+Shift+O' };
}

/**
 * What the item says for a window's projects and the target resolved from
 * them, or undefined when there is nothing to open at all and the item hides.
 */
export function openTargetHint(
    projectCount: number,
    target: ProjectTarget | undefined,
    platform: OsPlatform = osPlatform(),
): OpenTargetHint | undefined {
    if (projectCount === 0) {
        return undefined;
    }
    const { open, readOnly } = keys(platform);
    if (!target) {
        return {
            text: '$(link-external) ambiguous',
            tooltip: `${open} opens nothing: this window has ${projectCount} files, and nothing says which one. `
                + 'Select one in the XLIDE sidebar, or open one of its modules.',
            command: 'xlide.sidebar.focus',
        };
    }
    const name = path.basename(target.filePath);
    return {
        text: `$(link-external) ${name}`,
        tooltip: `${open} opens ${name} in ${containerAppNameForPath(target.filePath)} (${readOnly} opens it read-only), `
            + `because ${REASONS[target.source]}.`,
        command: 'xlide.openInOfficeApp',
    };
}

/** The status bar item, kept current with everything the ladder reads. */
export class OpenTargetStatusBar implements vscode.Disposable {
    private readonly item: vscode.StatusBarItem;
    private readonly disposables: vscode.Disposable[] = [];
    private projects: Promise<vscode.Uri[]> | undefined;
    private renderVersion = 0;

    constructor(private readonly workspaceState: vscode.Memento) {
        this.item = vscode.window.createStatusBarItem('xlide.openTarget', vscode.StatusBarAlignment.Left, 99);
        this.item.name = 'XLIDE: Where Open in Office Application Goes';
        const projectsChanged = (): void => {
            this.projects = undefined;
            this.render();
        };
        const watcher = vscode.workspace.createFileSystemWatcher(MACRO_CONTAINER_GLOB, false, true, false);
        this.disposables.push(
            this.item,
            watcher,
            watcher.onDidCreate(projectsChanged),
            watcher.onDidDelete(projectsChanged),
            vscode.workspace.onDidChangeWorkspaceFolders(projectsChanged),
            vscode.window.onDidChangeActiveTextEditor(() => this.render()),
            onDidChangeProjectTargetInputs(() => this.render()),
        );
        this.render();
    }

    /** Resolves the target again; a render that is overtaken draws nothing. */
    private render(): void {
        const version = ++this.renderVersion;
        void (async () => {
            this.projects ??= findMacroContainerFiles();
            const projects = await this.projects;
            const target = await resolveProjectTarget({ projects, workspaceState: this.workspaceState });
            if (version !== this.renderVersion) {
                return;
            }
            const hint = openTargetHint(projects.length, target);
            if (!hint) {
                this.item.hide();
                return;
            }
            this.item.text = hint.text;
            this.item.tooltip = hint.tooltip;
            this.item.command = hint.command;
            this.item.show();
        })().catch(() => {
            // A workspace that cannot be searched right now; the next change draws it.
        });
    }

    dispose(): void {
        for (const disposable of this.disposables) {
            disposable.dispose();
        }
    }
}
