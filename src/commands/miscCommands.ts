import * as vscode from 'vscode';
import * as path from 'path';
import { xlideExplorerAutoExpandCollapseFromConfig } from '../globalSettings';
import { MACRO_CONTAINER_GLOB } from '../macroContainerUi';
import { registerXlideCommand } from '../xlideCommandRegistration';
import type { XlideNode } from '../projectExplorer';
import { errorMessage } from '../util/errors';
import {
    type CommandDeps,
    outputLogger,
    openModuleDocument,
} from './shared';

export function registerMiscCommands(deps: CommandDeps): vscode.Disposable[] {
    const { bridge, explorer, out } = deps;

    const log = outputLogger(out);


    async function showClassModuleReferences(node: XlideNode): Promise<void> {
        if (!node.moduleName || !node.filePath) { return; }
        const originDoc = await openModuleDocument(node.filePath, node.moduleName);
        const editor = await vscode.window.showTextDocument(originDoc, { preview: false });
        const origin = new vscode.Position(0, 0);
        editor.selection = new vscode.Selection(origin, origin);
        await vscode.commands.executeCommand('references-view.findReferences', originDoc.uri, origin);
    }

    /**
     * The editor for a tree node's module: the project module's virtual
     * document, or - when the container's modules are files, a VB6 project -
     * the file itself, which already carries the VBA language by extension.
     */
    const showModuleEditor = async (node: XlideNode): Promise<vscode.TextEditor> => {
        if (node.moduleFilePath) {
            const doc = await vscode.workspace.openTextDocument(vscode.Uri.file(node.moduleFilePath));
            return vscode.window.showTextDocument(doc, { preview: false });
        }
        const doc = await openModuleDocument(node.filePath, node.moduleName ?? '');
        return vscode.window.showTextDocument(doc, { preview: false });
    };

    return [
        registerXlideCommand('xlide.refreshExplorer', () => {
            explorer.refresh();
        }),

        // In-tree "Load failed - click to retry" placeholder (e.g. after Excel
        // briefly held the project file). Retries just the failed listing
        // instead of collapsing the whole tree with a full refresh.
        registerXlideCommand('xlide.retryExplorerLoad', (node: XlideNode) => {
            explorer.retryLoad(node);
        }),

        // Open a module (or navigate to a sub's line inside one)
        registerXlideCommand('xlide.openModule', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            // A module row's own click opens the row too, under the setting that
            // has the tree follow the editor.
            if (node.kind === 'module'
                && xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) {
                void explorer.expandModuleRow(node);
            }
            const editor = await showModuleEditor(node);

            // If a specific line was requested (sub navigation), move cursor there
            if (node.line !== undefined && node.line > 0) {
                const pos = new vscode.Position(node.line - 1, 0);
                editor.selection = new vscode.Selection(pos, pos);
                editor.revealRange(
                    new vscode.Range(pos, pos),
                    vscode.TextEditorRevealType.InCenterIfOutsideViewport,
                );
            }
        }),

        // Find all references to the procedure or class represented by a tree node
        registerXlideCommand('xlide.findReferences', async (node: XlideNode) => {
            if (!node?.moduleName) { return; }
            if (node.kind === 'module' && node.moduleType === 'class') {
                await showClassModuleReferences(node);
                return;
            }
            if (node.kind !== 'sub') { return; }
            const editor = await showModuleEditor(node);
            const doc = editor.document;

            // Locate the procedure name on its declaration line so the reference
            // search starts on the identifier. The node label is "<kind> <name>"
            // (kind may be "Property Get" etc.), so the bare name is the last token.
            const procName = node.label.split(' ').pop() ?? '';
            let pos = new vscode.Position(Math.max(0, (node.line ?? 1) - 1), 0);
            if (procName && node.line !== undefined && node.line > 0) {
                const lineText = doc.lineAt(node.line - 1).text;
                // Whole-word match so a short proc name that is a substring of a
                // preceding token (e.g. "i" inside "Public") does not mis-anchor.
                const escaped = procName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                const match = new RegExp(`\\b${escaped}\\b`).exec(lineText);
                const col = match ? match.index : lineText.indexOf(procName);
                if (col >= 0) {
                    pos = new vscode.Position(node.line - 1, col);
                }
            }

            // Move the active editor's cursor onto the identifier so the
            // references command resolves the correct symbol, then trigger it.
            editor.selection = new vscode.Selection(pos, pos);
            editor.revealRange(
                new vscode.Range(pos, pos),
                vscode.TextEditorRevealType.InCenterIfOutsideViewport,
            );
            await vscode.commands.executeCommand('references-view.findReferences', doc.uri, pos);
        }),

        // DEV: smoke test - verifies listModules + readModule against a workspace project
        registerXlideCommand('xlide.dev.smoke', async () => {
            log('[smoke] Starting smoke test...');

            const uris = (await vscode.workspace.findFiles(MACRO_CONTAINER_GLOB,
                '{**/node_modules/**,**/.venv/**,**/venv/**}'))
                .filter(u => !path.basename(u.fsPath).startsWith('~$'));

            if (uris.length === 0) {
                vscode.window.showErrorMessage('XLIDE Smoke: No project found in the workspace.');
                return;
            }

            let projectPath: string;
            if (uris.length === 1) {
                projectPath = uris[0].fsPath;
            } else {
                const pick = await vscode.window.showQuickPick(
                    uris.map(u => ({ label: path.basename(u.fsPath), description: u.fsPath, fsPath: u.fsPath })),
                    { title: 'XLIDE Smoke Test: pick a project' },
                );
                if (!pick) { return; }
                projectPath = pick.fsPath;
            }

            log(`[smoke] File: ${projectPath}`);

            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: 'XLIDE: Running smoke test...', cancellable: false },
                async () => {
                    try {
                        // Step 1: listModules
                        const modules = await bridge.call<Array<{ name: string; type: string }>>(
                            'listModules', { path: projectPath },
                        );
                        log(`[smoke] listModules OK - ${modules.length} module(s): ${modules.map(m => m.name).join(', ')}`);

                        if (modules.length === 0) {
                            vscode.window.showWarningMessage('XLIDE Smoke: project has no VBA modules.');
                            return;
                        }

                        // Step 2: readModule (prefer a non-document module)
                        const target = modules.find(m => m.type !== 'document') ?? modules[0];
                        const source = await bridge.call<string>(
                            'readModule', { path: projectPath, module: target.name, full: false },
                        );
                        log(`[smoke] readModule "${target.name}" OK - ${source.length} chars`);

                        log('[smoke] All checks passed.');
                        void vscode.window.showInformationMessage(
                            `XLIDE Smoke: OK - ${modules.length} modules, read "${target.name}" (${source.length} chars). See XLIDE Output for details.`,
                        );
                    } catch (err) {
                        const msg = errorMessage(err);
                        log(`[smoke] FAILED: ${msg}`);
                        vscode.window.showErrorMessage(`XLIDE Smoke FAILED: ${msg}`);
                    }
                },
            );
        }),
    ];
}
