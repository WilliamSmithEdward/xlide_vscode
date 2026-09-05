import * as vscode from 'vscode';
import { MACRO_CONTAINER_GLOB } from './macroContainerUi';
import * as path from 'path';
import { errorMessage } from './util/errors';
import { debounce } from './util/debounce';
import { ProjectExplorer } from './projectExplorer';
import {
    XlideFileSystemProvider,
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    decodeModuleUri,
    isLocalXlideDocument,
} from './xlideFileSystem';
import { ProjectEngine } from './projectEngine';
import { analysisSourceForDocument, moduleLocationOfDocument } from './vbaDocumentLocation';
import { readFolderAnnotation } from './vba/folderAnnotation';
import { registerFormPreview } from './vbaFormPreview';
import { registerVb6FormDesigner } from './vb6FormDesigner';
import { registerAgentTools } from './agentTools';
import { registerCommands } from './commands';
import { registerVbaLanguageProviders } from './vbaLanguageProviders';
import { XlideStatusBar } from './statusBar';
import { VbaCaretProcedureTracker } from './vbaCaretProcedure';
import { registerXlideDirtyModuleBackups } from './xlideDirtyModuleBackups';
import { registerVbaEditorCommands } from './vbaEditorCommands';
import { registerXlideCommand } from './xlideCommandRegistration';
import { createRecordedOutputChannel } from './xlideOutputLog';
import { setExcelCoordinationLog } from './excelWorkbookCoordinator';
import { registerXlideGlobalSettingsWebview } from './globalSettingsWebview';
import {
    setXlideGlobalSettingValue,
    xlideExplorerAutoExpandCollapseFromConfig,
    xlideExplorerViewFromConfig,
    xlidePerformanceTraceFromConfig,
} from './globalSettings';
import { registerXlideSidebar } from './xlideSidebar';
import { AnalysisWorkerClient } from './analysisWorkerClient';
import { setExtensionAssetRoot } from './extensionAssets';
import { cleanupStaleVbaTestHostTempDirsAsync } from './vbaTestTempFiles';
import { setPerformanceTraceLogger } from './performanceTrace';
import { setProjectAnalysisWorker } from './vbaProjectWideAnalysis';
import { XLIDE_VBA_EDITOR_OVERRIDES } from './xlideVbaEditorOverrides';

// ---------------------------------------------------------------------------
// Activation
// ---------------------------------------------------------------------------

export function activate(context: vscode.ExtensionContext): void {
    // Bundled assets (assets/**) resolve against the installed extension root,
    // not the src/ tree (see extensionAssets.ts).
    setExtensionAssetRoot(context.extensionUri.fsPath);
    const out = createRecordedOutputChannel(vscode.window.createOutputChannel('XLIDE'));
    setPerformanceTraceLogger(
        (line) => out.appendLine(line),
        () => xlidePerformanceTraceFromConfig(vscode.workspace.getConfiguration('xlide')).value,
    );
    // Route Excel-coordination traces (file-system save path, shared module
    // operations) to the XLIDE output channel.
    setExcelCoordinationLog((line) => out.appendLine(line));
    out.appendLine('XLIDE activating...');

    // Dev-only commands stay out of the command palette unless the extension
    // host is running in Development mode (see menus.commandPalette).
    void vscode.commands.executeCommand(
        'setContext',
        'xlide.devMode',
        context.extensionMode === vscode.ExtensionMode.Development,
    );

    void cleanupStaleVbaTestHostTempDirsAsync()
        .then((cleanup) => {
            if (cleanup.deleted > 0 || cleanup.failed > 0) {
                out.appendLine(
                    `VBA test temp cleanup: scanned=${cleanup.scanned} deleted=${cleanup.deleted} failed=${cleanup.failed}`,
                );
            }
        })
        .catch((err) => {
            const message = errorMessage(err);
            out.appendLine(`VBA test temp cleanup skipped: ${message}`);
        });

    const bridge = new ProjectEngine(context, out);
    const fsProvider = new XlideFileSystemProvider(bridge);
    registerFormPreview(context, bridge);
    registerVb6FormDesigner(context, bridge);
    const explorer = new ProjectExplorer(bridge, out);
    // One answer to "which procedure is the caret in", shared by the status bar
    // and the tree so the two never disagree.
    const caret = new VbaCaretProcedureTracker();
    const statusBar = new XlideStatusBar(caret);
    // The project engine runs in-process: there is no backend to install,
    // start, probe, or recover, so nothing gates the tree or the sidebar.
    const sidebar = registerXlideSidebar({
        workspaceState: context.workspaceState,
    });
    sidebar.refresh();

    // Keep reference outside subscriptions for post-start auto-expand and reveal.
    const treeView = vscode.window.createTreeView('xlide.explorer', {
        treeDataProvider: explorer,
        showCollapseAll: true,
    });

    // The Tree / Folders buttons above the explorer are the setting, so the
    // button, the settings page, and settings.json all say the same thing.
    const applyExplorerView = (): void => {
        explorer.setView(xlideExplorerViewFromConfig(vscode.workspace.getConfiguration('xlide')).value);
    };
    applyExplorerView();

    /**
     * Select the row for the procedure the caret is in, the way the VBE's own
     * explorer marks it. Falls back to the module for a caret in the
     * declarations section, and for a procedure the tree cannot name yet -
     * an unsaved rename leaves the container calling it something else.
     */
    const revealCaretProcedure = (): void => {
        const position = caret.current;
        if (!position || !treeView.visible) { return; }
        if (!xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) {
            return;
        }
        const node = (position.procedure
            && explorer.getProcedureNode(position.projectPath, position.moduleName, position.label))
            ?? explorer.getModuleNode(position.projectPath, position.moduleName);
        if (node) {
            // reveal() rejects when the tree cannot place the element - a row
            // the last refresh dropped, say. Marking the caret is cosmetic, so
            // the rejection is swallowed rather than left unhandled.
            void treeView.reveal(node, { select: true, focus: false })
                .then(undefined, () => { /* the row is gone; nothing to mark */ });
        }
    };

    // VBA language services: syntax-aware symbol index + providers. The
    // analysis worker keeps full diagnostic passes off the extension-host
    // thread; if it cannot start, diagnostics fall back to in-host analysis.
    const analysisWorkerClient = new AnalysisWorkerClient(
        path.join(context.extensionPath, 'out', 'analysisWorker.js'),
        (line: string) => out.appendLine(line),
    );
    context.subscriptions.push(new vscode.Disposable(() => analysisWorkerClient.dispose()));
    // Analyze Workbook (the command, the agent tool, and the support bundle's
    // anonymized report) rides the same worker so a large module's analysis
    // never blocks the host mid-command.
    setProjectAnalysisWorker(analysisWorkerClient);
    const vbaIndex = registerVbaLanguageProviders(context, bridge, analysisWorkerClient);
    registerVbaEditorCommands(context);
    registerXlideVbaLanguageSync(context, out);
    void ensureXlideVbaEditorOverrides(out);

    // Register the explorer/palette commands up front in their own subscription
    // batch, isolated in a try/catch. The tree view above is created
    // independently, so if any other registration later in activation throws,
    // the tree would stay clickable while its commands were missing - surfacing
    // as "command 'xlide.openModule' not found" until the window is reloaded.
    // Registering commands first (and not letting one failure abort the rest)
    // keeps that from happening.
    try {
        context.subscriptions.push(
            ...registerCommands(context, bridge, explorer, fsProvider, out, vbaIndex),
        );
    } catch (err) {
        out.appendLine(
            'XLIDE: command registration failed during activation: '
            + (err instanceof Error ? err.message : String(err)),
        );
    }

    // Register the virtual filesystem resiliently too. openModule opens
    // xlide-vba:// documents through this provider, so if a stale re-activation
    // has already registered the scheme, letting registerFileSystemProvider throw
    // would abort the rest of activation and strand the tree with "no provider
    // for xlide-vba://..." - the same failure class as the command registration
    // above, just with a different symptom ("cannot open module").
    try {
        context.subscriptions.push(
            vscode.workspace.registerFileSystemProvider(XLIDE_SCHEME, fsProvider, {
                isCaseSensitive: process.platform !== 'win32',
                isReadonly: false,
            }),
        );
    } catch (err) {
        out.appendLine(
            'XLIDE: filesystem provider registration failed (already registered?): '
            + (err instanceof Error ? err.message : String(err)),
        );
    }

    context.subscriptions.push(
        out,

        registerXlideDirtyModuleBackups(context, out),

        treeView,
        explorer,

        // Item 6: Reveal active module in the XLIDE Explorer tree.
        // Also drives accordion collapse: only the active module stays expanded.
        // Debounced so rapid tab switches (e.g. Ctrl+W spam) coalesce into a
        // single setActiveModule + reveal, avoiding overlapping async reveal
        // calls that could leave stale modules expanded.
        (() => {
            let pending: { projectPath: string; moduleName: string } | undefined;
            const apply = debounce(() => {
                if (!pending) { return; }
                const { projectPath, moduleName } = pending;
                pending = undefined;
                // Honor the user's auto-expand/collapse preference: when off, the
                // tree never follows the active tab or accordion-collapses.
                if (!xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) {
                    return;
                }
                explorer.setActiveModule(projectPath, moduleName);
                const node = explorer.getModuleNode(projectPath, moduleName);
                if (node && treeView.visible) {
                    // Expanding the module is what lists its procedures, so the
                    // caret's own row only exists once this has landed.
                    void treeView.reveal(node, { select: true, focus: false, expand: true })
                        .then(() => revealCaretProcedure(), () => { /* reveal is best-effort */ });
                }
            }, 60);
            const subscription = vscode.window.onDidChangeActiveTextEditor((editor) => {
                // Focus moved off any text editor (the Output panel, terminal, the
                // tree, a webview, or the last tab closed). Leave the tree as-is: a
                // project only collapses when focus moves to a module in a DIFFERENT
                // project, never on transient focus loss.
                if (!editor) {
                    pending = undefined;
                    apply.cancel();
                    // Nothing open at all is the last editor closing, not a
                    // transient loss of focus: no module is being edited, so
                    // the folder layout goes back to its resting shape.
                    if (vscode.window.visibleTextEditors.length === 0
                        && xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) {
                        explorer.collapseAllFolders();
                    }
                    return;
                }
                // A project module's virtual document and a VB6 module's own
                // file are both modules of a project; the locator answers for
                // both, so the tree follows a `.frm` the way it follows a
                // `.bas` in a workbook.
                const location = moduleLocationOfDocument(editor.document);
                if (!location) { return; }
                pending = { projectPath: location.projectPath, moduleName: location.moduleName };
                apply();
            });
            return new vscode.Disposable(() => {
                apply.dispose();
                subscription.dispose();
            });
        })(),

        // Accordion: if the user manually clicks the expand arrow on a module node,
        // collapse all sibling modules under the same project. A folder opened
        // by hand keeps that until the editor moves to a different folder.
        treeView.onDidExpandElement((e) => {
            if (!xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) { return; }
            if (e.element.kind === 'module' && e.element.filePath && e.element.moduleName) {
                explorer.setActiveModule(e.element.filePath, e.element.moduleName);
            }
            explorer.notifyFolderExpansion(e.element, true);
        }),

        // When the user manually collapses the active project, stop forcing it
        // Expanded so it stays collapsed - otherwise the next refresh re-stamps it
        // Expanded against the still-set active-project key and springs it open.
        // A folder shut by hand is remembered the same way one opened by hand is.
        treeView.onDidCollapseElement((e) => {
            if (e.element.kind === 'project' && e.element.filePath) {
                explorer.notifyProjectCollapsed(e.element.filePath);
            }
            if (xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) {
                explorer.notifyFolderExpansion(e.element, false);
            }
        }),

        // The Tree / Folders buttons, which write the setting rather than a
        // second piece of state that could disagree with it.
        registerXlideCommand('xlide.explorer.showTree', async () => {
            await setXlideGlobalSettingValue(vscode.workspace.getConfiguration('xlide'), 'explorer.view', 'tree');
        }),
        registerXlideCommand('xlide.explorer.showFolders', async () => {
            await setXlideGlobalSettingValue(vscode.workspace.getConfiguration('xlide'), 'explorer.view', 'folders');
        }),
        vscode.workspace.onDidChangeConfiguration((e) => {
            if (e.affectsConfiguration('xlide.explorer.view')) {
                applyExplorerView();
            }
        }),

        // Refresh the explorer when macro-container files are added or removed
        // Debounced so rapid file-system events (save storms) coalesce into one refresh.
        (() => {
            const debouncedRefresh = debounce(() => explorer.refresh(), 200);
            const watcher = vscode.workspace.createFileSystemWatcher(MACRO_CONTAINER_GLOB);
            const createSubscription = watcher.onDidCreate(debouncedRefresh);
            const deleteSubscription = watcher.onDidDelete(debouncedRefresh);
            return new vscode.Disposable(() => {
                debouncedRefresh.dispose();
                createSubscription.dispose();
                deleteSubscription.dispose();
                watcher.dispose();
            });
        })(),

        // Show the XLIDE output channel (used by the explorer welcome view).
        registerXlideCommand('xlide.showOutput', () => {
            out.show(true);
        }),

        // The caret moving to another procedure moves the tree's selection with
        // it. The tracker only fires when the procedure actually changes, so
        // typing inside one does not touch the tree.
        caret.onDidChange(() => revealCaretProcedure()),
        caret,

        ...sidebar.disposables,
        registerXlideGlobalSettingsWebview(out),
        ...registerAgentTools(context, bridge, explorer, fsProvider, vbaIndex),

        statusBar,
        bridge,
    );

    // When the symbol index updates (e.g. after a rename or save), refresh
    // the matching module's sub list in the explorer so renamed procedures
    // appear immediately. A module an agent or a command rewrote can have a
    // different @Folder annotation, so the folder layout follows it too.
    context.subscriptions.push(
        vbaIndex.onDidChange(({ projectPath, moduleName }) => {
            if (!projectPath || !moduleName) {
                explorer.refresh();
                return;
            }
            explorer.refreshModuleSubs(projectPath, moduleName);
            const source = vbaIndex.peekModule(projectPath, moduleName)?.source;
            if (source !== undefined) {
                explorer.setModuleFolder(projectPath, moduleName, readFolderAnnotation(source).folder);
            }
        }),

        // The folder layout follows the open editor, not the file on disk: an
        // annotation edited in a module moves it while you type. Debounced,
        // since this reads the module text and a keystroke is not a folder.
        (() => {
            const pending = new Map<string, vscode.TextDocument>();
            const flush = debounce(() => {
                for (const document of pending.values()) {
                    if (document.isClosed) { continue; }
                    const location = moduleLocationOfDocument(document);
                    if (!location) { continue; }
                    explorer.setModuleFolder(
                        location.projectPath,
                        location.moduleName,
                        readFolderAnnotation(analysisSourceForDocument(document)).folder,
                    );
                }
                pending.clear();
            }, 300);
            const changed = vscode.workspace.onDidChangeTextDocument((e) => {
                if (!moduleLocationOfDocument(e.document)) { return; }
                pending.set(e.document.uri.toString(), e.document);
                flush();
            });
            // With the editor closed the container is the truth again, whether
            // the edit was saved into it or thrown away.
            const closed = vscode.workspace.onDidCloseTextDocument((document) => {
                pending.delete(document.uri.toString());
                const location = moduleLocationOfDocument(document);
                if (location) {
                    explorer.forgetModuleFolder(location.projectPath, location.moduleName);
                }
            });
            return new vscode.Disposable(() => {
                flush.dispose();
                changed.dispose();
                closed.dispose();
            });
        })(),
    );

    // Item 7: Auto-expand the first project so modules are visible. Listing
    // projects only globs the workspace; expanding the first one performs the
    // first bridge call, which lazy-starts the backend. Workspaces without
    // Excel workbooks (or with the explorer hidden) never open a workbook.
    if (treeView.visible) {
        const autoExpandTimer = setTimeout(() => {
            if (!treeView.visible) { return; }
            void explorer.warmProjectCache().then(firstNode => {
                if (firstNode && treeView.visible) {
                    // Returned, not fired and forgotten: a reveal that cannot
                    // place the node rejects, and the catch below is the only
                    // thing standing between that and an unhandled rejection.
                    return treeView.reveal(firstNode, { select: false, focus: false, expand: true });
                }
                return undefined;
            }).catch((err: unknown) => out.appendLine(
                `XLIDE auto-expand failed: ${err instanceof Error ? err.message : String(err)}`,
            ));
        }, 250);
        context.subscriptions.push(new vscode.Disposable(() => clearTimeout(autoExpandTimer)));
    }
}

export function deactivate(): void { /* nothing async needed */ }

interface ConfigurationLanguageInspect<T> {
    globalLanguageValue?: T;
    workspaceLanguageValue?: T;
    workspaceFolderLanguageValue?: T;
}

function languageOverrideValue<T>(
    inspected: ConfigurationLanguageInspect<T> | undefined,
): T | undefined {
    return inspected?.workspaceFolderLanguageValue
        ?? inspected?.workspaceLanguageValue
        ?? inspected?.globalLanguageValue;
}

/**
 * Extension configuration defaults lose to a user's global editor settings.
 * Shape the XLIDE minimap into a clean rail while preserving explicit
 * [xlide-vba] choices, except for the stale hidden-rail values written by an
 * earlier development build.
 */
async function ensureXlideVbaEditorOverrides(out: vscode.OutputChannel): Promise<void> {
    const config = vscode.workspace.getConfiguration('editor', { languageId: XLIDE_VBA_LANGUAGE_ID });
    const staleHiddenRail =
        languageOverrideValue<boolean>(config.inspect('minimap.enabled')) === false &&
        languageOverrideValue<number>(config.inspect('overviewRulerLanes')) === 0;
    for (const override of XLIDE_VBA_EDITOR_OVERRIDES) {
        const inspected = config.inspect(override.key);
        const hasLanguageOverride = languageOverrideValue(inspected) !== undefined;
        if (hasLanguageOverride && !staleHiddenRail) {
            continue;
        }
        if (config.get(override.key) === override.value) {
            continue;
        }
        try {
            await config.update(override.key, override.value, vscode.ConfigurationTarget.Global, true);
            out.appendLine(`Set editor.${override.key}=${override.value} for [${XLIDE_VBA_LANGUAGE_ID}].`);
        } catch (err) {
            const message = errorMessage(err);
            out.appendLine(`Could not set editor.${override.key} for [${XLIDE_VBA_LANGUAGE_ID}]: ${message}`);
        }
    }
}

function registerXlideVbaLanguageSync(context: vscode.ExtensionContext, out: vscode.OutputChannel): void {
    const syncDocument = (document: vscode.TextDocument): void => {
        if (document.uri.scheme !== XLIDE_SCHEME || document.languageId === XLIDE_VBA_LANGUAGE_ID) {
            return;
        }
        void Promise.resolve(vscode.languages.setTextDocumentLanguage(document, XLIDE_VBA_LANGUAGE_ID))
            .catch((err: Error) => {
                out.appendLine(`Could not set XLIDE VBA language for ${document.uri.toString()}: ${err.message}`);
            });
    };

    for (const document of vscode.workspace.textDocuments) {
        syncDocument(document);
    }
    for (const editor of vscode.window.visibleTextEditors) {
        syncDocument(editor.document);
    }

    context.subscriptions.push(
        vscode.workspace.onDidOpenTextDocument(syncDocument),
        vscode.window.onDidChangeVisibleTextEditors((editors) => {
            for (const editor of editors) {
                syncDocument(editor.document);
            }
        }),
    );
}

