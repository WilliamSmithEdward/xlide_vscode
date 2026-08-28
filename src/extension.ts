import * as vscode from 'vscode';
import { MACRO_CONTAINER_GLOB } from './macroContainerUi';
import * as path from 'path';
import { errorMessage } from './util/errors';
import { debounce } from './util/debounce';
import { XlsmExplorer } from './xlsmExplorer';
import {
    XlideFileSystemProvider,
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    decodeModuleUri,
    isLocalXlideDocument,
} from './xlideFileSystem';
import { WorkbookEngine } from './workbookEngine';
import { registerFormPreview } from './vbaFormPreview';
import { registerAgentTools } from './agentTools';
import { registerCommands } from './commands';
import { registerVbaLanguageProviders } from './vbaLanguageProviders';
import { XlideStatusBar } from './statusBar';
import { registerXlideDirtyModuleBackups } from './xlideDirtyModuleBackups';
import { registerVbaEditorCommands } from './vbaEditorCommands';
import { registerXlideCommand } from './xlideCommandRegistration';
import { createRecordedOutputChannel } from './xlideOutputLog';
import { setExcelCoordinationLog } from './excelWorkbookCoordinator';
import { registerXlideGlobalSettingsWebview } from './globalSettingsWebview';
import {
    xlideExplorerAutoExpandCollapseFromConfig,
    xlidePerformanceTraceFromConfig,
} from './globalSettings';
import { registerXlideSidebar } from './xlideSidebar';
import { AnalysisWorkerClient } from './analysisWorkerClient';
import { setExtensionAssetRoot } from './extensionAssets';
import { cleanupStaleVbaTestHostTempDirsAsync } from './vbaTestTempFiles';
import { setPerformanceTraceLogger } from './performanceTrace';
import { setWorkbookAnalysisWorker } from './vbaWorkbookAnalysis';
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

    const bridge = new WorkbookEngine(context, out);
    const fsProvider = new XlideFileSystemProvider(bridge);
    registerFormPreview(context, bridge, fsProvider);
    const explorer = new XlsmExplorer(bridge, out);
    const statusBar = new XlideStatusBar();
    // The workbook engine runs in-process: there is no backend to install,
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
    setWorkbookAnalysisWorker(analysisWorkerClient);
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
            let pending: { xlsmPath: string; moduleName: string } | undefined;
            const apply = debounce(() => {
                if (!pending) { return; }
                const { xlsmPath, moduleName } = pending;
                pending = undefined;
                // Honor the user's auto-expand/collapse preference: when off, the
                // tree never follows the active tab or accordion-collapses.
                if (!xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) {
                    return;
                }
                explorer.setActiveModule(xlsmPath, moduleName);
                const node = explorer.getModuleNode(xlsmPath, moduleName);
                if (node && treeView.visible) {
                    void treeView.reveal(node, { select: true, focus: false, expand: true });
                }
            }, 60);
            const subscription = vscode.window.onDidChangeActiveTextEditor((editor) => {
                // Focus moved off any text editor (the Output panel, terminal, the
                // tree, a webview, or the last tab closed). Leave the tree as-is: a
                // workbook only collapses when focus moves to a module in a DIFFERENT
                // workbook, never on transient focus loss.
                if (!editor) {
                    pending = undefined;
                    apply.cancel();
                    return;
                }
                if (!isLocalXlideDocument(editor.document)) { return; }
                pending = decodeModuleUri(editor.document.uri);
                apply();
            });
            return new vscode.Disposable(() => {
                apply.dispose();
                subscription.dispose();
            });
        })(),

        // Accordion: if the user manually clicks the expand arrow on a module node,
        // collapse all sibling modules under the same workbook.
        treeView.onDidExpandElement((e) => {
            if (!xlideExplorerAutoExpandCollapseFromConfig(vscode.workspace.getConfiguration('xlide')).value) { return; }
            if (e.element.kind === 'module' && e.element.filePath && e.element.moduleName) {
                explorer.setActiveModule(e.element.filePath, e.element.moduleName);
            }
        }),

        // When the user manually collapses the active workbook, stop forcing it
        // Expanded so it stays collapsed - otherwise the next refresh re-stamps it
        // Expanded against the still-set active-workbook key and springs it open.
        treeView.onDidCollapseElement((e) => {
            if (e.element.kind === 'xlsm' && e.element.filePath) {
                explorer.notifyWorkbookCollapsed(e.element.filePath);
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

        ...sidebar.disposables,
        registerXlideGlobalSettingsWebview(out),
        ...registerAgentTools(context, bridge, explorer, fsProvider, vbaIndex),

        statusBar,
        bridge,
    );

    // When the symbol index updates (e.g. after a rename or save), refresh
    // the matching module's sub list in the explorer so renamed procedures
    // appear immediately.
    context.subscriptions.push(
        vbaIndex.onDidChange(({ xlsmPath, moduleName }) => {
            if (!xlsmPath || !moduleName) {
                explorer.refresh();
            } else {
                explorer.refreshModuleSubs(xlsmPath, moduleName);
            }
        }),
    );

    // Item 7: Auto-expand the first workbook so modules are visible. Listing
    // workbooks only globs the workspace; expanding the first one performs the
    // first bridge call, which lazy-starts the backend. Workspaces without
    // Excel workbooks (or with the explorer hidden) never open a workbook.
    if (treeView.visible) {
        const autoExpandTimer = setTimeout(() => {
            if (!treeView.visible) { return; }
            void explorer.warmXlsmCache().then(firstNode => {
                if (firstNode && treeView.visible) {
                    void treeView.reveal(firstNode, { select: false, focus: false, expand: true });
                }
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

