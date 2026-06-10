import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as path from 'path';
import { XlsmExplorer } from './xlsmExplorer';
import {
    XlideFileSystemProvider,
    XLIDE_SCHEME,
    XLIDE_VBA_LANGUAGE_ID,
    XLIDE_LIVESHARE_AUTHORITY,
    decodeModuleUri,
} from './xlideFileSystem';
import { PythonBridge } from './pythonBridge';
import { registerAgentTools } from './agentTools';
import { registerCommands } from './commands';
import { registerVbaLanguageProviders } from './vbaLanguageProviders';
import { LiveShareIntegration } from './liveShare';
import { XlideStatusBar } from './statusBar';
import { registerXlideDirtyModuleBackups } from './xlideDirtyModuleBackups';
import { registerVbaEditorCommands } from './vbaEditorCommands';
import { registerXlideCommand } from './xlideCommandRegistration';
import { createRecordedOutputChannel } from './xlideOutputLog';
import { registerXlideGlobalSettingsWebview } from './globalSettingsWebview';
import {
    setXlideGlobalSettingValue,
    xlidePerformanceTraceFromConfig,
    xlidePythonPathFromConfig,
} from './globalSettings';
import { registerXlideSidebar } from './xlideSidebar';
import { isXlideSetupComplete, type XlideSidebarSetupStatus } from './xlideSidebarModel';
import { cleanupStaleVbaTestHostTempDirsAsync } from './vbaTestTempFiles';
import { setPerformanceTraceLogger } from './performanceTrace';
import { XLIDE_VBA_EDITOR_OVERRIDES } from './xlideVbaEditorOverrides';

const PYTHON_DOWNLOAD_URL = 'https://www.python.org/downloads/';

// ---------------------------------------------------------------------------
// Dependency installer
// ---------------------------------------------------------------------------

function installDependencies(
    bridge: PythonBridge,
    context: vscode.ExtensionContext,
    out: vscode.OutputChannel,
    onBridgeReady?: () => void,
    onBridgeFailed?: (err: Error) => void,
): Promise<void> {
    const pythonPath = bridge.resolvePython();
    const requirementsPath = path.join(context.extensionPath, 'python', 'requirements.txt');

    return Promise.resolve(vscode.window.withProgress(
        { location: vscode.ProgressLocation.Notification, title: 'XLIDE: Installing Python dependencies...', cancellable: false },
        () => new Promise<void>((resolve, reject) => {
            out.appendLine(`Running: ${pythonPath} -m pip install -r ${requirementsPath}`);
            const proc = cp.spawn(pythonPath, ['-m', 'pip', 'install', '-r', requirementsPath], {
                stdio: ['ignore', 'pipe', 'pipe'],
            });
            proc.stdout!.on('data', (d: Buffer) => out.appendLine(d.toString().trimEnd()));
            proc.stderr!.on('data', (d: Buffer) => out.appendLine(d.toString().trimEnd()));
            proc.on('error', (err) => reject(new Error(`pip failed: ${err.message}`)));
            proc.on('exit', (code) => {
                if (code === 0) {
                    bridge.start()
                        .then(() => {
                            out.appendLine('XLIDE ready.');
                            onBridgeReady?.();
                            void vscode.window.showInformationMessage(
                                'XLIDE: Dependencies installed and bridge started. If any files failed to open, click Try Again in the editor tab.',
                            );
                        })
                        .catch((err: Error) => {
                            out.appendLine(`ERROR after install: ${err.message}`);
                            onBridgeFailed?.(err);
                            vscode.window.showErrorMessage(`XLIDE: ${err.message}`);
                        });
                    resolve();
                } else {
                    reject(new Error(`pip install exited with code ${code}. See XLIDE output for details.`));
                }
            });
        }),
    ));
}

// ---------------------------------------------------------------------------
// Activation
// ---------------------------------------------------------------------------

export function activate(context: vscode.ExtensionContext): void {
    const out = createRecordedOutputChannel(vscode.window.createOutputChannel('XLIDE'));
    setPerformanceTraceLogger(
        (line) => out.appendLine(line),
        () => xlidePerformanceTraceFromConfig(vscode.workspace.getConfiguration('xlide')).value,
    );
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
            const message = err instanceof Error ? err.message : String(err);
            out.appendLine(`VBA test temp cleanup skipped: ${message}`);
        });

    const bridge = new PythonBridge(context, out);
    const fsProvider = new XlideFileSystemProvider(bridge);
    const explorer = new XlsmExplorer(bridge);
    const liveShare = new LiveShareIntegration(bridge, out);
    fsProvider.setLiveShare(liveShare);
    explorer.setLiveShare(liveShare);
    const statusBar = new XlideStatusBar(liveShare);
    const isMissingPackage = (msg: string) =>
        /No module named|ModuleNotFoundError|ImportError/i.test(msg);

    const isPythonNotFound = (msg: string) =>
        /python.*not found|not recognized|cannot find|no such file|ENOENT|spawn.*python/i.test(msg);

    const configuredPythonPath = () =>
        xlidePythonPathFromConfig(vscode.workspace.getConfiguration('xlide')).value;

    const pythonLauncherDetectsPython = (): Promise<boolean> => {
        if (process.platform !== 'win32') {
            return Promise.resolve(false);
        }
        return new Promise<boolean>((resolve) => {
            const proc = cp.spawn('py', ['-0p'], {
                windowsHide: true,
            });
            let stdout = '';
            let stderr = '';
            let settled = false;
            const finish = (value: boolean) => {
                if (settled) { return; }
                settled = true;
                clearTimeout(timer);
                resolve(value);
            };
            const timer = setTimeout(() => {
                finish(false);
                proc.kill();
            }, 1500);
            proc.stdout?.on('data', (chunk: Buffer) => { stdout += chunk.toString(); });
            proc.stderr?.on('data', (chunk: Buffer) => { stderr += chunk.toString(); });
            proc.on('error', () => finish(false));
            proc.on('exit', (code) => {
                finish(code === 0 && /python(?:\.exe)?/i.test(`${stdout}\n${stderr}`));
            });
        });
    };

    const checkingSetupStatus = (): XlideSidebarSetupStatus => ({
        pythonExecutable: {
            status: 'unknown',
            description: 'Checking',
            tooltip: 'XLIDE is checking the configured Python executable.',
        },
        pythonLibraries: {
            status: 'unknown',
            description: 'Checking',
            tooltip: 'XLIDE is checking required Python libraries.',
        },
    });
    let setupStatus: XlideSidebarSetupStatus = checkingSetupStatus();
    const sidebar = registerXlideSidebar({
        setupStatus: () => setupStatus,
        workspaceState: context.workspaceState,
    });
    const setSetupStatus = (status: XlideSidebarSetupStatus) => {
        setupStatus = status;
        const setupComplete = isXlideSetupComplete(status);
        explorer.setSetupComplete(setupComplete);
        void vscode.commands.executeCommand('setContext', 'xlide.setupComplete', setupComplete);
        sidebar.refresh();
    };
    setSetupStatus(setupStatus);
    const pythonBackendReady = () => setSetupStatus({
        pythonExecutable: {
            status: 'pass',
            description: bridge.resolvePython(),
            tooltip: 'XLIDE found a usable Python executable.',
        },
        pythonLibraries: {
            status: 'pass',
            description: 'Installed',
            tooltip: 'Required Python libraries are installed.',
        },
    });
    const pythonBackendNeedsAttention = async (err: Error): Promise<void> => {
        if (isPythonNotFound(err.message)) {
            const configured = configuredPythonPath();
            const installedOutsidePath = !configured && await pythonLauncherDetectsPython();
            const shouldSetPath = Boolean(configured) || installedOutsidePath;
            setSetupStatus({
                pythonExecutable: {
                    status: 'warn',
                    description: configured
                        ? 'Path Not Found'
                        : installedOutsidePath
                            ? 'Not On PATH'
                            : 'Not Found',
                    tooltip: shouldSetPath
                        ? 'Python appears to be installed, but XLIDE cannot start it from the current path. Set xlide.pythonPath to the Python executable.'
                        : err.message,
                    action: shouldSetPath ? 'setPythonPath' : 'downloadPython',
                },
                pythonLibraries: {
                    status: 'unknown',
                    description: 'Waiting For Python',
                    tooltip: 'Set a valid Python executable before installing required libraries.',
                },
            });
            return;
        }
        if (isMissingPackage(err.message)) {
            setSetupStatus({
                pythonExecutable: {
                    status: 'pass',
                    description: bridge.resolvePython(),
                    tooltip: 'XLIDE found a usable Python executable.',
                },
                pythonLibraries: {
                    status: 'warn',
                    description: 'Missing',
                    tooltip: err.message,
                },
            });
            return;
        }
        setSetupStatus({
            pythonExecutable: {
                status: 'unknown',
                description: 'Check Settings',
                tooltip: err.message,
            },
            pythonLibraries: {
                status: 'warn',
                description: 'Needs Attention',
                tooltip: err.message,
            },
        });
    };
    const recheckPythonBackend = () => {
        setSetupStatus(checkingSetupStatus());
        void bridge.restart()
            .then(() => {
                out.appendLine('XLIDE ready after Python path change.');
                pythonBackendReady();
            })
            .catch((err: Error) => {
                out.appendLine(`ERROR: Python backend failed after path change - ${err.message}`);
                void pythonBackendNeedsAttention(err);
            });
    };

    // Mirror Live Share guest state into a context key so the explorer welcome view
    // can show a "not supported" message instead of the generic empty-workspace one.
    const updateGuestContext = () => {
        void vscode.commands.executeCommand(
            'setContext',
            'xlide.isLiveShareGuest',
            liveShare.isInGuestSession,
        );
    };
    updateGuestContext();
    liveShare.onDidChange(updateGuestContext);

    // Keep reference outside subscriptions for post-start auto-expand and reveal.
    const treeView = vscode.window.createTreeView('xlide.explorer', {
        treeDataProvider: explorer,
        showCollapseAll: true,
    });

    // VBA language services: syntax-aware symbol index + providers.
    const vbaIndex = registerVbaLanguageProviders(context, bridge);
    registerVbaEditorCommands(context);
    registerXlideVbaLanguageSync(context, out);
    void ensureXlideVbaEditorOverrides(out);

    context.subscriptions.push(
        out,

        // Virtual read/write filesystem for xlide-vba:// URIs
        vscode.workspace.registerFileSystemProvider(XLIDE_SCHEME, fsProvider, {
            isCaseSensitive: process.platform !== 'win32',
            isReadonly: false,
        }),
        registerXlideDirtyModuleBackups(context, out),

        treeView,

        vscode.workspace.onDidChangeConfiguration((event) => {
            if (event.affectsConfiguration('xlide.pythonPath')) {
                out.appendLine('XLIDE Python path changed; rechecking Python backend.');
                recheckPythonBackend();
            }
        }),

        // Item 6: Reveal active module in the XLIDE Explorer tree.
        // Also drives accordion collapse: only the active module stays expanded.
        // Debounced so rapid tab switches (e.g. Ctrl+W spam) coalesce into a
        // single setActiveModule + reveal, avoiding overlapping async reveal
        // calls that could leave stale modules expanded.
        (() => {
            let timer: ReturnType<typeof setTimeout> | undefined;
            let pending: { xlsmPath: string; moduleName: string } | undefined;
            const apply = () => {
                timer = undefined;
                if (!pending) { return; }
                const { xlsmPath, moduleName } = pending;
                pending = undefined;
                explorer.setActiveModule(xlsmPath, moduleName);
                const node = explorer.getModuleNode(xlsmPath, moduleName);
                if (node && treeView.visible) {
                    void treeView.reveal(node, { select: true, focus: false, expand: true });
                }
            };
            return vscode.window.onDidChangeActiveTextEditor((editor) => {
                // No active editor (e.g. user closed the last tab) — collapse all modules.
                if (!editor) {
                    pending = undefined;
                    if (timer !== undefined) { clearTimeout(timer); timer = undefined; }
                    explorer.clearActiveModule();
                    return;
                }
                const uri = editor.document.uri;
                if (uri.scheme !== XLIDE_SCHEME || uri.authority === XLIDE_LIVESHARE_AUTHORITY) { return; }
                pending = decodeModuleUri(uri);
                if (timer !== undefined) { clearTimeout(timer); }
                timer = setTimeout(apply, 60);
            });
        })(),

        // Accordion: if the user manually clicks the expand arrow on a module node,
        // collapse all sibling modules under the same workbook.
        treeView.onDidExpandElement((e) => {
            if (e.element.kind === 'module' && e.element.filePath && e.element.moduleName) {
                explorer.setActiveModule(e.element.filePath, e.element.moduleName);
            }
        }),

        // Refresh the explorer when .xlsm/.xlsb/.xlam files are added or removed
        // Debounced so rapid file-system events (save storms) coalesce into one refresh.
        (() => {
            let timer: ReturnType<typeof setTimeout> | undefined;
            const debouncedRefresh = () => {
                if (timer !== undefined) { clearTimeout(timer); }
                timer = setTimeout(() => { timer = undefined; explorer.refresh(); }, 200);
            };
            const watcher = vscode.workspace.createFileSystemWatcher('**/*.{xlsm,xlsb,xlam}');
            watcher.onDidCreate(debouncedRefresh);
            watcher.onDidDelete(debouncedRefresh);
            return watcher;
        })(),

        // DEV ONLY: preview error notification UX
        registerXlideCommand('xlide.previewErrors', async () => {
            const pick = await vscode.window.showQuickPick([
                { label: 'Scenario A: Python not found', id: 'a' },
                { label: 'Scenario B: Packages missing', id: 'b' },
                { label: 'After install success', id: 'c' },
            ], { title: 'XLIDE: Preview error notification' });
            if (!pick) { return; }
            if (pick.id === 'a') {
                const choice = await vscode.window.showErrorMessage(
                    'XLIDE: Python 3.10+ was not found. Install Python and tick "Add Python to PATH", then reload the window. Or set xlide.pythonPath to your Python executable and reload.',
                    'Get Python', 'Set Python Path', 'Reload Window',
                );
                if (choice === 'Reload Window') {
                    void vscode.commands.executeCommand('workbench.action.reloadWindow');
                } else if (choice === 'Get Python') {
                    void vscode.env.openExternal(vscode.Uri.parse('https://www.python.org/downloads/'));
                    void vscode.window.showInformationMessage(
                        'After installing Python, reload the window to start XLIDE.',
                        'Reload Window',
                    ).then(a => { if (a === 'Reload Window') { void vscode.commands.executeCommand('workbench.action.reloadWindow'); } });
                } else if (choice === 'Set Python Path') {
                    void vscode.commands.executeCommand('workbench.action.openSettings', 'xlide.pythonPath');
                    void vscode.window.showInformationMessage(
                        'After setting the path, reload the window to start XLIDE.',
                        'Reload Window',
                    ).then(a => { if (a === 'Reload Window') { void vscode.commands.executeCommand('workbench.action.reloadWindow'); } });
                }
            } else if (pick.id === 'b') {
                const choice = await vscode.window.showErrorMessage(
                    'XLIDE: Required Python packages are missing (pyOpenVBA, openpyxl). Click "Install Now" to install them automatically.',
                    'Install Now', 'Copy Diagnostics', 'Dismiss',
                );
                if (choice === 'Copy Diagnostics') {
                    void vscode.commands.executeCommand('xlide.copyDiagnostics');
                }
            } else if (pick.id === 'c') {
                void vscode.window.showInformationMessage(
                    'XLIDE: Dependencies installed and bridge started. If any files failed to open, click Try Again in the editor tab.',
                );
            }
        }),

        // Manual setup command surfaced by the sidebar setup row.
        registerXlideCommand('xlide.setup', () =>
            installDependencies(bridge, context, out, pythonBackendReady, pythonBackendNeedsAttention).catch((err: Error) => {
                void pythonBackendNeedsAttention(err);
                out.appendLine(`Setup error: ${err.message}`);
                void vscode.window.showErrorMessage(
                    `XLIDE setup failed: ${err.message}`,
                    'Copy Diagnostics',
                ).then((choice) => {
                    if (choice === 'Copy Diagnostics') {
                        void vscode.commands.executeCommand('xlide.copyDiagnostics');
                    }
                });
            }),
        ),

        registerXlideCommand('xlide.downloadPython', () => {
            void vscode.env.openExternal(vscode.Uri.parse(PYTHON_DOWNLOAD_URL));
        }),

        registerXlideCommand('xlide.browsePythonPath', async () => {
            const configured = configuredPythonPath();
            const selected = await vscode.window.showOpenDialog({
                title: 'XLIDE: Select Python Executable',
                openLabel: 'Use Python',
                canSelectFiles: true,
                canSelectFolders: false,
                canSelectMany: false,
                defaultUri: configured ? vscode.Uri.file(configured) : undefined,
                filters: process.platform === 'win32'
                    ? { 'Python Executable': ['exe'], 'All Files': ['*'] }
                    : undefined,
            });
            const pythonPath = selected?.[0]?.fsPath;
            if (!pythonPath) {
                return;
            }
            await setXlideGlobalSettingValue(
                vscode.workspace.getConfiguration('xlide'),
                'pythonPath',
                pythonPath,
            );
        }),

        // Show the XLIDE output channel (used by the explorer welcome view).
        registerXlideCommand('xlide.showOutput', () => {
            out.show(true);
        }),

        ...sidebar.disposables,
        registerXlideGlobalSettingsWebview(),
        ...registerCommands(context, bridge, explorer, fsProvider, out, vbaIndex),
        ...registerAgentTools(context, bridge, explorer, fsProvider),

        statusBar,
        liveShare,
        bridge,
    );

    // Initialize Live Share integration (no-op if extension isn't installed)
    void liveShare.start().catch((err: Error) => {
        out.appendLine(`Live Share init failed: ${err.message}`);
    });

    // When the symbol index updates (e.g. after a rename or save), refresh
    // the matching module's sub list in the explorer so renamed procedures
    // appear immediately.
    context.subscriptions.push(
        vbaIndex.onDidChange(({ xlsmPath, moduleName }) => {
            if (!xlsmPath || !moduleName) {
                explorer.refresh();
            } else if (moduleName) {
                explorer.refreshModuleSubs(xlsmPath, moduleName);
            }
        }),
    );

    bridge.start().then(() => {
        out.appendLine('XLIDE ready.');
        pythonBackendReady();

        // Item 9: Show a one-time welcome notification on first ever activation.
        if (!context.globalState.get('xlide.welcomed')) {
            void context.globalState.update('xlide.welcomed', true);
            void vscode.window.showInformationMessage(
                'XLIDE is ready. Right-click a workbook in the XLIDE Explorer to export modules, ' +
                'or press F5 inside a module to run the macro at the cursor.',
                'Open Explorer',
            ).then(choice => {
                if (choice === 'Open Explorer') {
                    void vscode.commands.executeCommand('xlide.explorer.focus');
                }
            });
        }

        // Recommend disabling AI inline (ghost-text) completions for VBA, which
        // can hide XLIDE's IntelliSense suggestion menu. One-time, opt-in.
        recommendDisableInlineSuggest(context, out);

        // Item 7: Auto-expand the first workbook on activation so modules are visible.
        if (treeView.visible) {
            setTimeout(() => {
                if (!treeView.visible) { return; }
                void explorer.warmXlsmCache().then(firstNode => {
                    if (firstNode && treeView.visible) {
                        void treeView.reveal(firstNode, { select: false, focus: false, expand: true });
                    }
                });
            }, 250);
        }
    }).catch(async (err: Error) => {
        out.appendLine(`ERROR: Python backend failed to start - ${err.message}`);
        await pythonBackendNeedsAttention(err);

        if (isPythonNotFound(err.message) || isMissingPackage(err.message)) {
            out.appendLine('XLIDE setup is incomplete; use the XLIDE sidebar Setup section to finish Python setup.');
            return;
        }

        const choice = await vscode.window.showErrorMessage(
            `XLIDE: Failed to start Python backend. ${err.message}`,
            'Copy Diagnostics',
            'Set Python Path',
        );
        if (choice === 'Copy Diagnostics') {
            void vscode.commands.executeCommand('xlide.copyDiagnostics');
        } else if (choice === 'Set Python Path') {
            void vscode.commands.executeCommand('workbench.action.openSettings', 'xlide.pythonPath');
        }
    });
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
            const message = err instanceof Error ? err.message : String(err);
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

/**
 * One-time recommendation to disable AI inline (ghost-text) completions for XLIDE
 * VBA modules, which can visually obscure XLIDE's IntelliSense suggestion menu. Only
 * shown when inline suggestions are still effectively enabled for XLIDE VBA and the
 * user has not been asked before.
 */
function recommendDisableInlineSuggest(
    context: vscode.ExtensionContext,
    out: vscode.OutputChannel,
): void {
    if (context.globalState.get('xlide.inlineSuggestRecommended')) {
        return;
    }

    const config = vscode.workspace.getConfiguration('editor', { languageId: XLIDE_VBA_LANGUAGE_ID });
    const inspected = config.inspect<boolean>('inlineSuggest.enabled');
    const vbaOverride =
        inspected?.globalLanguageValue ??
        inspected?.workspaceLanguageValue ??
        inspected?.workspaceFolderLanguageValue;
    if (vbaOverride === false) {
        return; // Already disabled for XLIDE VBA; nothing to recommend.
    }
    if (config.get<boolean>('inlineSuggest.enabled', true) === false) {
        return; // Inline suggestions are off everywhere; no conflict to resolve.
    }

    void context.globalState.update('xlide.inlineSuggestRecommended', true);
    void vscode.window.showInformationMessage(
        'XLIDE provides VBA IntelliSense. AI inline completions (gray ghost text) can hide its ' +
        'suggestion menu. Disable inline completions for XLIDE VBA modules?',
        'Disable for XLIDE',
        'Keep',
    ).then(choice => {
        if (choice !== 'Disable for XLIDE') {
            return;
        }
        vscode.workspace
            .getConfiguration('editor', { languageId: XLIDE_VBA_LANGUAGE_ID })
            .update('inlineSuggest.enabled', false, vscode.ConfigurationTarget.Global, true)
            .then(
                () => out.appendLine(`Disabled editor.inlineSuggest.enabled for [${XLIDE_VBA_LANGUAGE_ID}].`),
                (err: Error) => out.appendLine(`Could not update setting: ${err.message}`),
            );
    });
}

