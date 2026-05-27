import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as path from 'path';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider, XLIDE_SCHEME, XLIDE_LIVESHARE_AUTHORITY, decodeModuleUri } from './xlideFileSystem';
import { PythonBridge } from './pythonBridge';
import { registerAgentTools } from './agentTools';
import { registerCommands } from './commands';
import { registerVbaLanguageProviders } from './vbaLanguageProviders';
import { LiveShareIntegration } from './liveShare';
import { XlideStatusBar } from './statusBar';

// ---------------------------------------------------------------------------
// Dependency installer
// ---------------------------------------------------------------------------

function installDependencies(
    bridge: PythonBridge,
    context: vscode.ExtensionContext,
    out: vscode.OutputChannel,
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
                            void vscode.window.showInformationMessage(
                                'XLIDE: Dependencies installed and bridge started. If any files failed to open, click Try Again in the editor tab.',
                            );
                        })
                        .catch((err: Error) => {
                            out.appendLine(`ERROR after install: ${err.message}`);
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
    const out = vscode.window.createOutputChannel('XLIDE');
    out.appendLine('XLIDE activating...');

    const bridge = new PythonBridge(context, out);
    const fsProvider = new XlideFileSystemProvider(bridge);
    const explorer = new XlsmExplorer(bridge);
    const liveShare = new LiveShareIntegration(bridge, out);
    fsProvider.setLiveShare(liveShare);
    explorer.setLiveShare(liveShare);
    const statusBar = new XlideStatusBar(liveShare);

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

    context.subscriptions.push(
        out,

        // Virtual read/write filesystem for xlide-vba:// URIs
        vscode.workspace.registerFileSystemProvider(XLIDE_SCHEME, fsProvider, {
            isCaseSensitive: process.platform !== 'win32',
            isReadonly: false,
        }),

        treeView,

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
                if (node) {
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
        vscode.commands.registerCommand('xlide.previewErrors', async () => {
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
                await vscode.window.showErrorMessage(
                    'XLIDE: Required Python packages are missing (pyOpenVBA, openpyxl). Click "Install Now" to install them automatically.',
                    'Install Now', 'View XLIDE Output', 'Dismiss',
                );
            } else if (pick.id === 'c') {
                void vscode.window.showInformationMessage(
                    'XLIDE: Dependencies installed and bridge started. If any files failed to open, click Try Again in the editor tab.',
                );
            }
        }),

        // Manual setup command (also auto-triggered on missing packages)
        vscode.commands.registerCommand('xlide.setup', () =>
            installDependencies(bridge, context, out).catch((err: Error) => {
                out.appendLine(`Setup error: ${err.message}`);
                vscode.window.showErrorMessage(`XLIDE setup failed: ${err.message}`);
            }),
        ),

        // Show the XLIDE output channel (used by the explorer welcome view).
        vscode.commands.registerCommand('xlide.showOutput', () => {
            out.show(true);
        }),

        ...registerCommands(context, bridge, explorer, fsProvider, out),
        ...registerAgentTools(context, bridge, explorer, fsProvider),

        statusBar,
        liveShare,
        bridge,
    );

    // Initialize Live Share integration (no-op if extension isn't installed)
    void liveShare.start().catch((err: Error) => {
        out.appendLine(`Live Share init failed: ${err.message}`);
    });

    // VBA language services: syntax-aware symbol index + providers
    const vbaIndex = registerVbaLanguageProviders(context, bridge);

    // When the symbol index updates (e.g. after a rename or save), refresh
    // the matching module's sub list in the explorer so renamed procedures
    // appear immediately.
    context.subscriptions.push(
        vbaIndex.onDidChange(({ xlsmPath, moduleName }) => {
            if (!xlsmPath) {
                explorer.refresh();
            } else if (moduleName) {
                explorer.refreshModuleSubs(xlsmPath, moduleName);
            }
        }),
    );

    const isMissingPackage = (msg: string) =>
        /No module named|ModuleNotFoundError|ImportError/i.test(msg);

    const isPythonNotFound = (msg: string) =>
        /python.*not found|not recognized|cannot find|no such file|ENOENT|spawn.*python/i.test(msg);

    bridge.start().then(() => {
        out.appendLine('XLIDE ready.');

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

        // Item 7: Auto-expand the first workbook on activation so modules are visible.
        void explorer.warmXlsmCache().then(firstNode => {
            if (firstNode) {
                void treeView.reveal(firstNode, { select: false, focus: false, expand: true });
            }
        });
    }).catch(async (err: Error) => {
        out.appendLine(`ERROR: Python backend failed to start - ${err.message}`);

        if (isPythonNotFound(err.message)) {
            const choice = await vscode.window.showErrorMessage(
                'XLIDE: Python 3.10+ was not found. Install Python and tick "Add Python to PATH", ' +
                'then reload the window. Or set xlide.pythonPath to your Python executable and reload.',
                'Get Python',
                'Set Python Path',
                'Reload Window',
            );
            if (choice === 'Get Python') {
                void vscode.env.openExternal(vscode.Uri.parse('https://www.python.org/downloads/'));
                void vscode.window.showInformationMessage(
                    'After installing Python, reload the window to start XLIDE.',
                    'Reload Window',
                ).then(action => {
                    if (action === 'Reload Window') {
                        void vscode.commands.executeCommand('workbench.action.reloadWindow');
                    }
                });
            } else if (choice === 'Set Python Path') {
                void vscode.commands.executeCommand('workbench.action.openSettings', 'xlide.pythonPath');
                void vscode.window.showInformationMessage(
                    'After setting the path, reload the window to start XLIDE.',
                    'Reload Window',
                ).then(action => {
                    if (action === 'Reload Window') {
                        void vscode.commands.executeCommand('workbench.action.reloadWindow');
                    }
                });
            } else if (choice === 'Reload Window') {
                void vscode.commands.executeCommand('workbench.action.reloadWindow');
            }
        } else if (isMissingPackage(err.message)) {
            const choice = await vscode.window.showErrorMessage(
                'XLIDE: Required Python packages are missing (pyOpenVBA, openpyxl). ' +
                'Click "Install Now" to install them automatically.',
                'Install Now',
                'View XLIDE Output',
                'Dismiss',
            );
            if (choice === 'Install Now') {
                await installDependencies(bridge, context, out).catch((e: Error) => {
                    out.appendLine(`Setup error: ${e.message}`);
                    vscode.window.showErrorMessage(`XLIDE setup failed: ${e.message}. See the XLIDE output channel for details.`);
                });
            } else if (choice === 'View XLIDE Output') {
                out.show(true);
            }
        } else {
            const choice = await vscode.window.showErrorMessage(
                `XLIDE: Failed to start Python backend. ${err.message}`,
                'View XLIDE Output',
                'Set Python Path',
            );
            if (choice === 'View XLIDE Output') {
                out.show(true);
            } else if (choice === 'Set Python Path') {
                void vscode.commands.executeCommand('workbench.action.openSettings', 'xlide.pythonPath');
            }
        }
    });
}

export function deactivate(): void { /* nothing async needed */ }
