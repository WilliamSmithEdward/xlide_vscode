import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as path from 'path';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider, XLIDE_SCHEME } from './xlideFileSystem';
import { PythonBridge } from './pythonBridge';
import { registerAgentTools } from './agentTools';
import { registerCommands } from './commands';
import { registerVbaLanguageProviders } from './vbaLanguageProviders';
import { LiveShareIntegration } from './liveShare';

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

    context.subscriptions.push(
        out,

        // Virtual read/write filesystem for xlide-vba:// URIs
        vscode.workspace.registerFileSystemProvider(XLIDE_SCHEME, fsProvider, {
            isCaseSensitive: process.platform !== 'win32',
            isReadonly: false,
        }),

        // Tree view in the Explorer sidebar
        vscode.window.createTreeView('xlide.explorer', {
            treeDataProvider: explorer,
            showCollapseAll: true,
        }),

        // Refresh the explorer when .xlsm/.xlsb/.xlam files are added or removed
        (() => {
            const watcher = vscode.workspace.createFileSystemWatcher('**/*.{xlsm,xlsb,xlam}');
            watcher.onDidCreate(() => explorer.refresh());
            watcher.onDidDelete(() => explorer.refresh());
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

        ...registerCommands(context, bridge, explorer, fsProvider, out),
        ...registerAgentTools(context, bridge, explorer, fsProvider),

        liveShare,
        bridge,
    );

    // Initialize Live Share integration (no-op if extension isn't installed)
    void liveShare.start().catch((err: Error) => {
        out.appendLine(`Live Share init failed: ${err.message}`);
    });

    // VBA language services: syntax-aware symbol index + providers
    registerVbaLanguageProviders(context, bridge);

    const isMissingPackage = (msg: string) =>
        /No module named|ModuleNotFoundError|ImportError/i.test(msg);

    const isPythonNotFound = (msg: string) =>
        /python.*not found|not recognized|cannot find|no such file|ENOENT|spawn.*python/i.test(msg);

    bridge.start().then(() => {
        out.appendLine('XLIDE ready.');
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
