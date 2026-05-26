import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as path from 'path';
import { XlsmExplorer } from './xlsmExplorer';
import { XlideFileSystemProvider, XLIDE_SCHEME } from './xlideFileSystem';
import { PythonBridge } from './pythonBridge';
import { registerAgentTools } from './agentTools';
import { registerCommands } from './commands';
import { registerVbaLanguageProviders } from './vbaLanguageProviders';

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
                    vscode.window.showInformationMessage('XLIDE: Dependencies installed. Starting...');
                    bridge.start()
                        .then(() => out.appendLine('XLIDE ready.'))
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

        // Manual setup command (also auto-triggered on missing packages)
        vscode.commands.registerCommand('xlide.setup', () =>
            installDependencies(bridge, context, out).catch((err: Error) => {
                out.appendLine(`Setup error: ${err.message}`);
                vscode.window.showErrorMessage(`XLIDE setup failed: ${err.message}`);
            }),
        ),

        ...registerCommands(context, bridge, explorer, fsProvider, out),
        ...registerAgentTools(context, bridge, explorer, fsProvider),

        bridge,
    );

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
                'XLIDE: Python was not found on your system. ' +
                'Install Python 3.9+ and tick "Add Python to PATH", or set the xlide.pythonPath setting to point at your Python executable.',
                'Get Python',
                'Set Python Path',
                'View XLIDE Output',
            );
            if (choice === 'Get Python') {
                void vscode.env.openExternal(vscode.Uri.parse('https://www.python.org/downloads/'));
            } else if (choice === 'Set Python Path') {
                void vscode.commands.executeCommand('workbench.action.openSettings', 'xlide.pythonPath');
            } else if (choice === 'View XLIDE Output') {
                out.show(true);
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
