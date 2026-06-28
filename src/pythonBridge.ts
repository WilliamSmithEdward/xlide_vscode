import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as path from 'path';
import * as readline from 'readline';
import { xlidePythonPathFromConfig } from './globalSettings';
import { isCancellationLike, startPerformanceTrace } from './performanceTrace';
import { BridgeError } from './pythonBridgeErrors';

interface JsonRpcRequest {
    jsonrpc: '2.0';
    id: number;
    method: string;
    params: unknown;
}

interface JsonRpcResponse {
    jsonrpc: '2.0';
    id: number;
    result?: unknown;
    error?: { code: number; message: string };
}

interface PendingRequest {
    resolve: (value: unknown) => void;
    reject: (err: Error) => void;
    cancel?: vscode.Disposable;
    timer?: ReturnType<typeof setTimeout>;
}

export interface PythonBridgeCallOptions {
    /** Per-request watchdog; defaults to DEFAULT_BRIDGE_CALL_TIMEOUT_MS. Pass 0 to disable. */
    timeoutMs?: number;
}

// Stderr is mirrored to the output channel; keep only the most recent lines
// for the startup-failure message so the buffer cannot grow unbounded.
const MAX_STDERR_LINES = 50;

// Generous default so long workbook operations (runOpenpyxl, large module
// writes) still finish, while a hung single-threaded server cannot wedge
// callers forever (python/server.py handles one request at a time).
export const DEFAULT_BRIDGE_CALL_TIMEOUT_MS = 120_000;

export class PythonBridge implements vscode.Disposable {
    private _proc: cp.ChildProcess | undefined;
    private _pending = new Map<number, PendingRequest>();
    private _nextId = 1;
    private _ready = false;
    private _startPromise: Promise<void> | undefined;
    private _readyResolve: (() => void) | undefined;
    private _readyReject: ((err: Error) => void) | undefined;
    private _stderrLines: string[] = [];
    private _stopping = false;
    private _onDemandStart: (() => Promise<void>) | undefined;

    constructor(
        private readonly _context: vscode.ExtensionContext,
        private readonly _out: vscode.OutputChannel,
    ) {}

    /** Exposed so callers can run pip install against the same Python. */
    resolvePython(): string { return this._resolvePython(); }

    /**
     * call() invokes this when the backend was never started, so the Python
     * process is spawned lazily on first bridge use instead of at activation.
     * The host memoizes the start attempt and owns its status/error UX.
     */
    setOnDemandStart(start: () => Promise<void>): void {
        this._onDemandStart = start;
    }

    async restart(): Promise<void> {
        await this.stop();
        return this.start();
    }

    async stop(): Promise<void> {
        const proc = this._proc;
        this._stopping = true;
        this._ready = false;
        this._startPromise = undefined;
        this._readyResolve = undefined;
        this._readyReject?.(new Error('Python bridge stopped.'));
        this._readyReject = undefined;
        this._rejectAll(new Error('Python bridge stopped.'));
        if (!proc) {
            this._stopping = false;
            return;
        }
        await new Promise<void>((resolve) => {
            let settled = false;
            const done = () => {
                if (!settled) {
                    settled = true;
                    resolve();
                }
            };
            proc.once('exit', done);
            proc.once('error', done);
            if (proc.exitCode !== null) {
                done();
                return;
            }
            proc.kill();
        });
        if (this._proc === proc) {
            this._proc = undefined;
        }
        this._stopping = false;
    }

    async start(): Promise<void> {
        // Idempotent: if a start is already in flight or a live process exists,
        // return that start rather than spawning a second interpreter (which
        // would orphan the first - e.g. xlide.setup's installDependencies calling
        // start() directly while a lazy start already ran). restart() stops first
        // (clearing _proc), so it still re-spawns.
        if (this._proc && this._startPromise) {
            return this._startPromise;
        }
        // Reset state so start() can be called again after a failed attempt.
        this._ready = false;
        this._stderrLines = [];

        const pythonPath = this._resolvePython();
        const serverScript = path.join(
            this._context.extensionPath,
            'python',
            'server.py',
        );
        const serverDir = path.dirname(serverScript);

        this._out.appendLine(`Starting Python bridge: ${pythonPath} ${serverScript}`);
        const trace = startPerformanceTrace('pythonBridge.start', pythonPath);

        const startPromise = new Promise<void>((resolve, reject) => {
            this._readyResolve = resolve;
            this._readyReject = reject;

            const proc = cp.spawn(pythonPath, [serverScript], {
                stdio: ['pipe', 'pipe', 'pipe'],
                cwd: serverDir,
            });
            this._proc = proc;

            proc.on('error', (err) => {
                this._out.appendLine(`Process error: ${err.message}`);
                if (!this._stopping) {
                    this._readyReject?.(new Error(`Cannot start Python: ${err.message}`));
                }
                this._rejectAll(new Error(`Python process error: ${err.message}`));
            });

            proc.on('exit', (code) => {
                this._out.appendLine(`Python backend exited with code ${code}`);
                const wasStopping = this._stopping;
                const wasReady = this._ready;
                // Reset state so the next call() doesn't hang on a dead bridge.
                this._ready = false;
                this._startPromise = undefined;
                this._readyResolve = undefined;
                if (this._proc === proc) {
                    this._proc = undefined;
                }
                if (!wasReady && !wasStopping) {
                    const stderr = this._stderrLines.join('\n');
                    this._readyReject?.(
                        new Error(`Python backend exited (code ${code}).\n${stderr}`),
                    );
                }
                this._readyReject = undefined;
                this._rejectAll(new Error(`Python backend exited with code ${code}`));
                if (wasReady && !wasStopping) {
                    // Surface unexpected crash to the user.
                    void vscode.window.showWarningMessage(
                        `XLIDE: Python backend exited unexpectedly (code ${code}). Restart the backend or set xlide.pythonPath, then try again.`,
                        'Restart Backend',
                        'Copy Diagnostics',
                    ).then((choice) => {
                        if (choice === 'Restart Backend') {
                            void this.restart().catch((err: Error) => {
                                void vscode.window.showErrorMessage(
                                    `XLIDE: Failed to restart Python backend. ${err.message}`,
                                );
                            });
                        } else if (choice === 'Copy Diagnostics') {
                            void vscode.commands.executeCommand('xlide.copyDiagnostics');
                        }
                    });
                }
            });

            proc.stderr!.on('data', (chunk: Buffer) => {
                const text = chunk.toString().trim();
                if (text) {
                    this._out.appendLine(`[python] ${text}`);
                    this._stderrLines.push(text);
                    if (this._stderrLines.length > MAX_STDERR_LINES) {
                        this._stderrLines.shift();
                    }
                }
            });

            const rl = readline.createInterface({ input: proc.stdout! });
            rl.on('line', (line) => this._onLine(line));
        });
        this._startPromise = startPromise.then(
            () => {
                trace.end('ok', pythonPath);
            },
            (err) => {
                trace.end('failed', pythonPath);
                throw err;
            },
        );
        return this._startPromise;
    }

    private _resolvePython(): string {
        const configured = xlidePythonPathFromConfig(vscode.workspace.getConfiguration('xlide')).value;
        if (configured) {
            return configured;
        }

        // Auto-detect: prefer a .venv inside each workspace folder
        const folders = vscode.workspace.workspaceFolders;
        if (folders && folders.length > 0) {
            const venvBin = process.platform === 'win32'
                ? path.join(folders[0].uri.fsPath, '.venv', 'Scripts', 'python.exe')
                : path.join(folders[0].uri.fsPath, '.venv', 'bin', 'python3');
            try {
                const fs = require('fs') as typeof import('fs');
                if (fs.existsSync(venvBin)) {
                    this._out.appendLine(`Auto-detected Python: ${venvBin}`);
                    return venvBin;
                }
            } catch { /* ignore */ }
        }

        return process.platform === 'win32' ? 'python' : 'python3';
    }

    private _onLine(line: string): void {
        // Ready handshake from server.py - flush queued calls.
        if (line.trim() === '{"ready":true}') {
            this._out.appendLine('Python backend ready.');
            this._ready = true;
            this._readyResolve?.();
            return;
        }

        let msg: JsonRpcResponse;
        try {
            msg = JSON.parse(line) as JsonRpcResponse;
        } catch {
            this._out.appendLine(`Unparsable response from Python: ${line}`);
            return;
        }
        const pending = this._takePending(msg.id);
        if (!pending) { return; }
        if (msg.error) {
            pending.reject(new BridgeError(msg.error.message, msg.error.code));
        } else {
            pending.resolve(msg.result);
        }
    }

    call<T = unknown>(
        method: string,
        params?: unknown,
        token?: vscode.CancellationToken,
        options?: PythonBridgeCallOptions,
    ): Promise<T> {
        const trace = startPerformanceTrace('pythonBridge.call', method);
        const timeoutMs = options?.timeoutMs ?? DEFAULT_BRIDGE_CALL_TIMEOUT_MS;
        const doSend = (): Promise<T> => new Promise<T>((resolve, reject) => {
            if (token?.isCancellationRequested) {
                reject(new vscode.CancellationError());
                return;
            }
            const id = this._nextId++;
            let cancel: vscode.Disposable | undefined;
            this._pending.set(id, {
                resolve: resolve as (v: unknown) => void,
                reject,
                cancel,
            });
            if (token) {
                cancel = token.onCancellationRequested(() => {
                    // Stops waiting locally; the single-threaded server keeps
                    // running the request, but its late response is ignored.
                    this._takePending(id)?.reject(new vscode.CancellationError());
                });
                const pending = this._pending.get(id);
                if (pending) {
                    pending.cancel = cancel;
                }
            }
            const proc = this._proc;
            if (!proc?.stdin) {
                this._takePending(id);
                reject(new Error('XLIDE: Python bridge not started.'));
                return;
            }
            const req: JsonRpcRequest = { jsonrpc: '2.0', id, method, params };
            try {
                proc.stdin.write(JSON.stringify(req) + '\n');
            } catch (err) {
                this._pending.delete(id);
                cancel?.dispose();
                reject(err instanceof Error ? err : new Error(String(err)));
                return;
            }
            const sent = this._pending.get(id);
            if (sent && timeoutMs > 0) {
                sent.timer = setTimeout(() => {
                    this._takePending(id)?.reject(new Error(
                        `XLIDE: Python backend request '${method}' timed out after ${Math.round(timeoutMs / 1000)}s. ` +
                        'The backend may be busy or hung; restart it if XLIDE stays unresponsive.',
                    ));
                }, timeoutMs);
            }
        });

        let promise: Promise<T>;
        if (this._ready) {
            promise = doSend();
        } else if (this._startPromise) {
            promise = this._startPromise.then(() => doSend());
        } else if (this._onDemandStart) {
            promise = this._onDemandStart().then(() => doSend());
        } else {
            promise = Promise.reject(new Error('XLIDE: Python bridge not started.'));
        }
        return promise.then(
            (value) => {
                trace.end('ok', method);
                return value;
            },
            (err) => {
                trace.end(isCancellationLike(err) ? 'canceled' : 'failed', method);
                throw err;
            },
        );
    }

    /** Removes a pending request and releases its cancellation/timeout hooks. */
    private _takePending(id: number): PendingRequest | undefined {
        const pending = this._pending.get(id);
        if (!pending) { return undefined; }
        this._pending.delete(id);
        pending.cancel?.dispose();
        if (pending.timer !== undefined) {
            clearTimeout(pending.timer);
        }
        return pending;
    }

    private _rejectAll(err: Error): void {
        for (const [, pending] of this._pending) {
            pending.cancel?.dispose();
            if (pending.timer !== undefined) {
                clearTimeout(pending.timer);
            }
            pending.reject(err);
        }
        this._pending.clear();
    }

    dispose(): void {
        void this.stop();
    }
}
