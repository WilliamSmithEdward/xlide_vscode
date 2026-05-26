import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as path from 'path';
import * as readline from 'readline';

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
}

export class PythonBridge implements vscode.Disposable {
    private _proc: cp.ChildProcess | undefined;
    private _pending = new Map<number, PendingRequest>();
    private _nextId = 1;
    private _ready = false;
    private _queue: Array<() => void> = [];

    constructor(
        private readonly _context: vscode.ExtensionContext,
        private readonly _out: vscode.OutputChannel,
    ) {}

    async start(): Promise<void> {
        const pythonPath = this._resolvePython();
        const serverScript = path.join(
            this._context.extensionPath,
            'python',
            'server.py',
        );
        const serverDir = path.dirname(serverScript);

        this._out.appendLine(`Starting Python bridge: ${pythonPath} ${serverScript}`);

        this._proc = cp.spawn(pythonPath, [serverScript], {
            stdio: ['pipe', 'pipe', 'pipe'],
            cwd: serverDir,   // xlide package is importable from here
        });

        this._proc.on('error', (err) => {
            this._out.appendLine(`Process error: ${err.message}`);
            this._rejectAll(new Error(`Python process error: ${err.message}`));
            throw err;
        });

        this._proc.on('exit', (code) => {
            this._out.appendLine(`Python backend exited with code ${code}`);
            this._rejectAll(new Error(`Python backend exited with code ${code}`));
        });

        this._proc.stderr!.on('data', (chunk: Buffer) => {
            const text = chunk.toString().trim();
            if (text) {
                this._out.appendLine(`[python] ${text}`);
            }
        });

        const rl = readline.createInterface({ input: this._proc.stdout! });
        rl.on('line', (line) => this._onLine(line));
        // _ready is set to true only when the server emits {"ready":true}
        // (after all Python imports finish).  Until then, calls are queued.
    }

    private _resolvePython(): string {
        const configured = vscode.workspace
            .getConfiguration('xlide')
            .get<string>('pythonPath');
        if (configured && configured.trim()) {
            return configured.trim();
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
        // Ready handshake from server.py — flush queued calls.
        if (line.trim() === '{"ready":true}') {
            this._out.appendLine('Python backend ready.');
            this._ready = true;
            for (const fn of this._queue) { fn(); }
            this._queue = [];
            return;
        }

        let msg: JsonRpcResponse;
        try {
            msg = JSON.parse(line) as JsonRpcResponse;
        } catch {
            this._out.appendLine(`Unparsable response from Python: ${line}`);
            return;
        }
        const pending = this._pending.get(msg.id);
        if (!pending) { return; }
        this._pending.delete(msg.id);
        if (msg.error) {
            pending.reject(new Error(msg.error.message));
        } else {
            pending.resolve(msg.result);
        }
    }

    call<T = unknown>(method: string, params: unknown): Promise<T> {
        return new Promise<T>((resolve, reject) => {
            const doSend = () => {
                const id = this._nextId++;
                this._pending.set(id, {
                    resolve: resolve as (v: unknown) => void,
                    reject,
                });
                const req: JsonRpcRequest = {
                    jsonrpc: '2.0',
                    id,
                    method,
                    params,
                };
                this._proc!.stdin!.write(JSON.stringify(req) + '\n');
            };

            if (this._ready) {
                doSend();
            } else {
                this._queue.push(doSend);
            }
        });
    }

    private _rejectAll(err: Error): void {
        for (const [, pending] of this._pending) {
            pending.reject(err);
        }
        this._pending.clear();
    }

    dispose(): void {
        this._proc?.kill();
    }
}
