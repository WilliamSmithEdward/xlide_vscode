// The loopback API a window serves, which the XLIDE MCP server (xlide_mcp)
// reports its edits to. The protocol is docs/xlide-vscode-bridge.md in
// xlide_mcp; what each report does to the window is mcpEditMirror.ts.
//
// The window listens on 127.0.0.1 at a port the system picks. Every path
// starts with a random token, compared in constant time, and the window says
// where it listens in a record only this user can read: xlide-api-{pid}.json,
// in %LOCALAPPDATA%\xlide_vscode on Windows and $XDG_STATE_HOME/xlide_vscode
// elsewhere. The record is rewritten when the workspace folders change and
// deleted when the window closes, and a window starting up deletes the records
// of windows that are gone. The server believes no record until `hello`
// answers.
//
// Node only: a browser has no socket to listen on, and the web build never
// imports this.

import { createHash, randomBytes, timingSafeEqual } from 'crypto';
import * as fs from 'fs';
import * as http from 'http';
import type { AddressInfo } from 'net';
import * as os from 'os';
import * as path from 'path';
import { errorMessage } from './util/errors';

export const XLIDE_API_PRODUCT = 'xlide_vscode';
export const XLIDE_API_PROTOCOL = 1;

/** What a window writes about itself for the server to find. */
export interface XlideApiRecord {
    pid: number;
    port: number;
    token: string;
    product: typeof XLIDE_API_PRODUCT;
    version: string;
    protocol: typeof XLIDE_API_PROTOCOL;
    workspaceFolders: string[];
}

/** A module one tool call wrote, created or deleted. */
export interface AgentEditReport {
    file: string;
    module: string;
    /** The module's code before the call; empty when it did not exist. */
    before: string;
    beforeExisted: boolean;
    /** The module's code after the call; empty when the call deleted it. */
    after: string;
    afterExists: boolean;
    /** What the server says the edit was. Passed along, never branched on. */
    kind?: string;
}

export interface ModuleRenamedReport {
    file: string;
    from: string;
    to: string;
}

/** A change to the file that is not a module's code: cells, shapes, references. */
export interface FileChangedReport {
    file: string;
    what?: string;
}

/** What every report answers. The server passes it on to the agent. */
export interface ReportAnswer {
    /** This window has the file in its tree, or a module of it open. */
    shown: boolean;
    /** Whether the module waits in the tree for Keep or Revert. */
    review: 'pending' | 'none';
}

export interface XlideApiHandlers {
    agentEdit(report: AgentEditReport): Promise<ReportAnswer>;
    moduleRenamed(report: ModuleRenamedReport): Promise<ReportAnswer>;
    fileChanged(report: FileChangedReport): Promise<ReportAnswer>;
}

export interface XlideApiServerOptions {
    handlers: XlideApiHandlers;
    /** The extension's version, for the record. */
    version: string;
    /** The window's workspace folders, as file system paths. */
    workspaceFolders(): string[];
    log(line: string): void;
    /** Where the record goes. The platform's state directory unless a test says otherwise. */
    stateDir?: string;
}

/**
 * A report carries a module's code twice, and a module holds at most 65,535
 * lines, so this is far past any real report and still bounds what a caller
 * can make the window hold.
 */
const MAX_REPORT_BYTES = 32 * 1024 * 1024;

/** The server gives each window a second; a request still open long after that is stuck. */
const REQUEST_TIMEOUT_MS = 15_000;
const HEADERS_TIMEOUT_MS = 10_000;

const RECORD_NAME = /^\.?xlide-api-(\d+)\.json(?:\.tmp)?$/;

/**
 * The folder the records go in. Windows keeps per-user state under
 * %LOCALAPPDATA%; elsewhere it is $XDG_STATE_HOME, which the XDG base
 * directory specification says to ignore when it is unset, empty or relative,
 * falling back to ~/.local/state.
 */
export function xlideApiStateDir(
    platform: NodeJS.Platform = process.platform,
    env: NodeJS.ProcessEnv = process.env,
    home: string = os.homedir(),
): string {
    if (platform === 'win32') {
        const local = env.LOCALAPPDATA || path.win32.join(home, 'AppData', 'Local');
        return path.win32.join(local, 'xlide_vscode');
    }
    const state = env.XDG_STATE_HOME && path.posix.isAbsolute(env.XDG_STATE_HOME)
        ? env.XDG_STATE_HOME
        : path.posix.join(home, '.local', 'state');
    return path.posix.join(state, 'xlide_vscode');
}

export function xlideApiRecordName(pid: number): string {
    return `xlide-api-${pid}.json`;
}

/**
 * Deletes the records, and any temp file a write left, of windows whose
 * process is gone. A window that crashed never deleted its own. The server
 * would find nothing listening there and move on, but the folder would
 * collect them. Returns how many went.
 */
export function sweepStaleXlideApiRecords(
    stateDir: string,
    isRunning: (pid: number) => boolean = processIsRunning,
    ownPid: number = process.pid,
): number {
    let names: string[];
    try {
        names = fs.readdirSync(stateDir);
    } catch {
        return 0;
    }
    let removed = 0;
    for (const name of names) {
        const match = RECORD_NAME.exec(name);
        if (!match) {
            continue;
        }
        const pid = Number(match[1]);
        if (pid === ownPid || isRunning(pid)) {
            continue;
        }
        try {
            fs.rmSync(path.join(stateDir, name), { force: true });
            removed += 1;
        } catch {
            // Another window starting at the same moment swept it first.
        }
    }
    return removed;
}

/** Signal 0 sends nothing: it only asks whether the process exists. */
function processIsRunning(pid: number): boolean {
    try {
        process.kill(pid, 0);
        return true;
    } catch (err) {
        // EPERM: it exists, and belongs to someone else.
        return (err as NodeJS.ErrnoException).code === 'EPERM';
    }
}

/** A request the window refuses, with the status that says why. */
class RefusedRequest extends Error {
    constructor(readonly status: number, message: string) {
        super(message);
    }
}

type Route = 'hello' | 'agent-edit' | 'module-renamed' | 'file-changed';

const ROUTE_METHODS: Record<Route, 'GET' | 'POST'> = {
    'hello': 'GET',
    'agent-edit': 'POST',
    'module-renamed': 'POST',
    'file-changed': 'POST',
};

function isRoute(name: string): name is Route {
    return Object.prototype.hasOwnProperty.call(ROUTE_METHODS, name);
}

export class XlideApiServer {
    private _recordWritten = false;
    private _disposed = false;

    private constructor(
        private readonly _server: http.Server,
        private readonly _options: XlideApiServerOptions,
        private readonly _token: string,
        readonly port: number,
        readonly recordPath: string,
    ) {}

    /**
     * Listens, sweeps the records of windows that are gone, and writes this
     * window's. Rejects when the port cannot be opened or the record cannot be
     * written: a server nobody can find is no use.
     */
    static async start(options: XlideApiServerOptions): Promise<XlideApiServer> {
        const token = randomBytes(32).toString('hex');
        const stateDir = options.stateDir ?? xlideApiStateDir();
        let api: XlideApiServer | undefined;
        const server = http.createServer((request, response) => {
            if (api) {
                void api._answer(request, response);
            } else {
                response.destroy();
            }
        });
        server.requestTimeout = REQUEST_TIMEOUT_MS;
        server.headersTimeout = HEADERS_TIMEOUT_MS;
        await new Promise<void>((resolve, reject) => {
            server.once('error', reject);
            server.listen(0, '127.0.0.1', () => {
                server.off('error', reject);
                resolve();
            });
        });
        // A failure after listening (a socket error, say) must not take the
        // extension host down with it.
        server.on('error', (err) => options.log(`XLIDE API server error: ${errorMessage(err)}`));
        const port = (server.address() as AddressInfo).port;
        api = new XlideApiServer(server, options, token, port, path.join(stateDir, xlideApiRecordName(process.pid)));
        try {
            sweepStaleXlideApiRecords(stateDir);
            api.writeRecord();
        } catch (err) {
            api.dispose();
            throw err;
        }
        return api;
    }

    /** Writes the record again, for a change of workspace folders. */
    writeRecord(): void {
        if (this._disposed) {
            return;
        }
        const record: XlideApiRecord = {
            pid: process.pid,
            port: this.port,
            token: this._token,
            product: XLIDE_API_PRODUCT,
            version: this._options.version,
            protocol: XLIDE_API_PROTOCOL,
            workspaceFolders: this._options.workspaceFolders(),
        };
        writePrivateFile(this.recordPath, `${JSON.stringify(record, null, 2)}\n`);
        this._recordWritten = true;
    }

    /**
     * Deletes the record first, synchronously, so a window that is closing
     * leaves nothing to find even if it never gets to run the rest.
     */
    dispose(): void {
        if (this._disposed) {
            return;
        }
        this._disposed = true;
        if (this._recordWritten) {
            try {
                fs.rmSync(this.recordPath, { force: true });
            } catch {
                // Left for the next window's sweep.
            }
        }
        this._server.close();
        this._server.closeAllConnections();
    }

    private async _answer(request: http.IncomingMessage, response: http.ServerResponse): Promise<void> {
        let route: Route | undefined;
        try {
            route = this._route(request);
            if (route === 'hello') {
                send(response, 200, { product: XLIDE_API_PRODUCT, protocol: XLIDE_API_PROTOCOL });
                return;
            }
            const body = parseJson(await readBody(request));
            send(response, 200, await this._dispatch(route, body));
        } catch (err) {
            if (err instanceof RefusedRequest) {
                send(response, err.status, { error: err.message });
                return;
            }
            this._options.log(`XLIDE API: ${route ?? 'a request'} failed: ${errorMessage(err)}`);
            send(response, 500, { error: 'The window could not take the report.' });
        }
    }

    /** The endpoint a request names, once its token is this window's. */
    private _route(request: http.IncomingMessage): Route {
        const { pathname } = new URL(request.url ?? '/', 'http://127.0.0.1');
        const [token = '', name = '', ...rest] = pathname.split('/').slice(1);
        // A wrong token and a missing endpoint answer alike, so a caller
        // without the token learns nothing about what is here.
        if (!sameToken(token, this._token) || !isRoute(name) || rest.some((segment) => segment !== '')) {
            throw new RefusedRequest(404, 'Not found.');
        }
        if (request.method !== ROUTE_METHODS[name]) {
            throw new RefusedRequest(405, `${name} takes ${ROUTE_METHODS[name]}.`);
        }
        return name;
    }

    private _dispatch(route: Exclude<Route, 'hello'>, body: Record<string, unknown>): Promise<ReportAnswer> {
        const handlers = this._options.handlers;
        switch (route) {
            case 'agent-edit':
                return handlers.agentEdit(agentEditReport(body));
            case 'module-renamed':
                return handlers.moduleRenamed({
                    file: filePathField(body),
                    from: textField(body, 'from', true),
                    to: textField(body, 'to', true),
                });
            case 'file-changed':
                return handlers.fileChanged({
                    file: filePathField(body),
                    what: informationalField(body, 'what'),
                });
        }
    }
}

/** Hashing first makes the two the same length, which timingSafeEqual needs. */
function sameToken(given: string, expected: string): boolean {
    const digest = (value: string): Buffer => createHash('sha256').update(value).digest();
    return timingSafeEqual(digest(given), digest(expected));
}

function agentEditReport(body: Record<string, unknown>): AgentEditReport {
    const beforeExisted = booleanField(body, 'beforeExisted');
    return {
        file: filePathField(body),
        module: textField(body, 'module', true),
        // Revert writes the before back, so a module that existed must come
        // with one: a missing before would revert it to nothing.
        before: beforeExisted ? textField(body, 'before', false) : codeField(body, 'before'),
        beforeExisted,
        // The window reads the after for itself; this is only a fallback.
        after: codeField(body, 'after'),
        afterExists: booleanField(body, 'afterExists'),
        kind: informationalField(body, 'kind'),
    };
}

function filePathField(body: Record<string, unknown>): string {
    const file = textField(body, 'file', true);
    if (!path.isAbsolute(file)) {
        throw new RefusedRequest(400, 'file must be an absolute path.');
    }
    return file;
}

function textField(body: Record<string, unknown>, name: string, required: boolean): string {
    const value = body[name];
    if (typeof value !== 'string' || (required && value === '')) {
        throw new RefusedRequest(400, `${name} must be a${required ? ' non-empty' : ''} string.`);
    }
    return value;
}

/** A field the window only passes along: never a reason to refuse a report. */
function informationalField(body: Record<string, unknown>, name: string): string | undefined {
    const value = body[name];
    return typeof value === 'string' ? value : undefined;
}

/** A module's code. Missing or null reads as none, for the side of a create or a delete that has no module. */
function codeField(body: Record<string, unknown>, name: string): string {
    return body[name] === undefined || body[name] === null ? '' : textField(body, name, false);
}

function booleanField(body: Record<string, unknown>, name: string): boolean {
    const value = body[name];
    if (typeof value !== 'boolean') {
        throw new RefusedRequest(400, `${name} must be true or false.`);
    }
    return value;
}

function parseJson(text: string): Record<string, unknown> {
    let value: unknown;
    try {
        value = JSON.parse(text);
    } catch {
        throw new RefusedRequest(400, 'The report is not JSON.');
    }
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
        throw new RefusedRequest(400, 'The report must be a JSON object.');
    }
    return value as Record<string, unknown>;
}

function readBody(request: http.IncomingMessage): Promise<string> {
    const declared = Number(request.headers['content-length']);
    if (Number.isFinite(declared) && declared > MAX_REPORT_BYTES) {
        return Promise.reject(new RefusedRequest(413, 'The report is too large.'));
    }
    return new Promise((resolve, reject) => {
        const chunks: Buffer[] = [];
        let size = 0;
        request.on('data', (chunk: Buffer) => {
            size += chunk.length;
            if (size > MAX_REPORT_BYTES) {
                // Stop taking it in; the refusal closes the connection.
                request.removeAllListeners('data');
                request.pause();
                reject(new RefusedRequest(413, 'The report is too large.'));
                return;
            }
            chunks.push(chunk);
        });
        request.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
        request.on('error', reject);
    });
}

function send(response: http.ServerResponse, status: number, body: unknown): void {
    if (response.headersSent) {
        return;
    }
    const text = JSON.stringify(body);
    response.writeHead(status, {
        'Content-Type': 'application/json; charset=utf-8',
        'Content-Length': Buffer.byteLength(text),
        'Cache-Control': 'no-store',
        // A refused body may still be arriving; do not wait to read it.
        ...(status === 413 ? { Connection: 'close' } : {}),
    });
    response.end(text);
}

/**
 * Writes a file only its owner can read, through a temp file and a rename so
 * a reader never sees half of it. On Windows a rename over a file another
 * process has open fails; the file is then written in place.
 */
function writePrivateFile(filePath: string, text: string): void {
    const dir = path.dirname(filePath);
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
    const temp = path.join(dir, `.${path.basename(filePath)}.tmp`);
    fs.writeFileSync(temp, text, { mode: 0o600 });
    try {
        fs.renameSync(temp, filePath);
    } catch {
        fs.rmSync(temp, { force: true });
        fs.writeFileSync(filePath, text, { mode: 0o600 });
    }
}
