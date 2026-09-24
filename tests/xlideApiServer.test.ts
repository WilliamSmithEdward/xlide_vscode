import { spawn } from 'child_process';
import * as fs from 'fs';
import * as http from 'http';
import * as os from 'os';
import * as nodePath from 'path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
    sweepStaleXlideApiRecords,
    xlideApiRecordName,
    xlideApiStateDir,
    XlideApiServer,
    type AgentEditReport,
    type XlideApiHandlers,
    type XlideApiRecord,
} from '../src/xlideApiServer';

const ANSWER = { shown: true, review: 'pending' as const };

function fakeHandlers(): XlideApiHandlers & { [K in keyof XlideApiHandlers]: ReturnType<typeof vi.fn> } {
    return {
        agentEdit: vi.fn(async () => ANSWER),
        moduleRenamed: vi.fn(async () => ANSWER),
        fileChanged: vi.fn(async () => ANSWER),
    };
}

/** A report the way the server sends one for a module it edited. */
function editReport(overrides: Record<string, unknown> = {}): Record<string, unknown> {
    return {
        file: nodePath.resolve('/work/Book.xlsm'),
        module: 'Module1',
        before: 'Sub A()\r\nEnd Sub\r\n',
        beforeExisted: true,
        after: 'Sub B()\r\nEnd Sub\r\n',
        afterExists: true,
        kind: 'write',
        ...overrides,
    };
}

describe('the loopback API a window serves', () => {
    let stateDir: string;
    let handlers: ReturnType<typeof fakeHandlers>;
    let log: ReturnType<typeof vi.fn>;
    let folders: string[];
    let server: XlideApiServer;

    beforeEach(async () => {
        stateDir = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'xlide-api-'));
        handlers = fakeHandlers();
        log = vi.fn();
        folders = [nodePath.resolve('/work')];
        server = await XlideApiServer.start({
            handlers,
            version: '10.7.2',
            workspaceFolders: () => folders,
            log,
            stateDir,
        });
    });

    afterEach(() => {
        server.dispose();
        fs.rmSync(stateDir, { recursive: true, force: true });
    });

    function record(): XlideApiRecord {
        return JSON.parse(fs.readFileSync(server.recordPath, 'utf8')) as XlideApiRecord;
    }

    function url(route: string, token = record().token): string {
        return `http://127.0.0.1:${server.port}/${token}/${route}`;
    }

    async function post(route: string, body: unknown, token?: string): Promise<{ status: number; json: unknown }> {
        const response = await fetch(url(route, token), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: typeof body === 'string' ? body : JSON.stringify(body),
        });
        return { status: response.status, json: await response.json() };
    }

    it('writes a record naming this process, the port and the token', () => {
        expect(server.recordPath).toBe(nodePath.join(stateDir, xlideApiRecordName(process.pid)));
        const written = record();
        expect(written).toEqual({
            pid: process.pid,
            port: server.port,
            token: written.token,
            product: 'xlide_vscode',
            version: '10.7.2',
            protocol: 1,
            workspaceFolders: folders,
        });
        expect(written.token).toMatch(/^[0-9a-f]{64}$/);
        expect(server.port).toBeGreaterThan(0);
    });

    it.skipIf(process.platform === 'win32')('lets only its owner read the record', () => {
        expect(fs.statSync(server.recordPath).mode & 0o777).toBe(0o600);
    });

    it('answers hello with the product and protocol, which is what the server checks first', async () => {
        const response = await fetch(url('hello'));

        expect(response.status).toBe(200);
        expect(await response.json()).toEqual({ product: 'xlide_vscode', protocol: 1 });
    });

    it('answers nothing to a caller without the token', async () => {
        const token = record().token;
        const wrong = `${token.slice(0, -1)}${token.endsWith('0') ? '1' : '0'}`;
        for (const path of [`/${wrong}/hello`, '/hello', '/', `/${token}`, `/${token}/nothing`, `/${token}/hello/more`]) {
            const response = await fetch(`http://127.0.0.1:${server.port}${path}`);
            expect(response.status, path).toBe(404);
        }
        const refused = await post('agent-edit', editReport(), wrong);
        expect(refused.status).toBe(404);
        expect(handlers.agentEdit).not.toHaveBeenCalled();
    });

    it('refuses a request made with the wrong method', async () => {
        expect((await fetch(url('agent-edit'))).status).toBe(405);
        expect((await post('hello', {})).status).toBe(405);
    });

    it('hands each report to its handler and answers with what the handler says', async () => {
        const edit = await post('agent-edit', editReport());
        const renamed = await post('module-renamed', { file: nodePath.resolve('/work/Book.xlsm'), from: 'Old', to: 'New' });
        const changed = await post('file-changed', { file: nodePath.resolve('/work/Book.xlsm'), what: 'cells' });

        for (const answer of [edit, renamed, changed]) {
            expect(answer).toEqual({ status: 200, json: ANSWER });
        }
        expect(handlers.agentEdit).toHaveBeenCalledWith({
            file: nodePath.resolve('/work/Book.xlsm'),
            module: 'Module1',
            before: 'Sub A()\r\nEnd Sub\r\n',
            beforeExisted: true,
            after: 'Sub B()\r\nEnd Sub\r\n',
            afterExists: true,
            kind: 'write',
        } satisfies AgentEditReport);
        expect(handlers.moduleRenamed).toHaveBeenCalledWith({ file: nodePath.resolve('/work/Book.xlsm'), from: 'Old', to: 'New' });
        expect(handlers.fileChanged).toHaveBeenCalledWith({ file: nodePath.resolve('/work/Book.xlsm'), what: 'cells' });
    });

    it('refuses a report it cannot act on safely', async () => {
        const refusals: Array<[string, unknown]> = [
            ['not JSON', '{'],
            ['not an object', '[1]'],
            ['no module', editReport({ module: undefined })],
            ['an empty module', editReport({ module: '' })],
            ['a relative path', editReport({ file: 'Book.xlsm' })],
            ['beforeExisted not a boolean', editReport({ beforeExisted: 'yes' })],
            ['afterExists missing', editReport({ afterExists: undefined })],
            // Revert writes the before back: none would revert the module to nothing.
            ['a module that existed, with no before', editReport({ before: undefined })],
            ['a module that existed, with a null before', editReport({ before: null })],
        ];
        for (const [what, body] of refusals) {
            const answer = await post('agent-edit', body);
            expect(answer.status, what).toBe(400);
            expect(answer.json, what).toHaveProperty('error');
        }
        expect((await post('module-renamed', { file: nodePath.resolve('/work/Book.xlsm'), from: 'Old' })).status).toBe(400);
        expect(handlers.agentEdit).not.toHaveBeenCalled();
        expect(handlers.moduleRenamed).not.toHaveBeenCalled();
    });

    it('reads the missing side of a create or a delete as no code', async () => {
        await post('agent-edit', editReport({ before: null, beforeExisted: false }));
        await post('agent-edit', editReport({ after: undefined, afterExists: false }));

        expect(handlers.agentEdit.mock.calls[0][0]).toMatchObject({ before: '', beforeExisted: false });
        expect(handlers.agentEdit.mock.calls[1][0]).toMatchObject({ after: '', afterExists: false });
    });

    it('never refuses a report for a field it only passes along', async () => {
        const answer = await post('agent-edit', editReport({ kind: 5 }));
        await post('file-changed', { file: nodePath.resolve('/work/Book.xlsm') });

        expect(answer.status).toBe(200);
        expect(handlers.agentEdit.mock.calls[0][0].kind).toBeUndefined();
        expect(handlers.fileChanged).toHaveBeenCalledWith({ file: nodePath.resolve('/work/Book.xlsm'), what: undefined });
    });

    it('refuses a report larger than any real one, without reading it', async () => {
        const status = await new Promise<number>((resolve, reject) => {
            const request = http.request(url('agent-edit'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Content-Length': 64 * 1024 * 1024 },
            }, (response) => {
                response.resume();
                resolve(response.statusCode ?? 0);
            });
            request.on('error', reject);
            request.write('{');
        });

        expect(status).toBe(413);
        expect(handlers.agentEdit).not.toHaveBeenCalled();
    });

    it('answers a failed handler with 500 and says why in the log', async () => {
        handlers.agentEdit.mockRejectedValueOnce(new Error('engine busy'));

        const answer = await post('agent-edit', editReport());

        expect(answer.status).toBe(500);
        expect(log).toHaveBeenCalledWith(expect.stringContaining('engine busy'));
    });

    it('writes the record again with the folders the window has now', () => {
        folders = [nodePath.resolve('/work'), nodePath.resolve('/other')];
        server.writeRecord();

        expect(record().workspaceFolders).toEqual(folders);
    });

    it('deletes its record and stops listening when the window closes', async () => {
        const port = server.port;
        server.dispose();

        expect(fs.existsSync(server.recordPath)).toBe(false);
        await expect(fetch(`http://127.0.0.1:${port}/x/hello`)).rejects.toThrow();
    });
});

describe('the records of windows that are gone', () => {
    let stateDir: string;

    beforeEach(() => {
        stateDir = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'xlide-api-sweep-'));
    });

    afterEach(() => {
        fs.rmSync(stateDir, { recursive: true, force: true });
    });

    function touch(name: string): void {
        fs.writeFileSync(nodePath.join(stateDir, name), '{}');
    }

    it('are deleted, and the records of running windows kept', () => {
        touch(xlideApiRecordName(111));
        touch(`.${xlideApiRecordName(111)}.tmp`);
        touch(xlideApiRecordName(222));
        touch(xlideApiRecordName(333));
        touch('notes.json');

        const removed = sweepStaleXlideApiRecords(stateDir, (pid) => pid === 222, 333);

        expect(removed).toBe(2);
        expect(fs.readdirSync(stateDir).sort()).toEqual(['notes.json', xlideApiRecordName(222), xlideApiRecordName(333)]);
    });

    it('tells a process that has exited from one still running', async () => {
        // A child of this test, left to exit on its own: the probe sends it nothing.
        const child = spawn(process.execPath, ['-e', 'setTimeout(() => {}, 1000)'], { stdio: 'ignore' });
        const pid = child.pid!;
        touch(xlideApiRecordName(pid));

        expect(sweepStaleXlideApiRecords(stateDir)).toBe(0);
        await new Promise((resolve) => child.once('exit', resolve));
        expect(sweepStaleXlideApiRecords(stateDir)).toBe(1);
        expect(fs.readdirSync(stateDir)).toEqual([]);
    });

    it('are looked for where nothing is written yet without failing', () => {
        expect(sweepStaleXlideApiRecords(nodePath.join(stateDir, 'missing'))).toBe(0);
    });
});

describe('where the records go', () => {
    it('is under %LOCALAPPDATA% on Windows', () => {
        expect(xlideApiStateDir('win32', { LOCALAPPDATA: 'C:\\Users\\me\\AppData\\Local' }, 'C:\\Users\\me'))
            .toBe('C:\\Users\\me\\AppData\\Local\\xlide_vscode');
        expect(xlideApiStateDir('win32', {}, 'C:\\Users\\me')).toBe('C:\\Users\\me\\AppData\\Local\\xlide_vscode');
    });

    it('is under $XDG_STATE_HOME elsewhere, and ~/.local/state when that is unset, empty or relative', () => {
        expect(xlideApiStateDir('linux', { XDG_STATE_HOME: '/state' }, '/home/me')).toBe('/state/xlide_vscode');
        expect(xlideApiStateDir('darwin', {}, '/Users/me')).toBe('/Users/me/.local/state/xlide_vscode');
        expect(xlideApiStateDir('linux', { XDG_STATE_HOME: '' }, '/home/me')).toBe('/home/me/.local/state/xlide_vscode');
        expect(xlideApiStateDir('linux', { XDG_STATE_HOME: 'state' }, '/home/me')).toBe('/home/me/.local/state/xlide_vscode');
    });
});
