import { readFileSync } from 'node:fs';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
    receive: undefined as undefined | ((message: unknown) => unknown),
    messages: [] as any[], docs: [] as any[], files: new Map<string, string>(),
    html: '',
    reply: true, search: undefined as any,
    applyEdit: vi.fn(), confirm: vi.fn(), open: vi.fn(),
}));

vi.mock('vscode', () => ({
    ViewColumn: { Beside: 2 },
    Range: class { constructor(public start: any, public end: any) {} },
    WorkspaceEdit: class {
        edits: any[] = [];
        replace(uri: any, range: any, text: string) { this.edits.push({ uri, range, text }); }
    },
    window: {
        createWebviewPanel: () => ({
            webview: { get html() { return state.html; }, set html(value) { state.html = value; }, postMessage: (message: any) => {
                state.messages.push(message);
                if (message.type === 'match' && state.reply) {
                    queueMicrotask(() => {
                        try { state.receive?.({ type: 'matched', id: message.id, result: state.search(message.text, message.query, message.limit) }); }
                        catch (error) { state.receive?.({ type: 'matched', id: message.id, error: String(error) }); }
                    });
                }
                return Promise.resolve(true);
            }, onDidReceiveMessage: (receive: any) => { state.receive = receive; return { dispose() {} }; } },
            onDidDispose: () => ({ dispose() {} }), dispose() {},
        }),
        showWarningMessage: (...args: any[]) => state.confirm(...args),
        showTextDocument: vi.fn(),
    },
    workspace: {
        get textDocuments() { return state.docs; },
        openTextDocument: (...args: any[]) => state.open(...args),
        applyEdit: (...args: any[]) => state.applyEdit(...args),
        fs: { readFile: (uri: any) => Promise.resolve(Buffer.from(state.files.get(uri.toString()) ?? '')) },
    },
}));
vi.mock('../../src/xlideFileSystem', () => ({ encodeModuleUri: (project: string, name: string) => uri(`${project}/${name}`),
    XLIDE_SCHEME: 'xlide-vba', decodeModuleUri: () => { throw new Error('Not a virtual test URI'); } }));
vi.mock('../../src/util/workspaceUris', () => ({ workspaceUriFor: (file: string) => uri(file) }));
vi.mock('../../src/extensionAssets', () => ({ readExtensionTextAsset: (file: string) => readFileSync(file, 'utf8').replace(/\r\n/g, '\n') }));
import { ModuleSearchPanel } from '../../src/moduleSearchPanel';

function uri(key: string) { return { toString: () => key }; }
function document(key: string, text: string) {
    const doc = { uri: uri(key), text, isClosed: false, getText: () => doc.text,
        positionAt: (offset: number) => ({ offset }) };
    state.docs.push(doc);
    return doc;
}
let panel: ModuleSearchPanel;
async function send(message: object) {
    return (panel as unknown as { receive(message: object): Promise<void> }).receive(message);
}
function setup(projects = ['/Book.xlsm'], listing = [{ name: 'A' }, { name: 'B' }] as Array<{ name: string; filePath?: string }>) {
    const call = vi.fn(async (method: string, _args: unknown) => method === 'listModules' ? listing : { ok: true });
    panel = new ModuleSearchPanel({
        bridge: { call }, explorer: { getChildren: async () => projects.map(filePath => ({ kind: 'project', filePath, label: filePath })) },
        fsProvider: { readFile: async (uri: any) => Buffer.from(state.files.get(uri.toString()) ?? '') },
    } as any, { kind: 'project', filePath: projects[0], label: projects[0] });
    return { call };
}
const search = (extra = {}) => send({ type: 'search', pattern: 'needle', scope: 'all', ...extra });

beforeEach(() => {
    state.messages = []; state.docs = []; state.files = new Map(); state.reply = true;
    state.search = require('../../assets/webview/moduleSearchWorker.js').searchModule;
    state.applyEdit.mockReset().mockResolvedValue(true);
    state.confirm.mockReset().mockResolvedValue('Apply replacements');
    state.open.mockReset().mockImplementation(async (uri: any) => state.docs.find(d => d.uri.toString() === uri.toString())
        ?? document(uri.toString(), state.files.get(uri.toString()) ?? ''));
});
afterEach(() => { panel?.dispose(); vi.useRealTimers(); });

describe('XLIDE project search controller', () => {
    it('embeds the exact worker source without interpreting dollar replacement tokens', () => {
        setup();
        const embedded = /const searchWorkerSource = (.+);/.exec(state.html)![1];
        expect(JSON.parse(embedded)).toBe(readFileSync('assets/webview/moduleSearchWorker.js', 'utf8').replace(/\r\n/g, '\n'));
    });

    it.each(['/Book.xlsm', '/Doc.docm', '/Slides.pptm', '/Database.accdb'])('searches unopened and dirty modules in %s', async project => {
        setup([project]);
        state.files.set(`${project}/A`, 'needle'); state.files.set(`${project}/B`, 'disk');
        document(`${project}/B`, 'unsaved needle');
        await search();
        expect(state.messages.find(m => m.type === 'done')).toMatchObject({ count: 2, incomplete: false });
        expect(state.open).not.toHaveBeenCalled();
        expect(state.applyEdit).not.toHaveBeenCalled();
    });

    it('uses VB6 native module paths and previews only selected matches', async () => {
        setup(['/Code.vbp'], [{ name: 'A', filePath: '/native/A.bas' }]);
        state.files.set('/native/A.bas', 'needle needle');
        await search();
        await send({ type: 'preview', ids: [1], replacement: 'updated' });
        expect(state.messages.find(m => m.type === 'preview')).toMatchObject({ count: 1, modules: 1 });
        expect(state.applyEdit).not.toHaveBeenCalled();
        await send({ type: 'apply' });
        const edits = state.applyEdit.mock.calls[0][0].edits;
        expect(edits).toMatchObject([{ uri: { toString: expect.any(Function) }, range: { start: { offset: 7 }, end: { offset: 13 } }, text: 'updated' }]);
    });

    it('refuses stale source before preview and stale documents after preview', async () => {
        setup(); state.files.set('/Book.xlsm/A', 'needle'); state.files.set('/Book.xlsm/B', 'needle');
        await search();
        const doc = document('/Book.xlsm/A', 'changed needle');
        await expect(send({ type: 'preview', all: true, replacement: 'x' })).rejects.toThrow('changed');
        doc.text = 'needle';
        await send({ type: 'preview', all: true, replacement: 'x' });
        state.docs.find(doc => doc.uri.toString() === '/Book.xlsm/B').text = 'new dirty content';
        await expect(send({ type: 'apply' })).rejects.toThrow('changed after preview');
        expect(state.applyEdit).not.toHaveBeenCalled();
    });

    it('does not apply an outdated replacement after confirmation was already opened', async () => {
        setup(); state.files.set('/Book.xlsm/A', 'needle');
        await search(); await send({ type: 'preview', all: true, replacement: 'x' });
        let confirm!: (answer: string) => void;
        state.confirm.mockImplementationOnce(() => new Promise(resolve => { confirm = resolve; }));
        const applying = send({ type: 'apply' });
        await send({ type: 'replacementChanged' });
        confirm('Apply replacements'); await applying;
        expect(state.applyEdit).not.toHaveBeenCalled();
    });

    it('restricts project and module scopes while all-tree scope searches both projects', async () => {
        setup(['/one.xlsm', '/two.docm']);
        for (const project of ['/one.xlsm', '/two.docm']) {
            for (const module of ['A', 'B']) { state.files.set(`${project}/${module}`, 'needle'); }
        }
        await search({ scope: 'project' });
        expect(state.messages.findLast(m => m.type === 'done').count).toBe(2);
        await search();
        expect(state.messages.findLast(m => m.type === 'done').count).toBe(4);
    });

    it('keeps a cancelled confirmation reusable and invalidates changed replacement previews', async () => {
        setup(); state.files.set('/Book.xlsm/A', 'needle');
        await search(); await send({ type: 'preview', all: true, replacement: 'x' });
        state.confirm.mockResolvedValueOnce(undefined);
        await send({ type: 'apply' });
        expect(state.messages.at(-1)).toMatchObject({ type: 'idle', hasPreview: true });
        await send({ type: 'replacementChanged' });
        await expect(send({ type: 'apply' })).rejects.toThrow('Preview');
        expect(state.applyEdit).not.toHaveBeenCalled();
    });

    it('never lets an unknown result ID select a filesystem path', async () => {
        setup(); state.files.set('/Book.xlsm/A', 'needle'); await search();
        await expect(send({ type: 'preview', ids: [999], replacement: 'x' })).rejects.toThrow('current search');
        await send({ type: 'open', id: 999, path: '/unrelated.txt' });
        expect(state.open).not.toHaveBeenCalled();
    });

    it('cancels pending regex work and rejects incomplete Replace All', async () => {
        setup(); state.files.set('/Book.xlsm/A', 'needle'); state.reply = false;
        const running = search();
        await vi.waitFor(() => expect(state.messages.some(m => m.type === 'match')).toBe(true));
        await send({ type: 'stop' }); await running;
        expect(state.messages.some(m => m.type === 'done')).toBe(false);
        expect(state.messages.some(m => m.type === 'resetWorker')).toBe(true);
        state.reply = true; state.files.set('/Book.xlsm/A', 'needle '.repeat(5001));
        await search();
        expect(state.messages.findLast(m => m.type === 'done')).toMatchObject({ incomplete: true });
        await expect(send({ type: 'preview', all: true, replacement: 'x' })).rejects.toThrow('incomplete');
    });

    it('terminates unresponsive matching without blocking the controller', async () => {
        vi.useFakeTimers(); setup(); state.reply = false; state.files.set('/Book.xlsm/A', 'needle');
        const running = search();
        await vi.advanceTimersByTimeAsync(2100); await running;
        expect(state.messages.find(m => m.type === 'done').errors.join(' ')).toContain('two seconds');
    });

    it('reports locked projects and leaves them untouched', async () => {
        const { call } = setup();
        call.mockRejectedValue(new Error('Project locked'));
        await search();
        expect(state.messages.find(m => m.type === 'done')).toMatchObject({ count: 0, incomplete: true });
        expect(state.applyEdit).not.toHaveBeenCalled();
    });
});
