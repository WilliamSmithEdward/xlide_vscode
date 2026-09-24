import * as fs from 'fs';
import * as os from 'os';
import * as nodePath from 'path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const hoisted = vi.hoisted(() => ({
    settings: {} as Record<string, unknown>,
    configListeners: [] as Array<(event: { affectsConfiguration(section: string): boolean }) => void>,
    folderListeners: [] as Array<() => void>,
    textDocuments: [] as unknown[],
}));

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    workspace: {
        textDocuments: hoisted.textDocuments,
        workspaceFolders: [],
        getConfiguration: () => ({
            get: (key: string, fallback?: unknown) => key in hoisted.settings ? hoisted.settings[key] : fallback,
            inspect: () => ({}),
        }),
        onDidChangeConfiguration: (listener: (typeof hoisted.configListeners)[number]) => {
            hoisted.configListeners.push(listener);
            return { dispose: () => hoisted.configListeners.splice(hoisted.configListeners.indexOf(listener), 1) };
        },
        onDidChangeWorkspaceFolders: (listener: () => void) => {
            hoisted.folderListeners.push(listener);
            return { dispose: () => hoisted.folderListeners.splice(hoisted.folderListeners.indexOf(listener), 1) };
        },
    },
}));

import * as vscode from 'vscode';
import { mcpEditMirrorHandlers, mirrorMcpEdits } from '../src/mcpEditMirror';
import { onDidChangeProjectFile } from '../src/projectFileChanges';
import { hasPendingAgentReview, revertAgentChange, type AgentWriteReviewDeps } from '../src/xlideAgentDiff';
import { xlideApiRecordName } from '../src/xlideApiServer';
import { encodeModuleUri } from '../src/xlideFileSystem';

interface FakeBridge {
    call: ReturnType<typeof vi.fn>;
}

describe('mirroring the MCP server\'s edits', () => {
    let dir: string;
    let file: string;
    let changed: string[];
    let subscription: { dispose(): void };
    let bridge: FakeBridge;
    let listed: boolean;
    let log: ReturnType<typeof vi.fn>;

    beforeEach(() => {
        // A file of its own per test: reviews are kept per file, for the session.
        dir = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'xlide-mirror-'));
        file = nodePath.join(dir, 'Book.xlsm');
        fs.writeFileSync(file, 'the file as the server left it');
        changed = [];
        subscription = onDidChangeProjectFile((projectPath) => changed.push(projectPath));
        bridge = { call: vi.fn(async () => ({ source: 'Sub B()\r\nEnd Sub\r\n' })) };
        listed = true;
        log = vi.fn();
        hoisted.settings = {};
        hoisted.textDocuments.length = 0;
        vi.mocked(vscode.commands.executeCommand).mockClear();
    });

    afterEach(() => {
        subscription.dispose();
        fs.rmSync(dir, { recursive: true, force: true });
    });

    function handlers() {
        return mcpEditMirrorHandlers({ bridge, treeListsProject: async () => listed, log });
    }

    function edit(overrides: Partial<Parameters<ReturnType<typeof handlers>['agentEdit']>[0]> = {}) {
        return {
            file,
            module: 'Module1',
            before: 'Sub A()\r\nEnd Sub\r\n',
            beforeExisted: true,
            after: 'Sub B()\r\nEnd Sub\r\n',
            afterExists: true,
            ...overrides,
        };
    }

    function diffsOpened(): unknown[][] {
        return vi.mocked(vscode.commands.executeCommand).mock.calls.filter(([command]) => command === 'vscode.diff');
    }

    /** Reverts through the tree's action, the module reading as `current`. */
    async function revert(moduleName: string, current: string): Promise<string[]> {
        const written: string[] = [];
        const deps: AgentWriteReviewDeps = {
            readModuleSource: async () => current,
            writeModuleSource: async (_file, _module, source) => { written.push(source); },
            deleteModule: async () => { written.push('<deleted>'); },
        };
        await revertAgentChange(deps, file, moduleName);
        return written;
    }

    it('takes the new state and presents a review, with XLIDE\'s own read as the after', async () => {
        // The server's copy differs from what XLIDE reads: Revert must expect XLIDE's.
        const answer = await handlers().agentEdit(edit({ after: 'Sub B()\r\n    \r\nEnd Sub\r\n' }));

        expect(answer).toEqual({ shown: true, review: 'pending' });
        expect(changed).toEqual([file]);
        expect(bridge.call).toHaveBeenCalledWith('readModule', { path: file, module: 'Module1' });
        expect(diffsOpened()).toHaveLength(1);
        // Revert compares the module with XLIDE's read, then writes the before back.
        expect(await revert('Module1', 'Sub B()\r\nEnd Sub\r\n')).toEqual(['Sub A()\r\nEnd Sub\r\n']);
    });

    it('takes the server\'s before in the engine\'s form, so a write that changed nothing leaves nothing to review', async () => {
        bridge.call.mockResolvedValue({ source: 'Sub A()\r\nEnd Sub\r\n' });

        const answer = await handlers().agentEdit(edit({ before: '\r\n\r\nSub A()\r\nEnd Sub\r\n' }));

        expect(answer).toEqual({ shown: true, review: 'none' });
        expect(diffsOpened()).toHaveLength(0);
    });

    it('presents a module the server created as one Revert deletes', async () => {
        const answer = await handlers().agentEdit(edit({ before: '', beforeExisted: false }));

        expect(answer.review).toBe('pending');
        expect(await revert('Module1', 'Sub B()\r\nEnd Sub\r\n')).toEqual(['<deleted>']);
    });

    it('ends the review of a module the server deleted', async () => {
        const mirror = handlers();
        await mirror.agentEdit(edit());

        const answer = await mirror.agentEdit(edit({ after: '', afterExists: false }));

        expect(answer).toEqual({ shown: true, review: 'none' });
        expect(hasPendingAgentReview(file, 'Module1')).toBe(false);
    });

    it('carries a review to the module\'s new name', async () => {
        const mirror = handlers();
        await mirror.agentEdit(edit());

        const answer = await mirror.moduleRenamed({ file, from: 'Module1', to: 'Renamed' });

        expect(answer).toEqual({ shown: true, review: 'pending' });
        expect(hasPendingAgentReview(file, 'Module1')).toBe(false);
        expect(hasPendingAgentReview(file, 'Renamed')).toBe(true);
        expect(changed).toEqual([file, file]);
    });

    it('takes the new state of a change to the file that is not code, with no review', async () => {
        const answer = await handlers().fileChanged({ file, what: 'cells' });

        expect(answer).toEqual({ shown: true, review: 'none' });
        expect(changed).toEqual([file]);
    });

    it('does nothing for a file this window does not show', async () => {
        listed = false;
        const mirror = handlers();

        expect(await mirror.agentEdit(edit())).toEqual({ shown: false, review: 'none' });
        expect(await mirror.fileChanged({ file })).toEqual({ shown: false, review: 'none' });
        expect(changed).toEqual([]);
        expect(diffsOpened()).toHaveLength(0);
        expect(bridge.call).not.toHaveBeenCalled();
    });

    it('counts a file as shown when a module of it is open, listed in the tree or not', async () => {
        listed = false;
        hoisted.textDocuments.push({ uri: encodeModuleUri(file, 'Other'), isClosed: false });

        expect(await handlers().fileChanged({ file })).toEqual({ shown: true, review: 'none' });
        expect(changed).toEqual([file]);
    });

    it('with Review Agent Writes off, takes the new state and opens nothing', async () => {
        hoisted.settings['agent.showWriteDiffs'] = false;

        const answer = await handlers().agentEdit(edit());

        expect(answer).toEqual({ shown: true, review: 'none' });
        expect(changed).toEqual([file]);
        expect(diffsOpened()).toHaveLength(0);
    });

    it('keeps an earlier review tracking the module when Review Agent Writes is off', async () => {
        const mirror = handlers();
        await mirror.agentEdit(edit());
        hoisted.settings['agent.showWriteDiffs'] = false;
        bridge.call.mockResolvedValue({ source: 'Sub C()\r\nEnd Sub\r\n' });

        const answer = await mirror.agentEdit(edit({ before: 'Sub B()\r\nEnd Sub\r\n', after: 'Sub C()\r\nEnd Sub\r\n' }));

        expect(answer.review).toBe('pending');
        // Still the original before, and the module as it now reads is what Revert expects.
        expect(await revert('Module1', 'Sub C()\r\nEnd Sub\r\n')).toEqual(['Sub A()\r\nEnd Sub\r\n']);
    });

    it('takes two reports for one file in the order they came', async () => {
        let finishFirstRead: (value: { source: string }) => void = () => undefined;
        bridge.call
            .mockImplementationOnce(() => new Promise((resolve) => { finishFirstRead = resolve; }))
            .mockResolvedValueOnce({ source: 'Sub C()\r\nEnd Sub\r\n' });
        const mirror = handlers();

        const first = mirror.agentEdit(edit());
        const second = mirror.agentEdit(edit({ before: 'Sub B()\r\nEnd Sub\r\n', after: 'Sub C()\r\nEnd Sub\r\n' }));
        await new Promise((resolve) => setTimeout(resolve, 20));
        expect(bridge.call).toHaveBeenCalledTimes(1);
        finishFirstRead({ source: 'Sub B()\r\nEnd Sub\r\n' });
        await Promise.all([first, second]);

        // The review keeps the before of the first write, and follows the second.
        expect(await revert('Module1', 'Sub C()\r\nEnd Sub\r\n')).toEqual(['Sub A()\r\nEnd Sub\r\n']);
    });

    it('falls back to the server\'s after when XLIDE cannot read the module', async () => {
        bridge.call.mockRejectedValue(new Error('locked'));

        await handlers().agentEdit(edit({ after: '\r\nSub B()\r\nEnd Sub\r\n' }));

        expect(await revert('Module1', 'Sub B()\r\nEnd Sub\r\n')).toEqual(['Sub A()\r\nEnd Sub\r\n']);
    });

    it('says in the log what each report did', async () => {
        await handlers().agentEdit(edit());

        expect(log).toHaveBeenCalledWith(`MCP server: wrote Module1 in ${file} (shown here, review pending).`);
    });
});

describe('serving the API while mirroring is on', () => {
    let stateDir: string;

    beforeEach(() => {
        stateDir = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'xlide-mirror-api-'));
        hoisted.settings = {};
    });

    afterEach(() => {
        fs.rmSync(stateDir, { recursive: true, force: true });
    });

    function recordPath(): string {
        return nodePath.join(stateDir, xlideApiRecordName(process.pid));
    }

    async function until(check: () => boolean, what: string): Promise<void> {
        const deadline = Date.now() + 5000;
        while (!check()) {
            if (Date.now() > deadline) {
                throw new Error(what);
            }
            await new Promise((resolve) => setTimeout(resolve, 10));
        }
    }

    function changeSetting(value: boolean): void {
        hoisted.settings['agent.mirrorMcpEdits'] = value;
        for (const listener of [...hoisted.configListeners]) {
            listener({ affectsConfiguration: (section) => 'xlide.agent.mirrorMcpEdits'.startsWith(section) });
        }
    }

    it('listens from the start, stops when turned off, and listens again when turned on', async () => {
        const log = vi.fn();
        const mirror = mirrorMcpEdits({
            bridge: { call: vi.fn() },
            treeListsProject: async () => true,
            log,
            version: '10.7.2',
            stateDir,
        });
        try {
            await until(() => fs.existsSync(recordPath()), 'the record should be written');
            const first = JSON.parse(fs.readFileSync(recordPath(), 'utf8')) as { port: number; token: string };
            const hello = await fetch(`http://127.0.0.1:${first.port}/${first.token}/hello`);
            expect(await hello.json()).toEqual({ product: 'xlide_vscode', protocol: 1 });

            changeSetting(false);
            await until(() => !fs.existsSync(recordPath()), 'the record should go when mirroring is turned off');
            await expect(fetch(`http://127.0.0.1:${first.port}/${first.token}/hello`)).rejects.toThrow();

            changeSetting(true);
            await until(() => fs.existsSync(recordPath()), 'the record should come back when mirroring is turned on');
            const second = JSON.parse(fs.readFileSync(recordPath(), 'utf8')) as { token: string };
            expect(second.token).not.toBe(first.token);
        } finally {
            mirror.dispose();
        }
        expect(fs.existsSync(recordPath())).toBe(false);
        expect(log).toHaveBeenCalledWith(expect.stringMatching(/^Mirroring MCP server edits: listening on 127\.0\.0\.1:\d+\.$/));
    });

    it('never listens while mirroring is off', async () => {
        hoisted.settings['agent.mirrorMcpEdits'] = false;
        const mirror = mirrorMcpEdits({
            bridge: { call: vi.fn() },
            treeListsProject: async () => true,
            log: vi.fn(),
            version: '10.7.2',
            stateDir,
        });
        await new Promise((resolve) => setTimeout(resolve, 100));
        mirror.dispose();

        expect(fs.readdirSync(stateDir)).toEqual([]);
    });

    it('writes the record again when the workspace folders change', async () => {
        const mirror = mirrorMcpEdits({
            bridge: { call: vi.fn() },
            treeListsProject: async () => true,
            log: vi.fn(),
            version: '10.7.2',
            stateDir,
        });
        try {
            await until(() => fs.existsSync(recordPath()), 'the record should be written');
            const folders = vscode.workspace as unknown as { workspaceFolders: Array<{ uri: { fsPath: string } }> };
            folders.workspaceFolders = [{ uri: { fsPath: nodePath.resolve('/added') } }];
            hoisted.folderListeners.forEach((listener) => listener());

            const written = JSON.parse(fs.readFileSync(recordPath(), 'utf8')) as { workspaceFolders: string[] };
            expect(written.workspaceFolders).toEqual([nodePath.resolve('/added')]);
        } finally {
            mirror.dispose();
        }
    });
});
