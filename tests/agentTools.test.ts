import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

interface RegisteredTool {
    invoke(options: { input: Record<string, unknown>; toolInvocationToken?: unknown }, token: unknown): Promise<unknown>;
}

const vscodeMock = vi.hoisted(() => ({
    registeredTools: new Map<string, RegisteredTool>(),
    registeredCommands: new Map<string, (...args: unknown[]) => unknown>(),
    // Assigned by the vscode mock factory below, which runs with vi live.
    executeCommand: undefined as unknown as ReturnType<typeof vi.fn>,
    showInformationMessage: undefined as unknown as ReturnType<typeof vi.fn>,
    showWarningMessage: undefined as unknown as ReturnType<typeof vi.fn>,
}));

vi.mock('vscode', async () => {
    vscodeMock.executeCommand = vi.fn();
    vscodeMock.showInformationMessage = vi.fn(async (): Promise<unknown> => undefined);
    vscodeMock.showWarningMessage = vi.fn(async (): Promise<unknown> => undefined);
    return (await import('./helpers/vscodeMock')).vscodeMock({
        commands: {
            executeCommand: vscodeMock.executeCommand,
            registerCommand: vi.fn((name: string, handler: (...args: unknown[]) => unknown) => {
                vscodeMock.registeredCommands.set(name, handler);
                return { dispose: vi.fn() };
            }),
        },
        window: {
            showInformationMessage: vscodeMock.showInformationMessage,
            showWarningMessage: vscodeMock.showWarningMessage,
        },
        lm: {
            registerTool: vi.fn((name: string, tool: RegisteredTool) => {
                vscodeMock.registeredTools.set(name, tool);
                return { dispose: vi.fn() };
            }),
        },
    });
});

vi.mock('../src/xlsmExplorer', () => ({ XlsmExplorer: class XlsmExplorer {} }));
vi.mock('../src/xlideFileSystem', () => ({
    XlideFileSystemProvider: class XlideFileSystemProvider {},
    encodeModuleUri: vi.fn((filePath: string, moduleName: string) => ({
        path: `/${String(filePath).split('\\').join('/')}/${moduleName}.bas`,
        toString: () => `xlide-vba:///${moduleName}.bas`,
    })),
    notifySignatureDropped: vi.fn(),
    moduleIdentityKey: (name: string) => name.toLowerCase(),
    workbookIdentityKey: (filePath: string) => filePath.toLowerCase(),
}));
vi.mock('../src/vbaMemberCompletion', () => ({ invalidateVbaMemberCompletionCache: vi.fn() }));
vi.mock('../src/moduleExport', () => ({ exportWorkbookModules: vi.fn() }));
vi.mock('../src/workbookModuleSyncSettings', () => ({ setWorkbookModuleSyncExportMode: vi.fn() }));
vi.mock('../src/vbaWorkbookAnalysis', () => ({ analyzeWorkbook: vi.fn() }));
vi.mock('../src/vbaTestRunPipeline', () => ({ executeVbaTestRun: vi.fn() }));
vi.mock('../src/agentVbaTestArtifacts', () => ({ agentVbaTestArtifactPayloadFromPipeline: vi.fn() }));
vi.mock('../src/vbaTestRunner', () => ({
    describeVbaTestSelection: vi.fn(() => ''),
    summarizeVbaTestRun: vi.fn(),
}));

import { registerAgentTools } from '../src/agentTools';
import { hasPendingAgentReview, trackModuleWriteForAgentReview } from '../src/xlideAgentDiff';
import { writeWorkbookModule } from '../src/workbookModuleOperations';
import { clearXlideWriteAudit, recentXlideWriteAudits } from '../src/xlideWriteAudit';

function registerTools(bridgeCall: ReturnType<typeof vi.fn>) {
    vscodeMock.registeredTools.clear();
    vscodeMock.registeredCommands.clear();
    const explorer = { refresh: vi.fn(), refreshModuleSubs: vi.fn() };
    registerAgentTools(
        {} as never,
        { call: bridgeCall } as never,
        explorer as never,
        { notifyFileChanged: vi.fn() } as never,
        { invalidate: vi.fn() } as never,
    );
    return { explorer };
}

describe('xlide_createWorkbook agent tool', () => {
    let tempDir: string;

    beforeEach(() => {
        clearXlideWriteAudit();
        tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-agent-tools-'));
    });

    afterEach(() => {
        fs.rmSync(tempDir, { recursive: true, force: true });
    });

    it('refuses to overwrite an existing workbook and audits the failure', async () => {
        const existing = path.join(tempDir, 'Book.xlsm');
        fs.writeFileSync(existing, 'stub');
        const bridgeCall = vi.fn();
        const { explorer } = registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_createWorkbook');

        await expect(tool?.invoke({ input: { filePath: existing } }, undefined))
            .rejects.toThrow(/already exists/);

        expect(bridgeCall).not.toHaveBeenCalled();
        expect(explorer.refresh).not.toHaveBeenCalled();
        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide_createWorkbook',
            operation: 'create-workbook',
            outcome: 'failed',
            workbookPath: existing,
            summary: 'Create workbook: 0 changed, 1 failed',
        }]);
    });

    it('creates a new workbook and audits the success', async () => {
        const target = path.join(tempDir, 'New.xlsm');
        const bridgeCall = vi.fn(async () => ({ ok: true, path: target }));
        const { explorer } = registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_createWorkbook');

        await tool?.invoke({ input: { filePath: target } }, undefined);

        expect(bridgeCall).toHaveBeenCalledWith('createWorkbook', { path: target });
        expect(explorer.refresh).toHaveBeenCalled();
        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide_createWorkbook',
            operation: 'create-workbook',
            outcome: 'succeeded',
            workbookPath: target,
            summary: 'Create workbook: 1 changed',
        }]);
    });

    it('audits bridge failures during workbook creation', async () => {
        const target = path.join(tempDir, 'New.xlsm');
        const bridgeCall = vi.fn(async () => {
            throw new Error('workbook engine unavailable');
        });
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_createWorkbook');

        await expect(tool?.invoke({ input: { filePath: target } }, undefined))
            .rejects.toThrow('workbook engine unavailable');

        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide_createWorkbook',
            operation: 'create-workbook',
            outcome: 'failed',
            workbookPath: target,
        }]);
    });
});

describe('agent write review (diff + tree badge, native surfaces only)', () => {
    let tempDir: string;

    beforeEach(() => {
        clearXlideWriteAudit();
        tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-agent-review-'));
        vscodeMock.executeCommand.mockClear();
        vscodeMock.showInformationMessage.mockClear();
        vscodeMock.showWarningMessage.mockClear();
    });

    afterEach(() => {
        fs.rmSync(tempDir, { recursive: true, force: true });
    });

    /** Marks an invocation as chat-driven; only those get the review. */
    const CHAT = { toolInvocationToken: {} };

    /** In-memory module store speaking the bridge protocol. */
    function fakeEngine(initialByModule: Record<string, string> = {}) {
        const store = new Map<string, string>(
            Object.entries(initialByModule).map(([name, source]) => [name.toLowerCase(), source]),
        );
        const calls: Array<{ method: string; args: Record<string, unknown> }> = [];
        const call = vi.fn(async (method: string, args: Record<string, unknown>) => {
            calls.push({ method, args });
            const key = String(args.module ?? '').toLowerCase();
            switch (method) {
                case 'readModule': {
                    const source = store.get(key);
                    if (source === undefined) {
                        throw new Error(`module not found: ${String(args.module)}`);
                    }
                    return { source };
                }
                case 'writeModule':
                    store.set(key, String(args.source));
                    return { ok: true, signatureDropped: false };
                case 'renameModule': {
                    const source = store.get(key);
                    if (source === undefined) {
                        throw new Error(`module not found: ${String(args.module)}`);
                    }
                    store.delete(key);
                    store.set(String(args.newName).toLowerCase(), source);
                    return { ok: true, signatureDropped: false };
                }
                case 'deleteModule':
                    store.delete(key);
                    return { ok: true, signatureDropped: false };
                default:
                    return { ok: true };
            }
        });
        return { call, calls, store };
    }

    function writeTool(bridgeCall: ReturnType<typeof vi.fn>) {
        registerTools(bridgeCall);
        return vscodeMock.registeredTools.get('xlide_writeModule');
    }

    async function runCommand(name: string, node: { filePath: string; moduleName: string }) {
        const handler = vscodeMock.registeredCommands.get(name);
        expect(handler).toBeDefined();
        await handler?.(node);
    }

    async function settle() {
        await new Promise((resolve) => setTimeout(resolve, 10));
    }

    it('a chat-driven write opens a diff quietly and badges the module', async () => {
        const target = path.join(tempDir, 'Book.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Old()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();

        const diffCall = vscodeMock.executeCommand.mock.calls
            .find((call: unknown[]) => call[0] === 'vscode.diff');
        expect(diffCall).toBeDefined();
        expect(String(diffCall?.[1])).toContain('xlide-vba-before:');
        // Native surfaces only: no notification prompt, badge until decided.
        expect(vscodeMock.showInformationMessage).not.toHaveBeenCalled();
        expect(hasPendingAgentReview(target, 'Module1')).toBe(true);
    });

    it('a write without a chat token gets no review and skips the pre-read', async () => {
        const target = path.join(tempDir, 'Plain.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Old()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' } }, undefined);
        await settle();

        expect(vscodeMock.executeCommand.mock.calls.some((call: unknown[]) => call[0] === 'vscode.diff')).toBe(false);
        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
        expect(engine.calls.some((entry) => entry.method === 'readModule')).toBe(false);
    });

    it('Keep Agent Change clears the badge', async () => {
        const target = path.join(tempDir, 'Keep.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Old()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await runCommand('xlide.keepAgentChange', { filePath: target, moduleName: 'Module1' });

        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
        expect(engine.store.get('module1')).toContain('NewCode');
    });

    it('Revert restores the before-image through the audited write path', async () => {
        const target = path.join(tempDir, 'Revert.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Old()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await runCommand('xlide.revertAgentChange', { filePath: target, moduleName: 'Module1' });

        expect(engine.store.get('module1')).toContain('Sub Old()');
        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide.revertAgentChange',
            operation: 'write-module',
            outcome: 'succeeded',
            workbookPath: target,
            moduleName: 'Module1',
            summary: 'Revert agent change: 1 changed',
        }]);
    });

    it('stacked writes revert to the state before the first', async () => {
        const target = path.join(tempDir, 'Stacked.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Original()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub First()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub Second()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await runCommand('xlide.revertAgentChange', { filePath: target, moduleName: 'Module1' });

        expect(engine.store.get('module1')).toContain('Sub Original()');
        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
    });

    it('reverting a module the agent created deletes it', async () => {
        const target = path.join(tempDir, 'Created.xlsm');
        const engine = fakeEngine();
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub Fresh()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await runCommand('xlide.revertAgentChange', { filePath: target, moduleName: 'Module1' });

        expect(engine.calls.some((entry) => entry.method === 'deleteModule')).toBe(true);
        expect(engine.store.has('module1')).toBe(false);
        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide.revertAgentChange',
            operation: 'delete-module',
            outcome: 'succeeded',
            summary: 'Revert agent change: 1 removed',
        }]);
    });

    it('refuses to revert over a change made after the agent wrote', async () => {
        const target = path.join(tempDir, 'Drift.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Old()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        // The module changes again behind the review's back.
        engine.store.set('module1', 'Sub UserEdited()\r\nEnd Sub\r\n');
        await runCommand('xlide.revertAgentChange', { filePath: target, moduleName: 'Module1' });

        expect(vscodeMock.showWarningMessage).toHaveBeenCalledWith(
            expect.stringContaining('changed again after the agent'));
        expect(engine.store.get('module1')).toContain('UserEdited');
        expect(hasPendingAgentReview(target, 'Module1')).toBe(true);
    });

    it('a write through the shared operation path keeps the review revertable', async () => {
        // The user's report: an agent's second change arriving through another
        // surface (Copilot editing the open document, a sidebar write) froze
        // the review at the first write, so Revert refused with the drift
        // warning. Any XLIDE write path must keep the after-image current.
        const target = path.join(tempDir, 'Tracked.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Original()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);
        const ops = {
            bridge: { call: engine.call },
            explorer: { refresh: vi.fn(), refreshModuleSubs: vi.fn() },
            fsProvider: { notifyFileChanged: vi.fn() },
            vbaIndex: { invalidate: vi.fn() },
        };

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub First()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await writeWorkbookModule(ops as never, {
            filePath: target,
            moduleName: 'Module1',
            source: 'Sub Second()\r\nEnd Sub\r\n',
        });

        expect(hasPendingAgentReview(target, 'Module1')).toBe(true);
        await runCommand('xlide.revertAgentChange', { filePath: target, moduleName: 'Module1' });

        expect(vscodeMock.showWarningMessage).not.toHaveBeenCalled();
        expect(engine.store.get('module1')).toContain('Sub Original()');
        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
    });

    it('an editor save is tracked the same way (the FSP call shape)', async () => {
        const target = path.join(tempDir, 'Saved.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Original()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub First()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        // xlideFileSystem.writeFile stores the new source, then reports it.
        engine.store.set('module1', 'Sub Second()\r\nEnd Sub\r\n');
        trackModuleWriteForAgentReview(target, 'Module1', 'Sub Second()\r\nEnd Sub\r\n');
        await runCommand('xlide.revertAgentChange', { filePath: target, moduleName: 'Module1' });

        expect(vscodeMock.showWarningMessage).not.toHaveBeenCalled();
        expect(engine.store.get('module1')).toContain('Sub Original()');
        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
    });

    it('a write that lands back on the pre-agent original resolves the review', async () => {
        const original = 'Sub Original()\r\nEnd Sub\r\n';
        const target = path.join(tempDir, 'Undone.xlsm');
        const engine = fakeEngine({ Module1: original });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub First()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        expect(hasPendingAgentReview(target, 'Module1')).toBe(true);
        // The save restores exactly what the user had: nothing left to review.
        engine.store.set('module1', original);
        trackModuleWriteForAgentReview(target, 'Module1', original);

        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
    });

    it('an agent rewrite of the pre-agent original leaves nothing pending', async () => {
        const original = 'Sub Original()\r\nEnd Sub\r\n';
        const target = path.join(tempDir, 'SelfUndo.xlsm');
        const engine = fakeEngine({ Module1: original });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub First()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: original }, ...CHAT }, undefined);
        await settle();

        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
        expect(engine.store.get('module1')).toBe(original);
    });

    it('a rename carries the pending review to the new name', async () => {
        const target = path.join(tempDir, 'Rename.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Old()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        const rename = vscodeMock.registeredTools.get('xlide_renameModule');
        await rename?.invoke({ input: { filePath: target, moduleName: 'Module1', newName: 'Module2' } }, undefined);

        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
        expect(hasPendingAgentReview(target, 'Module2')).toBe(true);
    });

    it('deleting the module discards the pending review', async () => {
        const target = path.join(tempDir, 'Delete.xlsm');
        const engine = fakeEngine({ Module1: 'Sub Old()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        const del = vscodeMock.registeredTools.get('xlide_deleteModule');
        await del?.invoke({ input: { filePath: target, moduleName: 'Module1' } }, undefined);

        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
    });
});
