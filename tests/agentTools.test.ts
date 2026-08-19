import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

interface RegisteredTool {
    invoke(options: { input: Record<string, unknown> }, token: unknown): Promise<unknown>;
}

const vscodeMock = vi.hoisted(() => ({
    registeredTools: new Map<string, RegisteredTool>(),
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
            registerCommand: vi.fn(() => ({ dispose: vi.fn() })),
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
import { hasPendingAgentReview } from '../src/xlideAgentDiff';
import { clearXlideWriteAudit, recentXlideWriteAudits } from '../src/xlideWriteAudit';

function registerTools(bridgeCall: ReturnType<typeof vi.fn>) {
    vscodeMock.registeredTools.clear();
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

describe('agent write review (keep/revert with tree badge)', () => {
    let tempDir: string;

    beforeEach(() => {
        tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-agent-review-'));
        vscodeMock.executeCommand.mockClear();
        vscodeMock.showInformationMessage.mockReset();
        vscodeMock.showInformationMessage.mockResolvedValue(undefined);
    });

    afterEach(() => {
        fs.rmSync(tempDir, { recursive: true, force: true });
    });

    function writeTool(bridgeCall: ReturnType<typeof vi.fn>) {
        registerTools(bridgeCall);
        return vscodeMock.registeredTools.get('xlide_writeModule');
    }

    it('opens a diff, prompts, and marks the module pending until resolved', async () => {
        const target = path.join(tempDir, 'Book.xlsm');
        const bridgeCall = vi.fn(async (method: string) => {
            if (method === 'readModule') { return { source: 'Sub Old()\r\nEnd Sub\r\n' }; }
            return { ok: true, signatureDropped: false };
        });
        const tool = writeTool(bridgeCall);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' } }, undefined);
        await new Promise((resolve) => setTimeout(resolve, 0));

        const diffCall = vscodeMock.executeCommand.mock.calls
            .find((call: unknown[]) => call[0] === 'vscode.diff');
        expect(diffCall).toBeDefined();
        expect(String(diffCall?.[1])).toContain('xlide-vba-before:');
        expect(vscodeMock.showInformationMessage).toHaveBeenCalledWith(
            expect.stringContaining('Keep the change?'), 'Keep', 'Revert');
        // Dismissed prompt: still pending, so the tree badge stays.
        expect(hasPendingAgentReview(target, 'Module1')).toBe(true);
    });

    it('Keep resolves the pending badge', async () => {
        const target = path.join(tempDir, 'Keep.xlsm');
        vscodeMock.showInformationMessage.mockResolvedValue('Keep');
        const bridgeCall = vi.fn(async (method: string) => {
            if (method === 'readModule') { return { source: 'Sub Old()\r\nEnd Sub\r\n' }; }
            return { ok: true, signatureDropped: false };
        });
        const tool = writeTool(bridgeCall);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' } }, undefined);
        await new Promise((resolve) => setTimeout(resolve, 0));

        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
    });

    it('Revert writes the before-image back and resolves the badge', async () => {
        const target = path.join(tempDir, 'Revert.xlsm');
        vscodeMock.showInformationMessage.mockResolvedValue('Revert');
        const written: string[] = [];
        let currentSource = 'Sub Old()\r\nEnd Sub\r\n';
        const bridgeCall = vi.fn(async (method: string, args: Record<string, unknown>) => {
            if (method === 'readModule') { return { source: currentSource }; }
            if (method === 'writeModule') {
                currentSource = String(args.source);
                written.push(currentSource);
                return { ok: true, signatureDropped: false };
            }
            return { ok: true };
        });
        const tool = writeTool(bridgeCall);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' } }, undefined);
        await new Promise((resolve) => setTimeout(resolve, 10));

        expect(written[0]).toContain('NewCode');
        expect(written[written.length - 1]).toContain('Sub Old()');
        expect(hasPendingAgentReview(target, 'Module1')).toBe(false);
    });

    it('refuses to revert over a change made after the agent wrote', async () => {
        const target = path.join(tempDir, 'Drift.xlsm');
        vscodeMock.showInformationMessage.mockResolvedValue('Revert');
        const written: string[] = [];
        const responses = [
            { source: 'Sub Old()\r\nEnd Sub\r\n' },        // before-image read
            { source: 'Sub NewCode()\r\nEnd Sub\r\n' },     // after-image read
            { source: 'Sub UserEdited()\r\nEnd Sub\r\n' },  // read at revert time: drifted
        ];
        const bridgeCall = vi.fn(async (method: string, args: Record<string, unknown>) => {
            if (method === 'readModule') { return responses.shift() ?? { source: '' }; }
            if (method === 'writeModule') {
                written.push(String(args.source));
                return { ok: true, signatureDropped: false };
            }
            return { ok: true };
        });
        const tool = writeTool(bridgeCall);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub NewCode()\r\nEnd Sub\r\n' } }, undefined);
        await new Promise((resolve) => setTimeout(resolve, 10));

        // Only the agent's own write happened; the revert was refused.
        expect(written).toHaveLength(1);
        expect(vscodeMock.showWarningMessage).toHaveBeenCalledWith(
            expect.stringContaining('changed again after the agent'));
        expect(hasPendingAgentReview(target, 'Module1')).toBe(true);
    });
});
