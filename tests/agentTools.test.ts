import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

interface RegisteredTool {
    invoke(options: { input: Record<string, unknown> }, token: unknown): Promise<unknown>;
}

const vscodeMock = vi.hoisted(() => ({
    registeredTools: new Map<string, RegisteredTool>(),
}));

vi.mock('vscode', () => ({
    lm: {
        registerTool: vi.fn((name: string, tool: RegisteredTool) => {
            vscodeMock.registeredTools.set(name, tool);
            return { dispose: vi.fn() };
        }),
    },
    LanguageModelToolResult: class {
        constructor(readonly parts: unknown[]) {}
    },
    LanguageModelTextPart: class {
        constructor(readonly value: string) {}
    },
    MarkdownString: class {
        constructor(readonly value = '') {}
    },
    workspace: {
        findFiles: vi.fn(async () => []),
    },
    window: {
        showWarningMessage: vi.fn(),
    },
}));

vi.mock('../src/pythonBridge', () => ({ PythonBridge: class PythonBridge {} }));
vi.mock('../src/xlsmExplorer', () => ({ XlsmExplorer: class XlsmExplorer {} }));
vi.mock('../src/xlideFileSystem', () => ({
    XlideFileSystemProvider: class XlideFileSystemProvider {},
    encodeModuleUri: vi.fn(() => ({ path: '/stub' })),
    notifySignatureDropped: vi.fn(),
}));
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
import { clearXlideWriteAudit, recentXlideWriteAudits } from '../src/xlideWriteAudit';

function registerTools(bridgeCall: ReturnType<typeof vi.fn>) {
    vscodeMock.registeredTools.clear();
    const explorer = { refresh: vi.fn() };
    registerAgentTools(
        {} as never,
        { call: bridgeCall } as never,
        explorer as never,
        { notifyFileChanged: vi.fn() } as never,
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
            throw new Error('python bridge unavailable');
        });
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_createWorkbook');

        await expect(tool?.invoke({ input: { filePath: target } }, undefined))
            .rejects.toThrow('python bridge unavailable');

        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide_createWorkbook',
            operation: 'create-workbook',
            outcome: 'failed',
            workbookPath: target,
        }]);
    });
});
