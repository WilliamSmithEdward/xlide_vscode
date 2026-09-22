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

// The module each encoded URI names, so decoding one gives it back.
const encodedModules = vi.hoisted(() => new Map<unknown, { projectPath: string; moduleName: string }>());

vi.mock('../src/projectExplorer', () => ({ ProjectExplorer: class ProjectExplorer {} }));
vi.mock('../src/xlideFileSystem', () => ({
    XlideFileSystemProvider: class XlideFileSystemProvider {},
    XLIDE_SCHEME: 'xlide-vba',
    encodeModuleUri: vi.fn((filePath: string, moduleName: string) => {
        const uri = {
            path: `/${String(filePath).split('\\').join('/')}/${moduleName}.bas`,
            toString: () => `xlide-vba:///${moduleName}.bas`,
        };
        encodedModules.set(uri, { projectPath: filePath, moduleName });
        return uri;
    }),
    decodeModuleUri: vi.fn((uri: unknown) => {
        const decoded = encodedModules.get(uri);
        if (!decoded) {
            throw new Error('not a module URI');
        }
        return decoded;
    }),
    encodeFormMarkupUri: vi.fn((filePath: string, moduleName: string) => ({
        path: `/${String(filePath).split('\\').join('/')}/${moduleName}.form`,
        toString: () => `xlide-vba:///${moduleName}.form`,
    })),
    activeLocalVbaEditor: vi.fn(),
    notifySignatureDropped: vi.fn(),
    moduleIdentityKey: (name: string) => name.toLowerCase(),
    projectIdentityKey: (filePath: string) => filePath.toLowerCase(),
}));
vi.mock('../src/vbaMemberCompletion', () => ({ invalidateVbaMemberCompletionCache: vi.fn() }));
// Each write passes straight through: the real coordinator talks to whatever
// Office application is running.
vi.mock('../src/officeWriteCoordinator', async (original) => ({
    ...(await original<typeof import('../src/officeWriteCoordinator')>()),
    runWriteWithHostCoordination: vi.fn((_filePath: string, write: () => Promise<unknown>) => write()),
}));
vi.mock('../src/moduleExport', () => ({ exportProjectModules: vi.fn() }));
vi.mock('../src/projectModuleSyncSettings', () => ({ setProjectModuleSyncExportMode: vi.fn() }));
vi.mock('../src/vbaProjectWideAnalysis', () => ({ analyzeProject: vi.fn() }));
vi.mock('../src/vbaTestRunPipeline', () => ({ executeVbaTestRun: vi.fn() }));
vi.mock('../src/agentVbaTestArtifacts', () => ({ agentVbaTestArtifactPayloadFromPipeline: vi.fn() }));
vi.mock('../src/vbaTestRunner', () => ({
    describeVbaTestSelection: vi.fn(() => ''),
    summarizeVbaTestRun: vi.fn(),
}));

import * as vscode from 'vscode';
import { registerAgentTools } from '../src/agentTools';
import { hasPendingAgentReview, pendingAgentReviewModules, trackModuleWriteForAgentReview } from '../src/xlideAgentDiff';
import { writeProjectModule } from '../src/projectModuleOperations';
import { clearXlideWriteAudit, recentXlideWriteAudits } from '../src/xlideWriteAudit';
import { runWriteWithHostCoordination } from '../src/officeWriteCoordinator';

function registerTools(bridgeCall: ReturnType<typeof vi.fn>) {
    vscodeMock.registeredTools.clear();
    vscodeMock.registeredCommands.clear();
    const explorer = { refresh: vi.fn(), refreshAgentReviewMarks: vi.fn() };
    registerAgentTools(
        {} as never,
        { call: bridgeCall } as never,
        explorer as never,
        { notifyFileChanged: vi.fn() } as never,
        { invalidate: vi.fn() } as never,
    );
    return { explorer };
}

describe('xlide_gitChanges agent tool', () => {
    const PROJECT = process.platform === 'win32' ? 'C:\\work\\Book.xlsm' : '/work/Book.xlsm';
    const ROOT = process.platform === 'win32' ? 'C:/work' : '/work';

    function registerWithGit(overrides: { inRepo?: boolean } = {}) {
        vscodeMock.registeredTools.clear();
        const gitDeps = {
            git: {
                run: vi.fn(async (args: readonly string[]) => ({
                    code: args[0] === 'rev-parse' && overrides.inRepo === false ? 128 : 0,
                    stdout: Buffer.from(args[0] === 'rev-parse' ? `${ROOT}\n` : ''),
                    stderr: '',
                })),
            },
            currentModules: vi.fn(async () => [{ name: 'Module1', source: 'Sub A()\r\n    x = 2\r\nEnd Sub\r\n' }]),
            modulesAtRevision: vi.fn(async () => [{ name: 'Module1', source: 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n' }]),
        };
        registerAgentTools(
            {} as never,
            { call: vi.fn() } as never,
            { refresh: vi.fn(), refreshAgentReviewMarks: vi.fn() } as never,
            { notifyFileChanged: vi.fn() } as never,
            { invalidate: vi.fn() } as never,
            gitDeps as never,
        );
        return { gitDeps, tool: vscodeMock.registeredTools.get('xlide_gitChanges')! };
    }

    async function report(tool: RegisteredTool, input: Record<string, unknown>) {
        // The mock's LanguageModelToolResult keeps its parts under `parts`.
        const result = await tool.invoke({ input }, undefined) as { parts: Array<{ value: string }> };
        return JSON.parse(result.parts[0].value) as {
            tracked: boolean;
            reason?: string;
            changes: Array<{ name: string; kind: string; diff: string }>;
        };
    }

    it('hands an agent one unified diff per changed module, against HEAD by default', async () => {
        const { gitDeps, tool } = registerWithGit();
        const answer = await report(tool, { filePath: PROJECT });

        expect(answer.tracked).toBe(true);
        expect(answer.changes).toHaveLength(1);
        expect(answer.changes[0]).toMatchObject({ name: 'Module1', kind: 'modified' });
        expect(answer.changes[0].diff).toContain('--- Module1 (HEAD)');
        expect(answer.changes[0].diff).toContain('-    x = 1');
        expect(answer.changes[0].diff).toContain('+    x = 2');
        expect(gitDeps.modulesAtRevision).toHaveBeenCalledWith(PROJECT, expect.objectContaining({ relativePath: 'Book.xlsm' }), 'HEAD');
    });

    it('compares against the revision it is given', async () => {
        const { gitDeps, tool } = registerWithGit();
        await report(tool, { filePath: PROJECT, revision: 'v1.0' });
        expect(gitDeps.modulesAtRevision).toHaveBeenCalledWith(PROJECT, expect.anything(), 'v1.0');
    });

    it('answers with a reason, not an error, for a file outside any repository', async () => {
        const { tool } = registerWithGit({ inRepo: false });
        const answer = await report(tool, { filePath: PROJECT });
        expect(answer).toMatchObject({ tracked: false, changes: [] });
        expect(answer.reason).toContain('not inside a git repository');
    });
});

describe('xlide_createProject agent tool', () => {
    let tempDir: string;

    beforeEach(() => {
        clearXlideWriteAudit();
        tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlide-agent-tools-'));
    });

    afterEach(() => {
        fs.rmSync(tempDir, { recursive: true, force: true });
    });

    it('refuses to overwrite an existing project and audits the failure', async () => {
        const existing = path.join(tempDir, 'Book.xlsm');
        fs.writeFileSync(existing, 'stub');
        const bridgeCall = vi.fn();
        const { explorer } = registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_createProject');

        await expect(tool?.invoke({ input: { filePath: existing } }, undefined))
            .rejects.toThrow(/already exists/);

        expect(bridgeCall).not.toHaveBeenCalled();
        expect(explorer.refresh).not.toHaveBeenCalled();
        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide_createProject',
            operation: 'create-project',
            outcome: 'failed',
            projectPath: existing,
            summary: 'Create project: 0 changed, 1 failed',
        }]);
    });

    it('creates a new project and audits the success', async () => {
        const target = path.join(tempDir, 'New.xlsm');
        const bridgeCall = vi.fn(async () => ({ ok: true, path: target }));
        const { explorer } = registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_createProject');

        await tool?.invoke({ input: { filePath: target } }, undefined);

        expect(bridgeCall).toHaveBeenCalledWith('createProject', { path: target });
        expect(explorer.refresh).toHaveBeenCalled();
        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide_createProject',
            operation: 'create-project',
            outcome: 'succeeded',
            projectPath: target,
            summary: 'Create project: 1 changed',
        }]);
    });

    it('audits bridge failures during project creation', async () => {
        const target = path.join(tempDir, 'New.xlsm');
        const bridgeCall = vi.fn(async () => {
            throw new Error('workbook engine unavailable');
        });
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_createProject');

        await expect(tool?.invoke({ input: { filePath: target } }, undefined))
            .rejects.toThrow('workbook engine unavailable');

        expect(recentXlideWriteAudits(1)).toMatchObject([{
            command: 'xlide_createProject',
            operation: 'create-project',
            outcome: 'failed',
            projectPath: target,
        }]);
    });
});

describe('xlide_writeCells agent tool', () => {
    it('writes only inside the Office coordination a module write gets', async () => {
        // With the workbook open in Excel, the file is locked: the coordinator
        // is what closes and reopens it as the user's setting says. The cell
        // write used to go to the engine directly and fail on the lock.
        const book = 'C:\\work\\Book.xlsm';
        const bridgeCall = vi.fn(async () => ({ ok: true }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_writeCells')!;
        vi.mocked(runWriteWithHostCoordination).mockClear();
        vi.mocked(runWriteWithHostCoordination).mockImplementationOnce(async () => 'held back');

        await tool.invoke({ input: { filePath: book, sheet: 'Sheet1', startCell: 'A1', data: [['x']] } }, undefined);

        expect(runWriteWithHostCoordination).toHaveBeenCalledWith(book, expect.any(Function));
        expect(bridgeCall).not.toHaveBeenCalledWith('writeCells', expect.anything());

        await tool.invoke({ input: { filePath: book, sheet: 'Sheet1', startCell: 'A1', data: [['x']] } }, undefined);

        expect(bridgeCall).toHaveBeenCalledWith('writeCells', { path: book, sheet: 'Sheet1', startCell: 'A1', data: [['x']] });
    });
});

describe('a listing tool over a file that holds no code', () => {
    const book = 'C:\\work\\Book.xlsm';

    const bridgeFor = (hasVbaProject: boolean) => vi.fn(async (method: string) => {
        if (method === 'listModules') { return []; }
        if (method === 'hasVbaProject') { return { hasVbaProject }; }
        return { modules: [], sheets: [{ name: 'Sheet1', dimensions: 'A1:A1' }], namedRanges: [] };
    });

    const listing = async (tool: string, bridgeCall: ReturnType<typeof vi.fn>): Promise<string> => {
        registerTools(bridgeCall);
        const result = await vscodeMock.registeredTools.get(tool)!
            .invoke({ input: { filePath: book } }, undefined) as { parts: Array<{ value: string }> };
        return result.parts[0].value;
    };

    it('tells an agent that an empty answer is the state of the file, not a failure', async () => {
        // A bare `[]` cannot be told from a read that went wrong, and an
        // agent that reads it as one retries or gives up.
        const answer = await listing('xlide_listModules', bridgeFor(false));
        expect(JSON.parse(answer.split('\n\n')[0])).toEqual([]);
        expect(answer).toContain('no VBA project in it at all');
        expect(answer).toContain('This is not an error');
        expect(answer).toContain('Excel');
    });

    it('distinguishes a project with nothing in it, which will take a module', async () => {
        const answer = await listing('xlide_listModules', bridgeFor(true));
        expect(answer).toContain('no modules in it yet');
        expect(answer).toContain('xlide_writeModule adds the first one');
    });

    it('says the same on xlide_getProjectInfo, whose sheets still answer', async () => {
        const answer = await listing('xlide_getProjectInfo', bridgeFor(false));
        expect(answer).toContain('Sheet1');
        expect(answer).toContain('no VBA project in it at all');
    });

    it('says nothing extra when the project has modules', async () => {
        const bridgeCall = vi.fn(async (method: string) => (
            method === 'listModules' ? [{ name: 'Module1', type: 'standard' }] : {}
        ));
        const answer = await listing('xlide_listModules', bridgeCall);
        expect(answer).toBe(JSON.stringify([{ name: 'Module1', type: 'standard' }], null, 2));
        expect(bridgeCall.mock.calls.filter(([method]) => method === 'hasVbaProject')).toHaveLength(0);
    });
});

describe('xlide_addReference agent tool', () => {
    const book = 'C:\\work\\Book.xlsm';

    it('writes inside the Office coordination, like every other project write', async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, added: true, name: 'Word' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_addReference')!;
        vi.mocked(runWriteWithHostCoordination).mockClear();

        const result = await tool.invoke(
            { input: { filePath: book, library: 'word' } }, undefined,
        ) as { parts: Array<{ value: string }> };

        expect(runWriteWithHostCoordination).toHaveBeenCalledWith(book, expect.any(Function));
        expect(bridgeCall).toHaveBeenCalledWith('addReference', { path: book, library: 'word' });
        expect(result.parts[0].value).toContain('Word');
    });

    it('says so rather than claiming a change when the project already has it', async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, added: false, name: 'Word' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_addReference')!;

        const result = await tool.invoke(
            { input: { filePath: book, library: 'word' } }, undefined,
        ) as { parts: Array<{ value: string }> };

        expect(result.parts[0].value).toContain('already references Word');
    });
});

describe('xlide_removeReference agent tool', () => {
    const book = 'C:\\work\\Book.xlsm';

    it('writes inside the Office coordination, and says what stops compiling', async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, removed: true, name: 'Word' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_removeReference')!;
        vi.mocked(runWriteWithHostCoordination).mockClear();

        const result = await tool.invoke(
            { input: { filePath: book, library: 'Word' } }, undefined,
        ) as { parts: Array<{ value: string }> };

        expect(runWriteWithHostCoordination).toHaveBeenCalledWith(book, expect.any(Function));
        expect(bridgeCall).toHaveBeenCalledWith('removeReference', { path: book, library: 'Word' });
        expect(result.parts[0].value).toContain('no longer compiles');
    });

    it('says so rather than claiming a change when the project never had it', async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, removed: false, name: 'Word' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_removeReference')!;

        const result = await tool.invoke(
            { input: { filePath: book, library: 'Word' } }, undefined,
        ) as { parts: Array<{ value: string }> };

        expect(result.parts[0].value).toContain('does not reference Word');
    });
});

describe('shape agent tools', () => {
    const book = 'C:\\work\\Book.xlsm';
    const deck = 'C:\\work\\Deck.pptm';
    const doc = 'C:\\work\\Report.docm';

    it('lists shapes for one surface or all, as JSON', async () => {
        const surfaces = [{ surface: 'Sheet1', shapes: [{ name: 'Go', kind: 'button', macro: 'DoIt' }] }];
        const bridgeCall = vi.fn(async () => ({ surfaces }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_listShapes')!;

        const result = await tool.invoke({ input: { filePath: book, surface: 'Sheet1' } }, undefined) as { parts: Array<{ value: string }> };
        await tool.invoke({ input: { filePath: book } }, undefined);

        expect(bridgeCall).toHaveBeenNthCalledWith(1, 'listShapes', { path: book, surface: 'Sheet1' }, undefined);
        expect(bridgeCall).toHaveBeenNthCalledWith(2, 'listShapes', { path: book }, undefined);
        expect(JSON.parse(result.parts[0].value)).toEqual(surfaces);
    });

    it('still takes sheet, the name the parameter had before the other hosts', async () => {
        const bridgeCall = vi.fn(async () => ({ surfaces: [] }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_listShapes')!;

        await tool.invoke({ input: { filePath: book, sheet: 'Sheet1' } }, undefined);

        expect(bridgeCall).toHaveBeenCalledWith('listShapes', { path: book, surface: 'Sheet1' }, undefined);
    });

    it('edits a shape only inside the Office coordination, and says what changed', async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, name: 'Go' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_editShape')!;
        vi.mocked(runWriteWithHostCoordination).mockClear();
        vi.mocked(runWriteWithHostCoordination).mockImplementationOnce(async () => 'held back');
        const input = { filePath: book, surface: 'Sheet1', action: 'update', name: 'Go', macro: 'Macros.Run' };

        await tool.invoke({ input }, undefined);
        expect(runWriteWithHostCoordination).toHaveBeenCalledWith(book, expect.any(Function));
        expect(bridgeCall).not.toHaveBeenCalled();

        const result = await tool.invoke({ input }, undefined) as { parts: Array<{ value: string }> };
        expect(bridgeCall).toHaveBeenCalledWith('editShape', { path: book, surface: 'Sheet1', action: 'update', name: 'Go', macro: 'Macros.Run' });
        expect(result.parts[0].value).toBe('Change shape: 1 changed\nShape "Go" changed on "Sheet1" in "C:\\work\\Book.xlsm".');
    });

    it('passes a slide and the points that place a shape on it', async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, name: 'Badge' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_editShape')!;

        await tool.invoke({
            input: { filePath: deck, surface: 'Slide 2', action: 'add', type: 'oval', left: 40, top: 50, macro: 'SayHello' },
        }, undefined);

        expect(bridgeCall).toHaveBeenCalledWith('editShape', {
            path: deck, surface: 'Slide 2', action: 'add', type: 'oval', left: 40, top: 50, macro: 'SayHello',
        });
    });

    it("leaves Word's surface empty when none is named, and says so without a sheet", async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, name: 'Stamp' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_editShape')!;

        const result = await tool.invoke({
            input: { filePath: doc, action: 'add', type: 'rectangle', name: 'Stamp' },
        }, undefined) as { parts: Array<{ value: string }> };

        expect(bridgeCall).toHaveBeenCalledWith('editShape', { path: doc, surface: '', action: 'add', type: 'rectangle', name: 'Stamp' });
        expect(result.parts[0].value).toBe('Add shape: 1 changed\nShape "Stamp" added in "C:\\work\\Report.docm".');
    });

    it('names the shape an add created, since the host chooses the name', async () => {
        const bridgeCall = vi.fn(async () => ({ ok: true, name: 'Button 4' }));
        registerTools(bridgeCall);
        const tool = vscodeMock.registeredTools.get('xlide_editShape')!;

        const result = await tool.invoke({ input: { filePath: book, surface: 'Sheet1', action: 'add', type: 'button', range: 'B2:C3' } }, undefined) as { parts: Array<{ value: string }> };

        expect(result.parts[0].value).toBe('Add shape: 1 changed\nShape "Button 4" added on "Sheet1" in "C:\\work\\Book.xlsm".');
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
        const classes = new Set<string>();
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
                    if (!store.has(key) && args.kind === 'class') {
                        classes.add(key);
                    }
                    store.set(key, String(args.source));
                    return { ok: true, signatureDropped: false };
                case 'listModules':
                    return [...store.keys()].map((name) => ({ name, type: classes.has(name) ? 'class' : 'standard' }));
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

    it("creates a class module when asked with kind='class', as the description says", async () => {
        // The description offered kind='class' and the tool never read it:
        // every module an agent created was a standard module.
        const target = path.join(tempDir, 'Kinds.xlsm');
        const engine = fakeEngine({ Helper: 'Sub A()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Person', source: 'Public Name As String\r\n', kind: 'class' } }, undefined);
        const write = engine.calls.find((entry) => entry.method === 'writeModule');
        expect(write?.args.kind).toBe('class');

        const refused = await tool?.invoke({ input: { filePath: target, moduleName: 'helper', source: 'Private m As Long\r\n', kind: 'class' } }, undefined);
        expect(JSON.stringify(refused)).toContain('is already a standard module');
        expect(engine.store.get('helper')).toBe('Sub A()\r\nEnd Sub\r\n');

        const bad = await tool?.invoke({ input: { filePath: target, moduleName: 'Other', source: '', kind: 'form' } }, undefined);
        expect(JSON.stringify(bad)).toContain("kind must be 'standard' or 'class'");
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
            projectPath: target,
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
            explorer: { refresh: vi.fn(), refreshAgentReviewMarks: vi.fn() },
            fsProvider: { notifyFileChanged: vi.fn() },
            vbaIndex: { invalidate: vi.fn() },
        };

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub First()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await writeProjectModule(ops as never, {
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
        // The project's list, which colours and counts the project row,
        // follows the rename too.
        expect(pendingAgentReviewModules(target)).toEqual(['Module2']);
    });

    it('lists a project\'s pending modules, and none for another project', async () => {
        const target = path.join(tempDir, 'Listing.xlsm');
        const engine = fakeEngine({ Module1: 'Sub A()\r\nEnd Sub\r\n', Module2: 'Sub B()\r\nEnd Sub\r\n' });
        const tool = writeTool(engine.call);

        await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub A2()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();
        await tool?.invoke({ input: { filePath: target, moduleName: 'Module2', source: 'Sub B2()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
        await settle();

        expect(pendingAgentReviewModules(target).sort()).toEqual(['Module1', 'Module2']);
        expect(pendingAgentReviewModules(path.join(tempDir, 'Elsewhere.xlsm'))).toEqual([]);
        if (process.platform === 'win32') {
            // Windows paths match whatever their case; a tool and the tree
            // can spell the same workbook differently.
            expect(pendingAgentReviewModules(target.toUpperCase()).sort()).toEqual(['Module1', 'Module2']);
        }
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
        expect(pendingAgentReviewModules(target)).toEqual([]);
    });

    // An agent often tests with throwaway work: a scratch module it creates
    // and deletes, a temporary edit it undoes. Each write opened a review
    // diff, and the diff stayed open after the change was gone - titled for a
    // module that no longer existed, or showing no difference at all. A review
    // diff now closes once what it shows is gone. Keep leaves it: the change
    // is still there.
    describe('review diffs of changes that are gone', () => {
        interface FakeTab { input: unknown; isDirty: boolean }
        let tabs: FakeTab[];

        beforeEach(() => {
            tabs = [];
            const tabGroups = vscode.window.tabGroups as unknown as {
                all: Array<{ tabs: FakeTab[] }>;
                close: ReturnType<typeof vi.fn>;
            };
            tabGroups.all = [{ tabs }];
            tabGroups.close = vi.fn(async (closing: FakeTab | FakeTab[]) => {
                for (const tab of Array.isArray(closing) ? closing : [closing]) {
                    tabs.splice(tabs.indexOf(tab), 1);
                }
                return true;
            });
            // VS Code opens a diff tab for every `vscode.diff`.
            vscodeMock.executeCommand.mockImplementation(async (command: string, original: unknown, modified: unknown) => {
                if (command === 'vscode.diff') {
                    tabs.push({ input: new vscode.TabInputTextDiff(original as never, modified as never), isDirty: false });
                }
            });
        });

        afterEach(() => {
            vscodeMock.executeCommand.mockReset();
        });

        const reviewDiffsOf = (moduleName: string): FakeTab[] =>
            tabs.filter((tab) => String((tab.input as { modified: unknown }).modified) === `xlide-vba:///${moduleName}.bas`);

        it('close once the agent deletes the scratch module it created', async () => {
            const target = path.join(tempDir, 'Scratch.xlsm');
            const engine = fakeEngine();
            const tool = writeTool(engine.call);

            await tool?.invoke({ input: { filePath: target, moduleName: 'TmpCheck', source: 'Sub Probe()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
            await settle();
            expect(reviewDiffsOf('TmpCheck')).toHaveLength(1);

            await vscodeMock.registeredTools.get('xlide_deleteModule')
                ?.invoke({ input: { filePath: target, moduleName: 'TmpCheck' } }, undefined);
            await settle();

            expect(reviewDiffsOf('TmpCheck')).toEqual([]);
            expect(hasPendingAgentReview(target, 'TmpCheck')).toBe(false);
        });

        it('close once the agent puts the module back as it was', async () => {
            const original = 'Sub Original()\r\nEnd Sub\r\n';
            const target = path.join(tempDir, 'Undone.xlsm');
            const engine = fakeEngine({ Module1: original });
            const tool = writeTool(engine.call);

            await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub Original()\r\n    Debug.Print 1\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
            await settle();
            expect(reviewDiffsOf('Module1')).toHaveLength(1);
            await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: original }, ...CHAT }, undefined);
            await settle();

            expect(reviewDiffsOf('Module1')).toEqual([]);
        });

        it('close once a save puts the module back as it was', async () => {
            const original = 'Sub Original()\r\nEnd Sub\r\n';
            const target = path.join(tempDir, 'SavedBack.xlsm');
            const engine = fakeEngine({ Module1: original });
            const tool = writeTool(engine.call);

            await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub First()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
            await settle();
            engine.store.set('module1', original);
            trackModuleWriteForAgentReview(target, 'Module1', original);
            await settle();

            expect(reviewDiffsOf('Module1')).toEqual([]);
        });

        it('close once the change is reverted, and stay open once it is kept', async () => {
            const target = path.join(tempDir, 'Decided.xlsm');
            const engine = fakeEngine({ Module1: 'Sub A()\r\nEnd Sub\r\n', Module2: 'Sub B()\r\nEnd Sub\r\n' });
            const tool = writeTool(engine.call);
            await tool?.invoke({ input: { filePath: target, moduleName: 'Module1', source: 'Sub A2()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
            await settle();
            await tool?.invoke({ input: { filePath: target, moduleName: 'Module2', source: 'Sub B2()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
            await settle();

            await runCommand('xlide.revertAgentChange', { filePath: target, moduleName: 'Module1' });
            await runCommand('xlide.keepAgentChange', { filePath: target, moduleName: 'Module2' });
            await settle();

            expect(reviewDiffsOf('Module1')).toEqual([]);
            expect(reviewDiffsOf('Module2')).toHaveLength(1);
        });

        it('leave another module s diff, and a diff with unsaved edits, open', async () => {
            const target = path.join(tempDir, 'Others.xlsm');
            const engine = fakeEngine();
            const tool = writeTool(engine.call);
            await tool?.invoke({ input: { filePath: target, moduleName: 'TmpOne', source: 'Sub One()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
            await settle();
            await tool?.invoke({ input: { filePath: target, moduleName: 'TmpTwo', source: 'Sub Two()\r\nEnd Sub\r\n' }, ...CHAT }, undefined);
            await settle();
            // The user typed into the live side of TmpTwo's diff.
            reviewDiffsOf('TmpTwo')[0].isDirty = true;

            const del = vscodeMock.registeredTools.get('xlide_deleteModule');
            await del?.invoke({ input: { filePath: target, moduleName: 'TmpOne' } }, undefined);
            await del?.invoke({ input: { filePath: target, moduleName: 'TmpTwo' } }, undefined);
            await settle();

            expect(reviewDiffsOf('TmpOne')).toEqual([]);
            expect(reviewDiffsOf('TmpTwo')).toHaveLength(1);
        });
    });
});
