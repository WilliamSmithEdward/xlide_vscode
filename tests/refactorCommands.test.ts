import { describe, expect, it, vi } from 'vitest';

const registered = vi.hoisted(() => new Map<string, (...args: unknown[]) => unknown>());

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock({
    commands: {
        registerCommand: vi.fn((name: string, handler: (...args: unknown[]) => unknown) => {
            registered.set(name, handler);
            return { dispose: vi.fn() };
        }),
    },
    window: {
        activeTextEditor: undefined,
        showQuickPick: vi.fn(async () => 'Helpers'),
        setStatusBarMessage: vi.fn(),
    },
    workspace: {
        applyEdit: vi.fn(async () => true),
    },
    WorkspaceEdit: class {
        replace(): void { /* the open module's own edits; not under test */ }
    },
}));
vi.mock('../src/projectModuleOperations', () => ({
    writeProjectModule: vi.fn(async () => ({ ok: true })),
    refreshProjectState: vi.fn(),
}));

import * as vscode from 'vscode';
import { registerRefactorCommands } from '../src/commands/refactorCommands';
import { refreshProjectState, writeProjectModule } from '../src/projectModuleOperations';
import { encodeModuleUri } from '../src/xlideFileSystem';

const PROJECT = process.platform === 'win32' ? 'C:\\work\\Reports.xlsm' : '/work/Reports.xlsm';

const REPORTS = [
    'Option Explicit',
    '',
    'Public Sub Build()',
    '    Debug.Print "built"',
    'End Sub',
    '',
].join('\r\n');

/** An editor on the Reports module with the caret in Build. */
function reportsEditor() {
    const caret = REPORTS.indexOf('Public Sub Build');
    const at = { line: 0, character: caret };
    return {
        document: {
            uri: encodeModuleUri(PROJECT, 'Reports'),
            languageId: 'vba',
            getText: () => REPORTS,
            offsetAt: () => caret,
            positionAt: () => at,
        },
        selection: { active: at, start: at, end: at },
        revealRange: vi.fn(),
    };
}

describe('Move to Module', () => {
    it('writes the module it moves to through the shared write path, then refreshes once', async () => {
        // A bare engine write skipped Office coordination, and left an open
        // copy of Helpers, the symbol index and the tree on the old text.
        const bridge = {
            call: vi.fn(async (method: string, params: { module?: string }) => {
                if (method === 'listModules') {
                    return [{ name: 'Reports' }, { name: 'Helpers' }];
                }
                if (method === 'readModule' && params.module === 'Helpers') {
                    return { source: 'Option Explicit\r\n' };
                }
                throw new Error(`unexpected engine call ${method}`);
            }),
        };
        const deps = { bridge, explorer: {}, fsProvider: {}, vbaIndex: {}, out: {}, context: {} };
        registerRefactorCommands(deps as never);
        (vscode.window as { activeTextEditor: unknown }).activeTextEditor = reportsEditor();

        await registered.get('xlide.refactor.moveToModule')!();

        expect(writeProjectModule).toHaveBeenCalledTimes(1);
        expect(writeProjectModule).toHaveBeenCalledWith(
            deps,
            { filePath: PROJECT, moduleName: 'Helpers', source: expect.stringContaining('Public Sub Build()') },
            { refreshProjectState: false },
        );
        expect(refreshProjectState).toHaveBeenCalledTimes(1);
        expect(refreshProjectState).toHaveBeenCalledWith(deps, PROJECT);
        expect(bridge.call).not.toHaveBeenCalledWith('writeModule', expect.anything());
    });
});
