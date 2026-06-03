import { beforeEach, describe, expect, it, vi } from 'vitest';

const vscodeMock = vi.hoisted(() => ({
    findFiles: vi.fn(),
    showErrorMessage: vi.fn(),
}));

vi.mock('vscode', () => ({
    EventEmitter: class {
        event = vi.fn();
        fire = vi.fn();
    },
    MarkdownString: class {
        supportThemeIcons = false;
        constructor(readonly value = '') {}
        appendMarkdown = vi.fn();
    },
    ThemeIcon: class {
        constructor(readonly id: string) {}
    },
    TreeItem: class {
        id?: string;
        iconPath?: unknown;
        tooltip?: unknown;
        description?: string;
        contextValue?: string;
        command?: unknown;
        constructor(readonly label: string, readonly collapsibleState: number) {}
    },
    TreeItemCollapsibleState: {
        None: 0,
        Collapsed: 1,
        Expanded: 2,
    },
    window: {
        showErrorMessage: vscodeMock.showErrorMessage,
    },
    workspace: {
        findFiles: vscodeMock.findFiles,
        workspaceFolders: [{ uri: { fsPath: 'C:\\work' } }],
    },
}));

import { XlsmExplorer } from '../src/xlsmExplorer';

describe('XlsmExplorer', () => {
    beforeEach(() => {
        vscodeMock.findFiles.mockReset();
        vscodeMock.showErrorMessage.mockReset();
        vscodeMock.findFiles.mockResolvedValue([
            { scheme: 'file', fsPath: 'C:\\work\\Book.xlsm' },
        ]);
    });

    it('keeps the workbook tree empty until setup is complete', async () => {
        const explorer = new XlsmExplorer(fakeBridge());

        await expect(explorer.getChildren()).resolves.toEqual([]);
        expect(vscodeMock.findFiles).not.toHaveBeenCalled();

        explorer.setSetupComplete(true);
        await expect(explorer.getChildren()).resolves.toEqual([{
            kind: 'xlsm',
            label: 'Book.xlsm',
            filePath: 'C:\\work\\Book.xlsm',
        }]);
        expect(vscodeMock.findFiles).toHaveBeenCalledWith(
            '**/*.{xlsm,xlsb,xlam}',
            '{**/node_modules/**,**/.venv/**,**/venv/**}',
        );
    });
});

function fakeBridge() {
    return {
        call: vi.fn(),
    } as unknown as ConstructorParameters<typeof XlsmExplorer>[0];
}
