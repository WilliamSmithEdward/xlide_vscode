import { afterEach, describe, expect, it, vi } from 'vitest';
import * as vscode from 'vscode';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import {
    AGENT_REVIEW_COLOR_ID,
    AGENT_REVIEW_DECORATION_SCHEME,
    AgentReviewDecorationProvider,
    moduleDecorationUri,
    parseDecorationUri,
    pendingCountBadge,
    projectDecorationUri,
} from '../src/agentReviewDecorations';
import { keepAgentChange, presentAgentModuleWrite } from '../src/xlideAgentDiff';

const WORKBOOK = 'C:\\Work\\Ledger Book.xlsm';

async function agentEdits(filePath: string, moduleName: string): Promise<void> {
    await presentAgentModuleWrite(filePath, moduleName, {
        before: 'Sub Old()\r\nEnd Sub\r\n',
        beforeExisted: true,
        after: 'Sub New()\r\nEnd Sub\r\n',
    });
}

describe('decoration URIs', () => {
    it('carry a project path back out, colons and separators included', () => {
        expect(parseDecorationUri(projectDecorationUri(WORKBOOK))).toEqual({
            kind: 'project',
            filePath: WORKBOOK,
        });
    });

    it('carry a module back out', () => {
        expect(parseDecorationUri(moduleDecorationUri(WORKBOOK, 'Ledger'))).toEqual({
            kind: 'module',
            filePath: WORKBOOK,
            moduleName: 'Ledger',
        });
    });

    it('keep the identity out of URI structure', () => {
        const uri = moduleDecorationUri(WORKBOOK, 'Ledger');

        expect(uri.scheme).toBe(AGENT_REVIEW_DECORATION_SCHEME);
        expect(uri.path).toMatch(/^\/module\/[A-Za-z0-9_-]+$/);
    });

    it('answer nothing for a URI this module did not build', () => {
        const from = (path: string, scheme = AGENT_REVIEW_DECORATION_SCHEME) => vscode.Uri.from({ scheme, path });

        expect(parseDecorationUri(from('/module/abc', 'file'))).toBeUndefined();
        expect(parseDecorationUri(from('/elsewhere/abc'))).toBeUndefined();
        expect(parseDecorationUri(from('/module/not%20base64'))).toBeUndefined();
        // A module payload must carry both parts, a project payload just one.
        const project = projectDecorationUri(WORKBOOK).path.split('/')[2];
        expect(parseDecorationUri(from(`/module/${project}`))).toBeUndefined();
    });
});

describe('the count badge', () => {
    it('fits the two characters a decoration badge allows', () => {
        expect(pendingCountBadge(1)).toBe('1');
        expect(pendingCountBadge(9)).toBe('9');
        expect(pendingCountBadge(10)).toBe('9+');
        expect(pendingCountBadge(250)).toBe('9+');
    });
});

describe('AgentReviewDecorationProvider', () => {
    const provider = new AgentReviewDecorationProvider();

    afterEach(() => {
        keepAgentChange(WORKBOOK, 'Ledger');
        keepAgentChange(WORKBOOK, 'Reports');
    });

    it('colours and badges a module only while its edit awaits review', async () => {
        const uri = moduleDecorationUri(WORKBOOK, 'Ledger');
        expect(provider.provideFileDecoration(uri)).toBeUndefined();

        await agentEdits(WORKBOOK, 'Ledger');
        const decoration = provider.provideFileDecoration(uri);

        expect(decoration?.badge).toBe('AI');
        expect((decoration?.color as unknown as { id: string }).id).toBe(AGENT_REVIEW_COLOR_ID);

        keepAgentChange(WORKBOOK, 'Ledger');
        expect(provider.provideFileDecoration(uri)).toBeUndefined();
    });

    it('counts the pending modules on the project row', async () => {
        const uri = projectDecorationUri(WORKBOOK);
        expect(provider.provideFileDecoration(uri)).toBeUndefined();

        await agentEdits(WORKBOOK, 'Ledger');
        expect(provider.provideFileDecoration(uri)?.badge).toBe('1');
        expect(provider.provideFileDecoration(uri)?.tooltip).toBe('1 agent edit awaiting review');

        await agentEdits(WORKBOOK, 'Reports');
        expect(provider.provideFileDecoration(uri)?.badge).toBe('2');
        expect(provider.provideFileDecoration(uri)?.tooltip).toBe('2 agent edits awaiting review');
    });

    it('leaves another project alone', async () => {
        await agentEdits(WORKBOOK, 'Ledger');

        expect(provider.provideFileDecoration(projectDecorationUri('C:\\Work\\Other.xlsm'))).toBeUndefined();
    });

    it('asks for every row again when a review starts or ends', async () => {
        const fired = vi.fn();
        const listener = provider.onDidChangeFileDecorations(fired);

        await agentEdits(WORKBOOK, 'Ledger');
        keepAgentChange(WORKBOOK, 'Ledger');

        expect(fired).toHaveBeenCalledTimes(2);
        expect(fired).toHaveBeenCalledWith(undefined);
        listener.dispose();
    });

    it('ignores every URI outside its own scheme', () => {
        expect(provider.provideFileDecoration(vscode.Uri.file(WORKBOOK))).toBeUndefined();
    });
});

describe('git marks through the same provider', () => {
    const marks = new Map<string, { byModule: Map<string, 'modified' | 'added'>; removed: number; head: string }>();
    const emitter = new vscode.EventEmitter<string>();
    const source = {
        marksFor: (filePath: string) => marks.get(filePath),
        onDidChange: emitter.event,
    };
    const provider = new AgentReviewDecorationProvider(source);

    afterEach(() => {
        marks.clear();
        keepAgentChange(WORKBOOK, 'Ledger');
    });

    it('badges a modified module M and an added one A, in git colours', () => {
        marks.set(WORKBOOK, { byModule: new Map([['ledger', 'modified'], ['fresh', 'added']]), removed: 0, head: 'abc' });

        const modified = provider.provideFileDecoration(moduleDecorationUri(WORKBOOK, 'Ledger'));
        expect(modified?.badge).toBe('M');
        expect((modified?.color as unknown as { id: string }).id).toBe('gitDecoration.modifiedResourceForeground');
        const added = provider.provideFileDecoration(moduleDecorationUri(WORKBOOK, 'Fresh'));
        expect(added?.badge).toBe('A');
        expect((added?.color as unknown as { id: string }).id).toBe('gitDecoration.addedResourceForeground');
        expect(provider.provideFileDecoration(moduleDecorationUri(WORKBOOK, 'Same'))).toBeUndefined();
    });

    it('counts modified, added and removed modules on the project row', () => {
        marks.set(WORKBOOK, { byModule: new Map([['ledger', 'modified'], ['fresh', 'added']]), removed: 1, head: 'abc' });

        const decoration = provider.provideFileDecoration(projectDecorationUri(WORKBOOK));
        expect(decoration?.badge).toBe('3');
        expect(decoration?.tooltip).toBe('3 modules changed since the last commit');
    });

    it('says nothing for a project without marks', () => {
        expect(provider.provideFileDecoration(projectDecorationUri(WORKBOOK))).toBeUndefined();
        marks.set(WORKBOOK, { byModule: new Map(), removed: 0, head: 'abc' });
        expect(provider.provideFileDecoration(projectDecorationUri(WORKBOOK))).toBeUndefined();
    });

    it('lets an agent edit awaiting review outrank the git mark on the same row', async () => {
        marks.set(WORKBOOK, { byModule: new Map([['ledger', 'modified']]), removed: 0, head: 'abc' });
        await agentEdits(WORKBOOK, 'Ledger');

        expect(provider.provideFileDecoration(moduleDecorationUri(WORKBOOK, 'Ledger'))?.badge).toBe('AI');
        expect(provider.provideFileDecoration(projectDecorationUri(WORKBOOK))?.badge).toBe('1');

        keepAgentChange(WORKBOOK, 'Ledger');
        expect(provider.provideFileDecoration(moduleDecorationUri(WORKBOOK, 'Ledger'))?.badge).toBe('M');
    });

    it('asks for every row again when the marks change', () => {
        const fired = vi.fn();
        const listener = provider.onDidChangeFileDecorations(fired);
        emitter.fire(WORKBOOK);
        expect(fired).toHaveBeenCalledWith(undefined);
        listener.dispose();
    });
});
