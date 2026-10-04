// Which modules a tab closure leaves with nothing open.
//
// The tree follows the editor into a module and expands it, and only a tab
// closing takes that back: focus moves to the Output panel, a terminal or a
// webview while the module is still open, so focus loss cannot be the signal.
// What matters here is the "still open" half - a module shown in another tab
// group, or in one side of a diff, is being edited whatever just closed.

import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', async () => (await import('./helpers/vscodeMock')).vscodeMock());

import * as vscode from 'vscode';
import { moduleLocationOfUri, modulesWithNoTabLeft, tabUris } from '../src/vbaDocumentLocation';
import { encodeFormMarkupUri, encodeModuleUri } from '../src/xlideFileSystem';

const BOOK = 'C:\\work\\Book.xlsm';
const OTHER = 'C:\\work\\Other.xlsm';

/** A tab showing one document, the way VS Code hands one over. */
const textTab = (uri: vscode.Uri): vscode.Tab =>
    ({ input: new vscode.TabInputText(uri) }) as unknown as vscode.Tab;
const diffTab = (left: vscode.Uri, right: vscode.Uri): vscode.Tab =>
    ({ input: new vscode.TabInputTextDiff(left, right) }) as unknown as vscode.Tab;
const customTab = (uri: vscode.Uri): vscode.Tab =>
    ({ input: new vscode.TabInputCustom(uri, 'xlide.formDesigner') }) as unknown as vscode.Tab;
/** A tab showing no document at all: the settings editor, a webview, a terminal. */
const otherTab = (): vscode.Tab => ({ input: undefined }) as unknown as vscode.Tab;

const named = (locations: ReturnType<typeof modulesWithNoTabLeft>): string[] =>
    locations.map((one) => one.moduleName);

describe('the documents a tab shows', () => {
    it('reads a text tab, a custom editor, and both sides of a diff', () => {
        const left = encodeModuleUri(BOOK, 'Module1');
        const right = encodeModuleUri(BOOK, 'Module2');

        expect(tabUris(textTab(left)).map(String)).toEqual([left.toString()]);
        expect(tabUris(customTab(left)).map(String)).toEqual([left.toString()]);
        expect(tabUris(diffTab(left, right)).map(String))
            .toEqual([left.toString(), right.toString()]);
    });

    it('answers with nothing for a tab that shows no document', () => {
        expect(tabUris(otherTab())).toEqual([]);
    });
});

describe('the modules a closure leaves with nothing open', () => {
    it('names the module whose last tab closed', () => {
        const uri = encodeModuleUri(BOOK, 'Module1');

        const [location] = modulesWithNoTabLeft([textTab(uri)], []);

        expect(location.moduleName).toBe('Module1');
        // The project is whatever the locator reads out of that same address.
        // Its spelling is the URI's own, which is not the literal above on
        // every platform, so the locator is the reference rather than BOOK.
        expect(location).toEqual(moduleLocationOfUri(uri));
    });

    it('leaves alone a module still open in another tab group', () => {
        const uri = encodeModuleUri(BOOK, 'Module1');

        expect(modulesWithNoTabLeft([textTab(uri)], [textTab(uri)])).toEqual([]);
    });

    it('leaves alone a module still open as one side of a diff', () => {
        // Comparing a module with its committed copy is an open editor of it.
        const uri = encodeModuleUri(BOOK, 'Module1');
        const head = encodeModuleUri(BOOK, 'Module2');

        expect(modulesWithNoTabLeft([textTab(uri)], [diffTab(uri, head)])).toEqual([]);
    });

    it('names a module whose diff tab closed and has nothing left', () => {
        const uri = encodeModuleUri(BOOK, 'Module1');
        const head = encodeModuleUri(BOOK, 'Module1');

        expect(named(modulesWithNoTabLeft([diffTab(uri, head)], []))).toEqual(['Module1']);
    });

    it('says nothing about a tab that was not a module', () => {
        expect(modulesWithNoTabLeft([textTab(vscode.Uri.file('C:\\work\\notes.txt'))], [])).toEqual([]);
        expect(modulesWithNoTabLeft([otherTab()], [])).toEqual([]);
    });

    it('names a module once when a form closes every face of it at once', () => {
        // A form's code, its markup and its designer are three tabs and one
        // module, and the tree folds that module once.
        const closed = [
            textTab(encodeModuleUri(BOOK, 'FrmMain')),
            textTab(encodeFormMarkupUri(BOOK, 'FrmMain')),
            customTab(encodeFormMarkupUri(BOOK, 'FrmMain')),
        ];

        expect(named(modulesWithNoTabLeft(closed, []))).toEqual(['FrmMain']);
    });

    it('names each module when several close together', () => {
        const closed = [
            textTab(encodeModuleUri(BOOK, 'Module1')),
            textTab(encodeModuleUri(OTHER, 'Module1')),
            textTab(encodeModuleUri(BOOK, 'Module2')),
        ];

        const found = modulesWithNoTabLeft(closed, []);

        expect(named(found)).toEqual(['Module1', 'Module1', 'Module2']);
        // The two Module1 rows are different modules because their projects
        // are, which is the part a shared name could hide.
        expect(found[0].projectPath).not.toBe(found[1].projectPath);
        expect(found[0].projectPath).toBe(found[2].projectPath);
    });

    it('keys a module by identity, so its name spelled two ways is one module', () => {
        // VBA module names are case-insensitive and the tree keys them that
        // way. The project half of the key follows the platform's own file
        // rules - lowercased on Windows, not on Linux, where CI runs - so
        // only the name is compared here.
        const closed = [
            textTab(encodeModuleUri(BOOK, 'Module1')),
            textTab(encodeModuleUri(BOOK, 'MODULE1')),
        ];

        expect(named(modulesWithNoTabLeft(closed, []))).toEqual(['Module1']);
    });
});


describe('unrelated tab closures', () => {
    it.each(['text', 'custom', 'other'])('does not inspect remaining tab inputs for a non-module %s closure', kind => {
        let reads = 0;
        const open = Array.from({ length: 1000 }, () => ({
            get input() { reads++; return new vscode.TabInputText(encodeModuleUri(BOOK, 'Module1')); },
        }) as unknown as vscode.Tab);
        const closed = kind === 'text' ? textTab(vscode.Uri.file('C:/work/notes.txt'))
            : kind === 'custom' ? customTab(vscode.Uri.file('C:/work/notes.txt')) : otherTab();
        expect(modulesWithNoTabLeft([closed], open)).toEqual([]);
        expect(reads).toBe(0);
    });
});


it('does not enumerate lazy tab groups for empty or unrelated closures', () => {
    const open = vi.fn(() => [textTab(encodeModuleUri(BOOK, 'Module1'))]);
    expect(modulesWithNoTabLeft([], open)).toEqual([]);
    expect(modulesWithNoTabLeft([textTab(vscode.Uri.file('C:/work/notes.txt'))], open)).toEqual([]);
    expect(modulesWithNoTabLeft([otherTab()], open)).toEqual([]);
    expect(open).not.toHaveBeenCalled();
});

it('enumerates lazy tab groups once for mixed and multiple module closures', () => {
    const one = encodeModuleUri(BOOK, 'Module1'), two = encodeModuleUri(BOOK, 'Module2');
    const closed = [otherTab(), textTab(one), diffTab(two, one)];
    const remaining = [diffTab(one, vscode.Uri.file('C:/work/notes.txt'))];
    const open = vi.fn(() => remaining);
    expect(modulesWithNoTabLeft(closed, open)).toEqual(modulesWithNoTabLeft(closed, remaining));
    expect(named(modulesWithNoTabLeft(closed, remaining))).toEqual(['Module2']);
    expect(open).toHaveBeenCalledTimes(1);
});
