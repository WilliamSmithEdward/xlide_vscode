import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { VbaTypeSemanticTokensProvider } from '../vbaSemanticTokensProvider';
import type { VbaProjectIndexService } from '../vbaProjectIndexService';
import { activate, closeAllEditors, open, until, workspaceRoot } from './support';

async function probe(name: string, source: string, marker: string): Promise<{ document: vscode.TextDocument; editor: vscode.TextEditor; caret: vscode.Position }> {
    const file = path.join(workspaceRoot(), `${name}.bas`);
    fs.writeFileSync(file, source);
    const document = await open(vscode.Uri.file(file));
    const editor = vscode.window.activeTextEditor!;
    const caret = document.positionAt(source.indexOf(marker) + marker.length);
    editor.selection = new vscode.Selection(caret, caret);
    await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
    return { document, editor, caret };
}
async function completions(document: vscode.TextDocument, caret: vscode.Position): Promise<vscode.CompletionList> {
    return (await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider', document.uri, caret))!;
}
function labels(list: vscode.CompletionList): string[] {
    return list.items.map(item => typeof item.label === 'string' ? item.label : item.label.label);
}

suite('Completion editor surface', () => {
    const previousSettings = new Map<string, unknown>();
    suiteSetup(async () => {
        await activate();
        for (const key of ['quickSuggestions', 'wordBasedSuggestions', 'suggestSelection']) {
            previousSettings.set(key, vscode.workspace.getConfiguration('editor').inspect(key)?.workspaceValue);
        }
        await vscode.workspace.getConfiguration('editor').update('quickSuggestions', false, vscode.ConfigurationTarget.Workspace);
        await vscode.workspace.getConfiguration('editor').update('wordBasedSuggestions', 'off', vscode.ConfigurationTarget.Workspace);
        await vscode.workspace.getConfiguration('editor').update('suggestSelection', 'first', vscode.ConfigurationTarget.Workspace);
    });
    suiteTeardown(async () => {
        for (const [key, value] of previousSettings) {
            await vscode.workspace.getConfiguration('editor').update(key, value, vscode.ConfigurationTarget.Workspace);
        }
    });
    teardown(async () => {
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
        await closeAllEditors();
    });
    test('invalidates semantic tokens across real document close/open language events', async () => {
        const provider = new VbaTypeSemanticTokensProvider({} as VbaProjectIndexService);
        const token = { isCancellationRequested: false } as vscode.CancellationToken;
        try {
            const { document } = await probe('SemanticReopen',
                'Sub Demo()\n    Debug.Print ThisWorkbook.Name\nEnd Sub\n', 'ThisWorkbook');
            const before = await provider.provideDocumentSemanticTokens(document, token);
            assert.ok(before.data.length, 'initial file should have semantic tokens');
            const version = document.version;
            const language = document.languageId;
            // Language changes emit close/open document events even when the
            // workbench retains the file model after a tab closes.
            const plain = await vscode.languages.setTextDocumentLanguage(document, 'plaintext');
            const reopened = await vscode.languages.setTextDocumentLanguage(plain, language);
            assert.equal(reopened.version, version, 'document lifecycle should preserve the same cache version');
            const after = await provider.provideDocumentSemanticTokens(reopened, token);
            assert.ok(after.data.length, 'reopened VBA document should have semantic tokens');
            assert.notEqual(after, before, 'closed document token cache should have been discarded');
            assert.deepEqual(Array.from(after.data), Array.from(before.data), 'unchanged source should retain its token positions');
        } finally { provider.dispose(); }
    });
    test('serves repeated completion requests in a large unchanged module', async () => {
        const source = 'Sub Demo()\nDim value As Long\n' + 'value = value + 1\n'.repeat(3000) +
            'ThisWorkbook.Sheets(1).ce\nEnd Sub\n';
        const { document, caret } = await probe('CompletionLarge', source, '.ce');
        const times: number[] = [];
        for (let i = 0; i < 10; i++) {
            const start = performance.now();
            assert.ok(labels(await completions(document, caret)).includes('Cells'));
            times.push(performance.now() - start);
        }
        console.log(`Completion large module ms: first=${times[0].toFixed(1)}, repeated=[${times.slice(1).map(time => time.toFixed(1)).join(',')}]`);
    });
    test('returns Cells for a corrected worksheet prefix', async () => {
        const { document, editor, caret } = await probe('CompletionPrefix', 'Sub Demo()\nThisWorkbook.Sheets(1).cez\nEnd Sub\n', '.cez');
        assert.deepEqual(labels(await completions(document, caret)), []);
        const start = caret.translate(0, -1);
        await editor.edit(edit => edit.delete(new vscode.Range(start, caret)));
        editor.selection = new vscode.Selection(start, start);
        assert.ok(labels(await completions(document, start)).includes('Cells'));
    });
    test('reopens the actual menu after Backspace, allowing Cells to be accepted', async () => {
        const { document } = await probe('CompletionMenu', 'Sub Demo()\nThisWorkbook.Sheets(1).cez\nEnd Sub\n', '.cez');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('deleteLeft');
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return document.lineAt(1).text.endsWith('.Cells') ? true : undefined;
        },
            `backspace should reopen Cells: ${document.lineAt(1).text}`, 4000);
    });
    test('offers procedures from the current module inside a macro-name string', async () => {
        const source = 'Public Sub Clicked()\nEnd Sub\nSub Demo()\nApplication.Run ""\nEnd Sub\n';
        const { document, caret } = await probe('CompletionMacro', source, 'Application.Run "');
        assert.ok(labels(await completions(document, caret)).includes('Module.Clicked'),
            'Application.Run should offer the current module procedure');
    });
    test('preserves both quotes when a macro suggestion replaces an empty string', async () => {
        const source = 'Public Sub Clicked()\nEnd Sub\nSub Demo()\nApplication.Run ""\nEnd Sub\n';
        const { document, caret } = await probe('CompletionMacroQuotes', source, 'Application.Run "');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Module.Clicked');
        assert.ok(item);
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        assert.equal(document.getText(range), '', 'the completion range must exclude the closing quote');
        const edit = new vscode.WorkspaceEdit();
        edit.replace(document.uri, range, 'Module.Clicked');
        await vscode.workspace.applyEdit(edit);
        assert.equal(document.lineAt(3).text, 'Application.Run "Module.Clicked"');
    });
    test('hovers a current-module procedure named in a macro string', async () => {
        const source = 'Public Sub Clicked()\nEnd Sub\nSub Demo()\nApplication.Run "Module.Clicked"\nEnd Sub\n';
        const { document, caret } = await probe('CompletionMacroHover', source, '"Module.Clicked');
        const hovers = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider', document.uri, caret.translate(0, -1));
        assert.ok(hovers?.some(hover => hover.contents.some(content =>
            (typeof content === 'string' ? content : content.value).includes('Sub Clicked'))),
            'hover should resolve the current module procedure named by Application.Run');
    });
    for (const [expression, name, prefix] of [
        ['value = Abs(-1)', 'Abs', 'value = Ab'],
        ['value = Left$("abc", 1)', 'Left$', 'value = Lef'],
        ['Set value = Application.Intersect(a, b)', 'Intersect', 'Application.Int'],
    ]) {
        test(`preserves existing arguments when completing ${name}`, async () => {
            const source = `Sub Demo()\n${expression}\nEnd Sub\n`;
            const { document, editor, caret } = await probe(`CompletionArguments${name}`, source, prefix);
            const item = (await completions(document, caret)).items.find(candidate => candidate.label === name);
            assert.ok(item);
            const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
            assert.ok(range);
            const insertText = item.insertText;
            if (insertText instanceof vscode.SnippetString) {
                await editor.insertSnippet(insertText, range);
            } else {
                await editor.edit(edit => edit.replace(range, insertText ?? name));
            }
            assert.equal(document.lineAt(1).text, expression, 'completion must retain the original argument list');
        });
    }
    test('completes and accepts a bracketed worksheet member in the actual menu', async () => {
        const source = 'Sub Demo()\nThisWorkbook.Sheets(1).[Ce]\nEnd Sub\n';
        const { document, caret } = await probe('CompletionBracketed', source, '.[Ce');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Cells');
        assert.ok(item, 'Cells must be offered inside a bracketed name');
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        assert.equal(document.getText(range), '[Ce]');
        await vscode.commands.executeCommand('editor.action.triggerSuggest');
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return document.lineAt(1).text.endsWith('.[Cells]') ? true : undefined;
        },
            'accepting Cells must replace the bracketed name once', 4000);
    });
    test('reopens the bracketed member menu after an unmatched prefix is corrected', async () => {
        const source = 'Sub Demo()\nThisWorkbook.Sheets(1).[Cez]\nEnd Sub\n';
        const { document } = await probe('CompletionBracketedRecovery', source, '.[Cez');
        await vscode.commands.executeCommand('hideSuggestWidget');
        await vscode.commands.executeCommand('deleteLeft');
        await until(async () => {
            await vscode.commands.executeCommand('acceptSelectedSuggestion');
            return document.lineAt(1).text.endsWith('.[Cells]') ? true : undefined;
        }, 'Backspace must reopen the bracketed member menu', 4000);
    });
    test('inserts Err as an object rather than an empty function call', async () => {
        const { document, editor, caret } = await probe('CompletionRuntimeObject', 'Sub Demo()\nSet obj = Er\nEnd Sub\n', 'Set obj = Er');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Err');
        assert.ok(item);
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        const insertText = item.insertText;
        if (insertText instanceof vscode.SnippetString) {
            await editor.insertSnippet(insertText, range);
        } else {
            await editor.edit(edit => edit.replace(range, insertText ?? 'Err'));
        }
        assert.equal(document.lineAt(1).text, 'Set obj = Err');
    });
    test('inserts parentheses for a bare host method in an expression', async () => {
        const { document, editor, caret } = await probe('CompletionGlobalMethod', 'Sub Demo()\nSet obj = Uni\nEnd Sub\n', 'Set obj = Uni');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === 'Union');
        assert.ok(item);
        assert.ok(item.insertText instanceof vscode.SnippetString);
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        await editor.insertSnippet(item.insertText, range);
        assert.equal(document.lineAt(1).text, 'Set obj = Union()');
    });
    test('keeps deleting through member prefixes with the smart Backspace command', async () => {
        const expression = 'ThisWorkbook.Sheets(1).az';
        const source = 'Public Property Get Demo() As Variant\nIf True Then\nEnd If\n' + expression + '\nEnd Property\n';
        const { document } = await probe('CompletionRepeatedBackspace', source, expression);
        for (let removed = 1; removed <= 10; removed++) {
            await vscode.commands.executeCommand('xlide.vba.smartBackspace');
            assert.equal(document.lineAt(3).text, expression.slice(0, -removed), `Backspace ${removed} must delete another character`);
        }
    });
    test('keeps smart Backspace working while joining trailing blank lines', async () => {
        const source = 'Sub Demo()\nEnd Sub\n' + '\n'.repeat(12);
        const { document } = await probe('CompletionBackspaceLineJoin', source, 'End Sub');
        await vscode.commands.executeCommand('cursorBottom');
        for (let removed = 1; removed <= 12; removed++) {
            await vscode.commands.executeCommand('xlide.vba.smartBackspace');
            assert.equal(document.getText(), source.slice(0, -removed), `Backspace ${removed} must join the next blank line`);
        }
    });
    test('measures typing and smart Backspace in a large module', async () => {
        const source = 'Sub Demo()\nDim value As Long\n' + 'value = value + 1\n'.repeat(3000) + 'value = 12345\nEnd Sub\n';
        const { document } = await probe('CompletionTypingWork', source, 'value = 12345');
        const times: number[] = [];
        for (const text of ['6', '7', '8', '9', '0']) {
            const start = performance.now();
            await vscode.commands.executeCommand('type', { text });
            times.push(performance.now() - start);
        }
        for (let removed = 0; removed < 5; removed++) {
            await vscode.commands.executeCommand('xlide.vba.smartBackspace');
        }
        assert.equal(document.lineAt(3002).text, 'value = 12345');
        console.log(`Typing large module ms: [${times.map(time => time.toFixed(1)).join(',')}]`);
    });
    test('does not offer code completions inside an ordinary string', async () => {
        const { document, caret } = await probe('CompletionString', 'Sub Demo()\nDebug.Print "hello"\nEnd Sub\n', 'hello');
        assert.equal((await completions(document, caret)).items.length, 0);
    });
    test('does not offer code completions in a comment', async () => {
        const { document, caret } = await probe('CompletionComment', "Sub Demo()\n' ordinary comment\nEnd Sub\n", 'comment');
        assert.equal((await completions(document, caret)).items.length, 0);
    });
    test('inserts a member containing combining marks as an ordinary identifier', async () => {
        const name = '\u0915\u093eValue';
        const source = `Public Type Record\n${name} As Long\nEnd Type\nSub Demo()\nDim obj As Record\nobj.\u0915\u093e\nEnd Sub\n`;
        const { document, caret } = await probe('CompletionCombiningMarks', source, 'obj.\u0915\u093e');
        const item = (await completions(document, caret)).items.find(candidate => candidate.label === name);
        assert.ok(item);
        assert.equal(item.insertText, name, 'a valid combining-mark identifier must not be bracketed');
    });
    test('replaces a Unicode UDT member prefix with the correct editor range', async () => {
        const source = 'Public Type Record\nCaf\u00e9Value As Long\nEnd Type\nSub Demo()\nDim obj As Record\nobj.caf\u00e9\nEnd Sub\n';
        const { document, caret } = await probe('CompletionUnicode', source, 'obj.caf\u00e9');
        const list = await completions(document, caret);
        const item = list.items.find(item => item.label === 'Caf\u00e9Value');
        assert.ok(item, 'UDT field should be offered');
        const range = item.range instanceof vscode.Range ? item.range : item.range?.replacing;
        assert.ok(range);
        assert.equal(document.getText(range), 'caf\u00e9');
    });
});
