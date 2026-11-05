import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { VBA_SMART_BLOCK_SNIPPETS } from '../vbaSmartBlockSnippets';
import { activate, closeAllEditors, open, until, writeModule } from './support';

suite('Smart Enter surfaces', () => {
    suiteSetup(activate);
    teardown(async () => {
        for (const document of vscode.workspace.textDocuments) {
            if (document.isDirty && document.uri.scheme === 'xlide-vba') { await document.save(); }
        }
        await closeAllEditors();
    });

    async function enter(moduleName: string, source: string, line: number, indentation: string) {
        const document = await open(await writeModule(moduleName, source));
        const editor = vscode.window.activeTextEditor!;
        const column = document.lineAt(line).text.length;
        editor.selection = new vscode.Selection(line, column, line, column);
        const start = performance.now();
        assert.ok(await editor.edit(edit => edit.insert(new vscode.Position(line, column), '\r\n' + indentation),
            { undoStopBefore: false, undoStopAfter: false }));
        console.log(moduleName, 'Enter edit command (ms):', performance.now() - start);
        return document;
    }

    test('ordinary Enter in a large module leaves the new body line unchanged', async () => {
        const procedures = Array.from({ length: 1200 }, (_, index) =>
            'Sub Probe' + index + '()\r\n    Debug.Print "probe"\r\nEnd Sub\r\n').join('\r\n');
        const source = procedures + '\r\nSub ActiveProbe()\r\n    value = 1\r\nEnd Sub\r\n';
        const line = source.split('\r\n').length - 3;
        const document = await enter('EnterLarge', source, line, '    ');
        await new Promise(resolve => setTimeout(resolve, 800));
        assert.equal(document.lineAt(line + 1).text, '    ');
        assert.equal(document.lineAt(line + 2).text, 'End Sub');
    });

    test('comment Enter preserves indentation and apostrophe spacing', async () => {
        const document = await enter('EnterComment', "Sub T()\r\n    '''  note\r\nEnd Sub\r\n", 1, '    ');
        await until(() => document.lineAt(2).text === "    '''  " ? true : undefined,
            'comment continuation should preserve the prefix');
    });

    test('With member Enter seeds a dot inside the open block', async () => {
        const document = await enter('EnterWith',
            'Sub T()\r\n    With ActiveSheet\r\n        .Name = "probe"\r\n    End With\r\nEnd Sub\r\n', 2, '        ');
        await until(() => document.lineAt(3).text === '        .' ? true : undefined,
            'With continuation should seed a dot');
        const completions = await vscode.commands.executeCommand<vscode.CompletionList>(
            'vscode.executeCompletionItemProvider', document.uri, new vscode.Position(3, 9), '.');
        assert.ok(completions?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Name'));
    });

    // Issue #1620. Excel rejects a lone `.` (oracle bare_leading_member_access_inside_with_compile),
    // so the dot XLIDE seeds goes when its line is left with nothing after it.
    /** Opens a module and presses Enter, typed, at the end of `line`. */
    async function typeEnter(moduleName: string, source: string, line: number) {
        const document = await open(await writeModule(moduleName, source));
        const editor = vscode.window.activeTextEditor!;
        const column = document.lineAt(line).text.length;
        editor.selection = new vscode.Selection(line, column, line, column);
        await vscode.commands.executeCommand('type', { text: '\n' });
        return document;
    }

    test('removes the dot Smart Enter seeds after With once the caret leaves it unused', async () => {
        const document = await typeEnter('EnterWithLeave', 'Sub T()\r\n    With ActiveSheet\r\nEnd Sub\r\n', 1);
        const dotLine = (): number => document.getText().split('\r\n').findIndex(line => /^\s*\.$/.test(line));
        await until(() => dotLine() > 0 ? true : undefined, 'Smart Enter should seed a dot');
        const line = dotLine();
        // A real Enter key leaves the caret on the dot; the type command does not.
        const editor = vscode.window.activeTextEditor!;
        const caret = new vscode.Position(line, document.lineAt(line).text.length);
        editor.selection = new vscode.Selection(caret, caret);
        await new Promise(resolve => setTimeout(resolve, 100));
        assert.ok(/^\s*\.$/.test(document.lineAt(line).text), 'placing the caret keeps the dot');
        await vscode.commands.executeCommand('cursorUp');
        await until(() => document.lineAt(line).text.trim() === '' ? true : undefined, 'the unused dot should go');
        assert.match(document.getText(), /End With/);
    });

    test('keeps a dot with a member typed after it, and removes a lone one on Enter', async () => {
        const document = await typeEnter('EnterWithMember', 'Sub T()\r\n    With ActiveSheet\r\nEnd Sub\r\n', 1);
        const editor = vscode.window.activeTextEditor!;
        const dotLine = (): number => document.getText().split('\r\n').findIndex(line => /^\s*\.$/.test(line));
        await until(() => dotLine() > 0 ? true : undefined, 'Smart Enter should seed a dot');
        const line = dotLine();
        const caret = new vscode.Position(line, document.lineAt(line).text.length);
        editor.selection = new vscode.Selection(caret, caret);
        await vscode.commands.executeCommand('type', { text: 'Name = "a"' });
        await vscode.commands.executeCommand('type', { text: '\n' });
        await until(() => /^\s*\.$/.test(document.lineAt(line + 1).text) ? true : undefined, 'a member line seeds the next dot');
        assert.equal(document.lineAt(line).text.trim(), '.Name = "a"');
        await vscode.commands.executeCommand('type', { text: '\n' });
        await until(() => document.lineAt(line + 1).text.trim() === '' ? true : undefined, 'Enter on the lone dot removes it');
        assert.ok(!/^\s*\.$/.test(document.lineAt(line + 2).text), 'and seeds no other');
    });

    test('removes the With snippet\'s dot only after its line is visited and left', async () => {
        const document = await open(await writeModule('SnippetWith', 'Sub T()\r\n    \r\nEnd Sub\r\n'));
        const editor = vscode.window.activeTextEditor!;
        editor.selection = new vscode.Selection(1, 4, 1, 4);
        const snippet = VBA_SMART_BLOCK_SNIPPETS.find(spec => spec.label === 'With')!;
        assert.ok(await editor.insertSnippet(new vscode.SnippetString(snippet.insertText)));
        await vscode.commands.executeCommand('type', { text: 'ws' });
        const dotLine = (): number => document.getText().split('\r\n').findIndex(line => /^\s*\.$/.test(line));
        await new Promise(resolve => setTimeout(resolve, 200));
        assert.ok(dotLine() > 0, 'typing the object keeps the dot');
        await vscode.commands.executeCommand('jumpToNextSnippetPlaceholder');
        const line = dotLine();
        await until(() => editor.selection.active.line === line ? true : undefined, 'caret on the dot');
        await vscode.commands.executeCommand('cursorDown');
        await until(() => dotLine() < 0 ? true : undefined, 'the unused snippet dot should go');
        assert.match(document.lineAt(1).text, /With ws/);
    });

    test('block Enter inserts the missing closer', async () => {
        const document = await enter('EnterBlock', 'Sub T()\r\n    If ready Then\r\nEnd Sub\r\n', 1, '    ');
        await until(() => /End If/i.test(document.getText()) ? true : undefined,
            'If continuation should insert End If');
        assert.equal((document.getText().match(/End If/gi) ?? []).length, 1);
    });
});
