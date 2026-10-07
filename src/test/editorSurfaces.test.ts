import * as assert from 'node:assert/strict';
import * as vscode from 'vscode';
import { VbaCaretProcedureTracker } from '../vbaCaretProcedure';
import { writeModule as writeProjectModule } from '../vba/projectService';
import { activate, closeAllEditors, moduleUri, open, until, workbookPath, writeModule } from './support';

suite('Editor surfaces', () => {
    let document: vscode.TextDocument;
    const source = [
        'Option Explicit', 'Public Counter As Long', '',
        'Sub SurfaceProbe()', '    ThisWorkbook.Sheets(1).Name = "probe"',
        '    Counter = Counter + 1', '    Debug.Print Left("probe", 2)',
        'End Sub', '',
    ].join('\r\n');

    suiteSetup(async () => {
        await activate();
        await writeModule('SurfaceOther', 'Sub OtherProbe()\r\n    Debug.Print SurfaceProbeModule.Counter\r\nEnd Sub\r\n');
        document = await open(await writeModule('SurfaceProbeModule', source));
    });
    suiteTeardown(async () => {
        if (document?.isDirty) { await document.save(); }
        await closeAllEditors();
    });

    test('dot completion and hover survive an actual document edit', async () => {
        const position = new vscode.Position(4, '    ThisWorkbook.Sheets(1).'.length);
        const complete = () => vscode.commands.executeCommand<vscode.CompletionList>(
            'vscode.executeCompletionItemProvider', document.uri, position, '.');
        const first = await until(async () => {
            const result = await complete();
            return result?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Name') ? result : undefined;
        }, 'sheet member completion should contain Name');
        assert.ok(first.items.length > 0);
        const editor = vscode.window.activeTextEditor!;
        assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(4, position.character, 4, document.lineAt(4).text.length), 'Name = "edited"')));
        const second = await complete();
        assert.ok(second?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Name'));
        const hover = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider',
            document.uri, new vscode.Position(5, 6));
        assert.ok(hover?.length, 'Counter hover should remain available after editing');
        await document.save();
    });

    test('signature help and document-local references retain their results', async () => {
        const signature = await vscode.commands.executeCommand<vscode.SignatureHelp>(
            'vscode.executeSignatureHelpProvider', document.uri, new vscode.Position(6, '    Debug.Print Left("probe", '.length));
        assert.ok(signature?.signatures.some(item => /Left/i.test(item.label)));
        const highlights = await vscode.commands.executeCommand<vscode.DocumentHighlight[]>(
            'vscode.executeDocumentHighlights', document.uri, new vscode.Position(5, 6));
        assert.ok(highlights && highlights.length >= 3, 'declaration, write and read should be highlighted');
        assert.ok(highlights.some(item => item.kind === vscode.DocumentHighlightKind.Write));
        assert.ok(highlights.every(item => item.range.start.line === 1 || item.range.start.line === 5));
    });

    test('caret tracker follows body edits and procedure renames through real events', async () => {
        const editor = vscode.window.activeTextEditor!;
        editor.selection = new vscode.Selection(5, 4, 5, 4);
        const tracker = new VbaCaretProcedureTracker();
        try {
            assert.equal(tracker.current?.label, 'Sub SurfaceProbe');
            assert.ok(await editor.edit(edit => edit.insert(new vscode.Position(5, document.lineAt(5).text.length), ' + 2')));
            await until(() => tracker.current?.label === 'Sub SurfaceProbe' ? true : undefined, 'body edit should retain procedure');
            assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(3, 4, 3, 16), 'RenamedProbe')));
            await until(() => tracker.current?.label === 'Sub RenamedProbe' ? true : undefined, 'header rename should invalidate procedure ranges');
        } finally {
            tracker.dispose();
            await document.save();
        }
    });

    test('records warm provider command timings in the real extension host', async () => {
        const samples: Record<string, number[]> = {};
        const measure = async (name: string, invoke: () => Thenable<unknown>) => {
            await invoke();
            const times: number[] = [];
            for (let sample = 0; sample < 11; sample++) {
                const start = performance.now();
                await invoke();
                times.push(performance.now() - start);
            }
            samples[name] = times.sort((a, b) => a - b);
        };
        await measure('completion', () => vscode.commands.executeCommand('vscode.executeCompletionItemProvider', document.uri, new vscode.Position(4, 26), '.'));
        await measure('hover', () => vscode.commands.executeCommand('vscode.executeHoverProvider', document.uri, new vscode.Position(5, 6)));
        await measure('highlights', () => vscode.commands.executeCommand('vscode.executeDocumentHighlights', document.uri, new vscode.Position(5, 6)));
        console.log('Editor surface command medians (ms):', JSON.stringify(Object.fromEntries(Object.entries(samples).map(([key, times]) => [key, times[5]]))));
    });

    test('large-module typing keeps completion, hover and semantic tokens available', async () => {
        const largeSource = source + Array.from({ length: 1200 }, (_, index) =>
            '\r\nSub Extra' + index + '()\r\n    Debug.Print ' + index + '\r\nEnd Sub\r\n').join('');
        const large = await open(await writeModule('SurfaceLarge', largeSource));
        const editor = vscode.window.activeTextEditor!;
        const edits: number[] = [];
        try {
            for (let index = 0; index < 12; index++) {
                const column = large.lineAt(5).text.length;
                const start = performance.now();
                assert.ok(await editor.edit(edit => edit.insert(new vscode.Position(5, column), ' '),
                    { undoStopBefore: false, undoStopAfter: false }));
                editor.selection = new vscode.Selection(5, column + 1, 5, column + 1);
                edits.push(performance.now() - start);
            }
            const completion = await vscode.commands.executeCommand<vscode.CompletionList>(
                'vscode.executeCompletionItemProvider', large.uri, new vscode.Position(4, 26), '.');
            assert.ok(completion?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Name'));
            const hover = await vscode.commands.executeCommand<vscode.Hover[]>(
                'vscode.executeHoverProvider', large.uri, new vscode.Position(5, 6));
            assert.ok(hover?.length);
            const tokens = await until(async () => {
                const result = await vscode.commands.executeCommand<vscode.SemanticTokens>(
                    'vscode.provideDocumentSemanticTokens', large.uri);
                return result?.data.length ? result : undefined;
            }, 'large-module semantic tokens should be available');
            assert.ok(tokens.data.length > 0);
            edits.sort((a, b) => a - b);
            console.log('Large-module edit command median (ms):', edits[6], 'procedures:', 1201);
        } finally {
            if (large.isDirty) { await large.save(); }
        }
    });

    test('non-symbol hover positions remain empty, including a continued comment', async () => {
        const source = [
            'Option Explicit', 'Public Counter As Long', 'Sub T()',
            "    ' Counter _", '    Counter', '    Debug.Print 42',
            '    Debug.Print #1/1/2026#', 'End Sub', '',
        ].join('\r\n');
        const probe = await open(await writeModule('SurfaceEmptyHover', source));
        for (const [line, column] of [[3, 8], [4, 8], [5, 17], [6, 20], [5, 2]]) {
            const hover = await vscode.commands.executeCommand<vscode.Hover[]>(
                'vscode.executeHoverProvider', probe.uri, new vscode.Position(line, column));
            assert.equal(hover?.length ?? 0, 0, 'non-symbol position ' + line + ':' + column);
        }
        const symbol = await vscode.commands.executeCommand<vscode.Hover[]>(
            'vscode.executeHoverProvider', probe.uri, new vscode.Position(1, 9));
        assert.ok(symbol?.length, 'a real symbol still has hover information');
    });

    test('fresh large-class local highlights retain declaration, read and write ranges', async () => {
        const padding = Array.from({ length: 700 }, (_, index) => [
            `Public Function Padding${index}() As Object`,
            `    Set Padding${index} = Nothing`, 'End Function', '',
        ].join('\r\n')).join('');
        const probe = [
            'Public Sub LocalReferenceProbe(ByVal InputValue As Long)',
            '    Dim Target As Long', "    Target = InputValue ' revision 00",
            '    Debug.Print Target', '    ThisWorkbook.Sheets(1).cez', 'End Sub', '',
        ].join('\r\n');
        const moduleName = 'SurfaceLocalReferenceClass';
        writeProjectModule(workbookPath(), moduleName, padding + probe, 'class');
        const large = await open(moduleUri(moduleName));
        const editor = vscode.window.activeTextEditor!;
        const probeLine = large.positionAt(large.getText().indexOf('Public Sub LocalReferenceProbe(')).line;
        const expected = [
            [probeLine + 1, 8, 14, vscode.DocumentHighlightKind.Write],
            [probeLine + 2, 4, 10, vscode.DocumentHighlightKind.Write],
            [probeLine + 3, 16, 22, vscode.DocumentHighlightKind.Read],
        ];
        try {
            for (let revision = 1; revision <= 6; revision++) {
                const line = probeLine + 2;
                const end = large.lineAt(line).text.length;
                assert.ok(await editor.edit(edit => edit.replace(
                    new vscode.Range(line, end - 2, line, end), String(revision).padStart(2, '0'))));
                const highlights = await vscode.commands.executeCommand<vscode.DocumentHighlight[]>(
                    'vscode.executeDocumentHighlights', large.uri, new vscode.Position(line, 5));
                assert.deepEqual(highlights?.map(item => [
                    item.range.start.line, item.range.start.character, item.range.end.character, item.kind,
                ]), expected, `local highlight ranges on fresh revision ${revision}`);
                const miss = await vscode.commands.executeCommand<vscode.DocumentHighlight[]>(
                    'vscode.executeDocumentHighlights', large.uri, new vscode.Position(probeLine + 4, 27));
                // An undefined semantic result lets VS Code's generic word
                // provider return Text highlights. It must never claim reads/writes.
                assert.ok(miss?.every(item => item.kind === vscode.DocumentHighlightKind.Text) ?? true,
                    'an unmatched member has no semantic read/write highlights');
            }
        } finally {
            if (large.isDirty) { await large.save(); }
        }
    });

    test('macro-name strings still hover the cross-module procedure', async () => {
        await writeModule('SurfaceMacroTarget', 'Public Sub Run()\r\nEnd Sub\r\n');
        const source = 'Sub MacroProbe()\r\n    Application.Run "SurfaceMacroTarget.Run"\r\nEnd Sub\r\n';
        const probe = await open(await writeModule('SurfaceMacroCaller', source));
        await until(async () => {
            const hover = await vscode.commands.executeCommand<vscode.Hover[]>(
                'vscode.executeHoverProvider', probe.uri, new vscode.Position(1, source.split('\r\n')[1].indexOf('SurfaceMacroTarget') + 3));
            return hover?.some(item => item.contents.some(content =>
                typeof content === 'string' ? /Sub Run/.test(content) : /Sub Run/.test(content.value))) ? true : undefined;
        }, 'a macro-name string should hover the target procedure in another module');
    });
});
