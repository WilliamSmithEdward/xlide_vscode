import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { Session } from 'node:inspector';
import * as vscode from 'vscode';
import { activate, closeAllEditors, insertTypedCharacter, open, workspaceRoot } from './support';
import { readModule, writeModule } from '../vba/projectService';
import { encodeModuleUri } from '../xlideFileSystem';
import { backspaceNeedsExtension } from '../vbaEditorCommands';
import { runRendererBackspaceProbe } from './rendererBackspaceProbe';

(process.env.XLIDE_PERF_WORKBOOK ? suite : suite.skip)('Actual large class latency', () => {
    let document: vscode.TextDocument;
    let line: number;
    suiteSetup(async () => {
        await activate();
        const copy = path.join(workspaceRoot(), 'LargeClassLatency.xlsm');
        fs.copyFileSync(process.env.XLIDE_PERF_WORKBOOK!, copy);
        const original = readModule(copy, 'ROneCOne').source;
        const probe = '\r\nPublic Sub XlideLatencyProbe()\r\n    Dim LatencyValue As Long\r\n    Latency\r\n    ThisWorkbook.Sheets(1).\r\nEnd Sub\r\n';
        writeModule(copy, 'ROneCOne', original + probe, 'class');
        document = await open(encodeModuleUri(copy, 'ROneCOne'));
        line = document.lineCount - 4;
        console.log('Actual class fixture:', JSON.stringify({ characters: original.length, lines: document.lineCount }));
    });
    suiteTeardown(async () => {
        if (document?.isDirty) { await document.save(); }
        await closeAllEditors();
    });
    test('records actual edit-to-completion latency and context phases in a disposable workbook copy', async () => {
        const editor = vscode.window.activeTextEditor!;
        const samples: unknown[] = [];
        const profileEnabled = process.env.XLIDE_PERF_CPU_PROFILE === '1';
        const profiler = new Session();
        if (profileEnabled) { profiler.connect(); }
        const post = (method: string) => new Promise<any>((resolve, reject) => profiler.post(method, (error, value) => error ? reject(error) : resolve(value)));
        if (profileEnabled) { await post('Profiler.enable'); await post('Profiler.start'); }
        try {
            for (const character of 'value') {
                await new Promise(resolve => setTimeout(resolve, 30));
                const before = performance.now();
                const editAttempts = await insertTypedCharacter(document, editor, line, character);
                const afterEdit = performance.now();
                const result = await vscode.commands.executeCommand<vscode.CompletionList>(
                    'vscode.executeCompletionItemProvider', document.uri, document.lineAt(line).range.end);
                const afterResult = performance.now();
                assert.ok(result?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'LatencyValue'));
                samples.push({ character, editAttempts, editMs: afterEdit-before, completionMs: afterResult-afterEdit,
                    totalMs: afterResult-before });
            }
            console.log('Actual large class typing latency:', JSON.stringify(samples));
            if (profileEnabled) {
                const profile = await post('Profiler.stop');
                fs.writeFileSync(path.join(workspaceRoot(), 'large-class.cpuprofile'), JSON.stringify(profile.profile));
            }
            const clipboard = await vscode.env.clipboard.readText();
            try {
                await vscode.commands.executeCommand('xlide.copyPerformanceSnapshot');
                console.log('Actual extension traces:', await vscode.env.clipboard.readText());
            } finally { await vscode.env.clipboard.writeText(clipboard); }
        } finally {
            if (profileEnabled) { profiler.disconnect(); }
        }
    });

    test('records sustained typing and native/cleanup Backspace routes in the actual class', async () => {
        const editor = vscode.window.activeTextEditor!;
        // Probe both the original reported area and the appended procedure.
        const samples: { line: number; idleMs: number; typingMs: number; backspaceMs: number; backspaceCommand: string }[] = [];
        const profiler = new Session();
        const profileEnabled = process.env.XLIDE_PERF_CPU_PROFILE === '1';
        const post = (method: string) => new Promise<any>((resolve, reject) => profiler.post(method, (error, value) => error ? reject(error) : resolve(value)));
        if (profileEnabled) { profiler.connect(); await post('Profiler.enable'); await post('Profiler.start'); }
        let expected = performance.now() + 10;
        let maxHostDelayMs = 0;
        const heartbeat = setInterval(() => {
            const now = performance.now();
            maxHostDelayMs = Math.max(maxHostDelayMs, now - expected);
            expected = now + 10;
        }, 10);
        try {
            for (const probeLine of [Math.min(1500, document.lineCount - 1), line]) {
                const end = document.lineAt(probeLine).range.end;
                editor.selection = new vscode.Selection(end, end);
                await new Promise(resolve => setTimeout(resolve, 300));
                await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
                const before = document.getText();
                for (let i = 0; i < 24; i++) {
                    const idleMs = [0, 30, 250, 350][i % 4];
                    await new Promise(resolve => setTimeout(resolve, idleMs));
                    await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
                    const originalLine = document.lineAt(probeLine).text;
                    const start = performance.now();
                    await vscode.commands.executeCommand('type', { text: 'z' });
                    const typed = performance.now();
                    assert.ok(document.lineAt(probeLine).text === originalLine + 'z',
                        'Keyboard typing must insert before measuring Backspace: ' + JSON.stringify({
                            probeLine, idleMs, beforeLength: originalLine.length,
                            afterLength: document.lineAt(probeLine).text.length, caret: editor.selection.active }));
                    const backspaceCommand = backspaceNeedsExtension(editor) ? 'xlide.vba.smartBackspace' : 'deleteLeft';
                    await vscode.commands.executeCommand(backspaceCommand);
                    samples.push({ line: probeLine, idleMs, typingMs: typed - start, backspaceMs: performance.now() - typed, backspaceCommand });
                }
                const after = document.getText();
                if (after !== before) {
                    let first = 0;
                    while (first < Math.min(before.length, after.length) && before[first] === after[first]) { first++; }
                    console.log('Typing restoration difference:', JSON.stringify({ probeLine,
                        beforeLength: before.length, afterLength: after.length, first: document.positionAt(first),
                        caseOnly: before.toLowerCase() === after.toLowerCase(), caret: editor.selection.active }));
                }
                assert.ok(after === before, 'Typing and Backspace must restore the document');
            }
            console.log('Actual sustained typing latency:', JSON.stringify({ samples, maxHostDelayMs }));
        } finally {
            clearInterval(heartbeat);
            if (profileEnabled) {
                const stopped = await post('Profiler.stop');
                fs.writeFileSync(path.join(workspaceRoot(), 'sustained-typing.cpuprofile'), JSON.stringify(stopped.profile));
                profiler.disconnect();
            }
        }
    });

    test('keeps member completion and hover working across repeated missed prefixes and Backspace', async () => {
        const editor = vscode.window.activeTextEditor!;
        const memberLine = line + 1;
        const original = document.lineAt(memberLine).text;
        assert.ok(original.endsWith('ThisWorkbook.Sheets(1).'));
        editor.selection = new vscode.Selection(document.lineAt(memberLine).range.end, document.lineAt(memberLine).range.end);
        await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
        const samples: { backspaceMs: number; completionMs: number; hoverMs: number }[] = [];
        for (let i = 0; i < 8; i++) {
            await vscode.commands.executeCommand('type', { text: 'cez' });
            assert.ok(document.lineAt(memberLine).text.endsWith('.cez'), 'Keyboard typing must reach the member prefix');
            assert.equal(backspaceNeedsExtension(editor), false);
            const start = performance.now();
            await vscode.commands.executeCommand('deleteLeft');
            const deleted = performance.now();
            assert.ok(document.lineAt(memberLine).text.endsWith('.ce'));
            const result = await vscode.commands.executeCommand<vscode.CompletionList>(
                'vscode.executeCompletionItemProvider', document.uri, document.lineAt(memberLine).range.end);
            const completed = performance.now();
            assert.ok(result?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Cells'));
            const offset = document.getText().toLowerCase().lastIndexOf('latencyvalue') + 2;
            const hovers = await vscode.commands.executeCommand<vscode.Hover[]>(
                'vscode.executeHoverProvider', document.uri, document.positionAt(offset));
            assert.ok(hovers?.length, 'Hover must still respond after Backspace');
            samples.push({ backspaceMs: deleted - start, completionMs: completed - deleted, hoverMs: performance.now() - completed });
            await vscode.commands.executeCommand('hideSuggestWidget');
            await vscode.commands.executeCommand('deleteLeft');
            await vscode.commands.executeCommand('deleteLeft');
            assert.ok(document.lineAt(memberLine).text === original, 'Repeated Backspace must restore the member line');
        }
        console.log('Actual member recovery latency:', JSON.stringify(samples));
    });

    test('records visible typing, Backspace and menu recovery in the actual class', async function () {
        if (!process.env.XLIDE_UI_DEBUG_PORT) { this.skip(); }
        const editor = vscode.window.activeTextEditor!;
        const memberLine = line + 1;
        const original = document.lineAt(memberLine).text;
        const caret = document.lineAt(memberLine).range.end;
        editor.selection = new vscode.Selection(caret, caret);
        editor.revealRange(new vscode.Range(caret, caret), vscode.TextEditorRevealType.InCenter);
        await vscode.commands.executeCommand('workbench.action.focusActiveEditorGroup');
        await vscode.commands.executeCommand('type', { text: 'cez' });
        await vscode.commands.executeCommand('hideSuggestWidget');
        assert.ok(document.lineAt(memberLine).text.endsWith('.cez'));
        assert.equal(backspaceNeedsExtension(editor), false);
        const result = await runRendererBackspaceProbe('stress');
        assert.equal(result.samples?.length, 24);
        assert.ok(document.lineAt(memberLine).text.endsWith('.cez'));
        console.log('Actual renderer typing latency:', JSON.stringify(result));
        for (let i = 0; i < 3; i++) { await vscode.commands.executeCommand('deleteLeft'); }
        assert.equal(document.lineAt(memberLine).text, original);
    });

    test('records hover after a declaration edit in the actual large class', async () => {
        const hoverText = async () => {
            const source = document.getText();
            const offset = source.toLowerCase().lastIndexOf('latencyvalue') + 2;
            const hovers = await vscode.commands.executeCommand<vscode.Hover[]>(
                'vscode.executeHoverProvider', document.uri, document.positionAt(offset));
            return (hovers ?? []).flatMap(item => item.contents.map(content =>
                typeof content === 'string' ? content : content.value)).join('\n');
        };
        const beforeWarm = performance.now();
        assert.match(await hoverText(), /LatencyValue As Long/);
        const warmMs = performance.now() - beforeWarm;
        const beforeEdit = performance.now();
        let edited = false;
        let editAttempts = 0;
        // Automatic casing can update the document while VS Code applies the
        // test edit. A rejected edit changed nothing; retry with fresh positions
        // and include all attempts in the measured edit-to-hover duration.
        for (; editAttempts < 3 && !edited; editAttempts++) {
            const source = document.getText();
            const declaration = source.lastIndexOf('Dim LatencyValue As Long');
            assert.notEqual(declaration, -1);
            const start = declaration + 'Dim LatencyValue As '.length;
            edited = await vscode.window.activeTextEditor!.edit(edit => edit.replace(
                new vscode.Range(document.positionAt(start), document.positionAt(start + 4)), 'Double'));
        }
        assert.ok(edited);
        assert.match(await hoverText(), /LatencyValue As Double/);
        console.log('Actual large class hover latency:', JSON.stringify({ warmMs, editAttempts, editToHoverMs: performance.now() - beforeEdit }));
    });
    test('records code actions after typing at a large-class caret', async () => {
        const actionsAtCaret = async () => {
            const caret = document.lineAt(line).range.end;
            const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>(
                'vscode.executeCodeActionProvider', document.uri, new vscode.Range(caret, caret));
            assert.ok(actions?.some(action => action.command?.command === 'xlide.refactor.introduceParameter'));
        };
        const warmStart = performance.now();
        await actionsAtCaret();
        const warmMs = performance.now() - warmStart;
        const editStart = performance.now();
        let edited = false, editAttempts = 0;
        for (; editAttempts < 3 && !edited; editAttempts++) {
            const caret = document.lineAt(line).range.end;
            edited = await vscode.window.activeTextEditor!.edit(edit => edit.insert(caret, ' '));
        }
        assert.ok(edited);
        await actionsAtCaret();
        console.log('Actual large class code action latency:', JSON.stringify({ warmMs, editAttempts, editToActionsMs: performance.now() - editStart }));
    });

});
