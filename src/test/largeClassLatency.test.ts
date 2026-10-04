import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { Session } from 'node:inspector';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, workspaceRoot } from './support';
import { readModule, writeModule } from '../vba/projectService';
import { encodeModuleUri } from '../xlideFileSystem';

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
                const end = document.lineAt(line).range.end;
                editor.selection = new vscode.Selection(end, end);
                const before = performance.now();
                assert.ok(await editor.edit(edit => edit.insert(end, character)));
                const afterEdit = performance.now();
                const result = await vscode.commands.executeCommand<vscode.CompletionList>(
                    'vscode.executeCompletionItemProvider', document.uri, document.lineAt(line).range.end);
                const afterResult = performance.now();
                assert.ok(result?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'LatencyValue'));
                samples.push({ character, editMs: afterEdit-before, completionMs: afterResult-afterEdit,
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
});
