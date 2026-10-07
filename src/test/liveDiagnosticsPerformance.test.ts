import * as assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { activate, closeAllEditors, open, until, workspaceRoot } from './support';
import { encodeModuleUri } from '../xlideFileSystem';
import { readModules } from '../vba/projectService';
import { parseModule, createConditionalActivityTracker } from '../analyzer';


async function waitForDiagnostics(document: vscode.TextDocument, check: () => boolean, label = 'publication'): Promise<void> {
    if (check()) { return; }
    await new Promise<void>((resolve, reject) => {
        const timeout = setTimeout(() => {
            listener.dispose();
            reject(new Error(`Timed out waiting for ${label}: ${JSON.stringify(vscode.languages.getDiagnostics(document.uri).map(d => ({ code: d.code, line: d.range.start.line, message: d.message })))}`));
        }, 15_000);
        const listener = vscode.languages.onDidChangeDiagnostics(event => {
            if (event.uris.some(uri => uri.toString() === document.uri.toString()) && check()) {
                clearTimeout(timeout);
                listener.dispose();
                resolve();
            }
        });
    });
}

// Opt-in actual editor probe; the original workbook is only copied and read.
suite('Live diagnostics performance', () => {
    test('publishes and clears errors, warnings and information across large-class edits', async function () {
        const original = process.env.XLIDE_LIVE_DIAGNOSTICS_PERF_WORKBOOK;
        if (!original) { this.skip(); return; }
        this.timeout(180_000);
        await activate();
        await closeAllEditors();
        const filePath = path.join(workspaceRoot(), 'LiveDiagnosticsPerformance.xlsm');
        fs.copyFileSync(original, filePath);
        const module = [...readModules(filePath)].sort((a, b) => (b.source?.length ?? 0) - (a.source?.length ?? 0))[0];
        assert.ok(module.source);
        const parsed = parseModule(module.source);
        const activity = createConditionalActivityTracker(parsed);
        const procedure = parsed.members.find(member => member.kind === 'Procedure'
            && !activity?.isInactive(member.span) && member.body.length > 2);
        assert.ok(procedure?.kind === 'Procedure');
        const document = await open(encodeModuleUri(filePath, module.name));
        assert.ok(document.lineCount > 8000, 'probe requires a large module');
        const editor = vscode.window.activeTextEditor!;
        const at = document.positionAt(procedure.body[1].span.start);
        const codeOf = (diagnostic: vscode.Diagnostic): string => typeof diagnostic.code === 'object'
            ? String(diagnostic.code.value) : String(diagnostic.code);
        const errors = () => vscode.languages.getDiagnostics(document.uri).filter(d =>
            codeOf(d) === 'string-arithmetic-coercion' && d.range.start.line === at.line);
        let previous = Date.now();
        let maximumGap = 0;
        const heartbeat = setInterval(() => {
            const now = Date.now();
            maximumGap = Math.max(maximumGap, now - previous);
            previous = now;
        }, 25);
        const times: { appearMs: number; clearMs: number; enterMs: number }[] = [];
        try {
            // Drain initial analysis before timing edits.
            await until(() => vscode.languages.getDiagnostics(document.uri).length > 0 ? true : undefined,
                'initial workbook diagnostics should publish', 30_000);
            await new Promise(resolve => setTimeout(resolve, 500));
            const existingSemanticCodes = vscode.languages.getDiagnostics(document.uri)
                .filter(d => d.severity !== vscode.DiagnosticSeverity.Error).map(codeOf);
            for (let round = 0; round < 3; round++) {
                editor.selection = new vscode.Selection(at, at);
                const start = Date.now();
                const text = 'MsgBox "hello world"-';
                assert.ok(await editor.edit(edit => edit.insert(at, text)));
                const end = at.translate(0, text.length);
                editor.selection = new vscode.Selection(end, end);
                const enterStart = Date.now();
                await vscode.commands.executeCommand('type', { text: '\n' });
                const enterMs = Date.now() - enterStart;
                const nextLine = at.translate(1, 0);
                editor.selection = new vscode.Selection(nextLine, nextLine);
                await waitForDiagnostics(document, () => errors().length > 0);
                assert.equal(errors().length, 1, 'publish one early finding for the trailing minus');
                const currentCodes = vscode.languages.getDiagnostics(document.uri).map(codeOf);
                for (const code of existingSemanticCodes) {
                    assert.ok(currentCodes.includes(code), `early publication dropped existing ${code}`);
                }
                const appearMs = Date.now() - start;
                const clearStart = Date.now();
                assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(at, nextLine))));
                editor.selection = new vscode.Selection(at, at);
                await waitForDiagnostics(document, () => errors().length === 0);
                times.push({ appearMs, clearMs: Date.now() - clearStart, enterMs });
            }

            console.log(JSON.stringify({ trailingMinusTimes: times }));
            // Exact screenshot location: a widely called Function near line
            // 840, with declaration errors changing into an assignment error.
            const value = parsed.members.find(member => member.kind === 'Procedure' && member.name === 'Value');
            assert.ok(value?.kind === 'Procedure');
            const screenshotAt = document.positionAt(module.source.lastIndexOf('End Function', value.span.end - 1));
            const screenshotEnd = screenshotAt.translate(2, 0);
            const screenshotDiagnostics = () => vscode.languages.getDiagnostics(document.uri).filter(d =>
                d.range.start.line >= screenshotAt.line && d.range.start.line < screenshotEnd.line);
            const screenshotErrors = () => screenshotDiagnostics().filter(d => d.severity === vscode.DiagnosticSeverity.Error);
            const unreadWorkbook = () => screenshotDiagnostics().some(d => codeOf(d) === 'variable-never-read');
            const screenshotWarmup = Date.now();
            assert.ok(await editor.edit(edit => edit.insert(screenshotAt, '    Dim wb As Workbook\n    Set wb = ThisWorkbook\n')));
            editor.selection = new vscode.Selection(screenshotEnd, screenshotEnd);
            await waitForDiagnostics(document, unreadWorkbook, 'screenshot information warmup');
            const screenshotWarmupMs = Date.now() - screenshotWarmup;
            const screenshotTimes: { transition: string; ms: number }[] = [];
            for (let round = 0; round < 3; round++) {
                for (const [declaration, transition, ready] of [
                    ['Debug wb As Worksheet', 'declaration typo', () => screenshotErrors().length === 3
                        && screenshotErrors().every(d => codeOf(d) === 'undeclared-variable') && !unreadWorkbook()],
                    ['Dim wb As Worksheet', 'declaration repair / assignment error', () => screenshotErrors().length === 1
                        && codeOf(screenshotErrors()[0]) === 'assignment-object-type-mismatch' && unreadWorkbook()],
                    ['Dim wb As Workbook', 'assignment repair', () => screenshotErrors().length === 0 && unreadWorkbook()],
                ] as const) {
                    const start = Date.now();
                    assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(screenshotAt, screenshotAt.translate(1, 0)),
                        `    ${declaration}\n`)));
                    editor.selection = new vscode.Selection(screenshotEnd, screenshotEnd);
                    await waitForDiagnostics(document, ready, `screenshot ${transition}`);
                    screenshotTimes.push({ transition, ms: Date.now() - start });
                    console.log(JSON.stringify(screenshotTimes.at(-1)));
                }
            }
            assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(screenshotAt, screenshotEnd))));
            await waitForDiagnostics(document, () => !unreadWorkbook(), 'screenshot removal');
            // An isolated procedure in the same large class lets us exercise
            // every semantic family without changing a heavily used helper's
            // return values and making all of its callers dirty too.
            const endOfDocument = document.positionAt(document.getText().length);
            const probe = '\nPrivate Sub XlidePerfErrors()\n    Debug.Print 0\n    Debug.Print (1\nEnd Sub\n';
            const createStart = Date.now();
            assert.ok(await editor.edit(edit => edit.insert(endOfDocument, probe)));
            const probeStart = document.getText().indexOf('    Debug.Print (1', document.offsetAt(endOfDocument));
            const semanticAt = document.positionAt(probeStart);
            const sentinel = () => vscode.languages.getDiagnostics(document.uri)
                .some(d => codeOf(d) === 'unbalanced-parens' && d.range.start.line === semanticAt.line);
            await waitForDiagnostics(document, sentinel, 'new procedure warmup');
            const procedureCreateMs = Date.now() - createStart;
            const paren = document.positionAt(probeStart + '    Debug.Print '.length);
            assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(paren, paren.translate(0, 1)))));
            await waitForDiagnostics(document, () => !sentinel(), 'warmup clear');
            const semanticTimes: { code: string; appearMs: number; clearMs: number }[] = [];
            const scenarios = [
                ['Debug.Print xlideMissingIdentifier', 'undeclared-variable'],
                ['xlideMissingProcedure', 'unknown-call'],
                ['Debug.Print Left$("abc")', 'argument-count'],
                ['Dim ws As Worksheet\nws.XlideMissingMember = 1', 'member-not-found'],
                ['Dim n As Long\nn = "abc"', 'assignment-type-mismatch'],
                ['Dim a(1) As Long\nDebug.Print a(2)', 'array-subscript-out-of-bounds'],
                ['Dim c As Collection\nDebug.Print c.Count', 'object-variable-not-set'],
                ['Debug.Print 1 / 0', 'division-by-zero'],
            ];
            for (const [text, code] of scenarios) {
                const insertedLines = text.split('\n').length;
                const after = semanticAt.translate(insertedLines, 0);
                const matching = () => vscode.languages.getDiagnostics(document.uri).filter(d => codeOf(d) === code
                    && d.severity === vscode.DiagnosticSeverity.Error && d.range.start.line >= semanticAt.line
                    && d.range.start.line < after.line);
                editor.selection = new vscode.Selection(semanticAt, semanticAt);
                const start = Date.now();
                assert.ok(await editor.edit(edit => edit.insert(semanticAt, text + '\n')));
                editor.selection = new vscode.Selection(after, after);
                await waitForDiagnostics(document, () => matching().length > 0, `${code} appearance`);
                const appearMs = Date.now() - start;
                assert.equal(matching().length, 1, `one ${code} finding`);
                const clearStart = Date.now();
                assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(semanticAt, after))));
                editor.selection = new vscode.Selection(semanticAt, semanticAt);
                await waitForDiagnostics(document, () => matching().length === 0, `${code} clear`);
                semanticTimes.push({ code, appearMs, clearMs: Date.now() - clearStart });
                console.log(JSON.stringify(semanticTimes.at(-1)));
            }
            const nonErrorTimes: { code: string; appearMs: number; clearMs: number }[] = [];
            const nonErrorScenarios = [
                ['Dim xlideUnused As Long', 'unused-variable', vscode.DiagnosticSeverity.Information],
                ['Dim xlideWritten As Long\nxlideWritten = 1', 'variable-never-read', vscode.DiagnosticSeverity.Information],
                ['Exit Sub\nDebug.Print 1', 'unreachable-code', vscode.DiagnosticSeverity.Information],
                ['Dim xlideWorkbook As Workbook\nIf TypeOf xlideWorkbook Is Worksheet Then\nDebug.Print 1\nEnd If', 'typeof-is-always-false', vscode.DiagnosticSeverity.Warning],
                ['ActiveSheet.Range("A1").Formula = "=1+"', 'formula-string-unparsed', vscode.DiagnosticSeverity.Warning],
            ] as const;
            for (const [text, code, severity] of nonErrorScenarios) {
                const after = semanticAt.translate(text.split('\n').length, 0);
                const matching = () => vscode.languages.getDiagnostics(document.uri).filter(d => codeOf(d) === code
                    && d.severity === severity && d.range.start.line >= semanticAt.line && d.range.start.line < after.line);
                const start = Date.now();
                assert.ok(await editor.edit(edit => edit.insert(semanticAt, text + '\n')));
                editor.selection = new vscode.Selection(after, after);
                await waitForDiagnostics(document, () => matching().length > 0, `${code} appearance`);
                const appearMs = Date.now() - start;
                const clearStart = Date.now();
                assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(semanticAt, after))));
                await waitForDiagnostics(document, () => matching().length === 0, `${code} clear`);
                nonErrorTimes.push({ code, appearMs, clearMs: Date.now() - clearStart });
                console.log(JSON.stringify(nonErrorTimes.at(-1)));
            }

            const surfaceTimes: { transition: string; ms: number }[] = [{ transition: 'procedure addition with syntax error', ms: procedureCreateMs }];
            const probeInformation = () => vscode.languages.getDiagnostics(document.uri).some(d => codeOf(d) === 'unused-procedure'
                && d.message.includes('XlidePerfErrors'));
            assert.ok(probeInformation(), 'new private procedure has an information finding');
            const removeProcedureStart = Date.now();
            assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(endOfDocument, document.positionAt(document.getText().length)))));
            await waitForDiagnostics(document, () => !probeInformation(), 'procedure information removal');
            surfaceTimes.push({ transition: 'procedure removal', ms: Date.now() - removeProcedureStart });

            const declarationAt = document.positionAt(procedure.span.start);
            const globalInformation = () => vscode.languages.getDiagnostics(document.uri).some(d => codeOf(d) === 'unused-variable'
                && d.message.includes('XlidePerfGlobal'));
            const declareStart = Date.now();
            assert.ok(await editor.edit(edit => edit.insert(declarationAt, 'Private XlidePerfGlobal As Long\n')));
            await waitForDiagnostics(document, globalInformation, 'module declaration information');
            surfaceTimes.push({ transition: 'module declaration addition', ms: Date.now() - declareStart });
            const removeDeclarationStart = Date.now();
            assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(declarationAt, declarationAt.translate(1, 0)))));
            await waitForDiagnostics(document, () => !globalInformation(), 'module declaration information removal');
            surfaceTimes.push({ transition: 'module declaration removal', ms: Date.now() - removeDeclarationStart });

            const closerAt = document.positionAt(document.getText().lastIndexOf('End Function', value.span.end - 1));
            const closerError = () => vscode.languages.getDiagnostics(document.uri).some(d => codeOf(d) === 'mismatched-end-keyword'
                && d.message.includes("'Function Value'"));
            const breakCloserStart = Date.now();
            assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(closerAt, closerAt.translate(0, 'End Function'.length)), 'End Sub')));
            editor.selection = new vscode.Selection(closerAt.translate(1, 0), closerAt.translate(1, 0));
            await waitForDiagnostics(document, closerError, 'procedure closer error');
            surfaceTimes.push({ transition: 'procedure closer error', ms: Date.now() - breakCloserStart });
            const repairCloserStart = Date.now();
            assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(closerAt, closerAt.translate(0, 'End Sub'.length)), 'End Function')));
            await waitForDiagnostics(document, () => !closerError(), 'procedure closer repair');
            surfaceTimes.push({ transition: 'procedure closer repair', ms: Date.now() - repairCloserStart });

            const burstInformation = () => vscode.languages.getDiagnostics(document.uri).some(d =>
                codeOf(d) === 'variable-never-read' && d.message.includes('xlideBurstInfo'));
            const burstNext = new vscode.Position(at.line + 1, 0);
            const burstAfter = new vscode.Position(at.line + 2, 0);
            assert.ok(await editor.edit(edit => edit.insert(at, 'Dim xlideBurstInfo As Long\nDebug.Print xlideBurstMissing\n')));
            editor.selection = new vscode.Selection(burstAfter, burstAfter);
            await new Promise(resolve => setTimeout(resolve, 180));
            const burstStart = Date.now();
            assert.ok(await editor.edit(edit => edit.replace(new vscode.Range(burstNext, burstAfter), 'xlideBurstInfo = 1\n')));
            await waitForDiagnostics(document, burstInformation, 'latest burst snapshot');
            surfaceTimes.push({ transition: 'rapid edit supersedes an in-flight error', ms: Date.now() - burstStart });
            assert.ok(!vscode.languages.getDiagnostics(document.uri).some(d => d.message.includes('xlideBurstMissing')), 'superseded error must not publish');
            const burstClearStart = Date.now();
            assert.ok(await editor.edit(edit => edit.delete(new vscode.Range(at, burstAfter))));
            await waitForDiagnostics(document, () => !burstInformation(), 'burst information clear');
            surfaceTimes.push({ transition: 'rapid edit information clear', ms: Date.now() - burstClearStart });
            console.log(JSON.stringify({ surfaceTimes }));
            // Stay on an incomplete syntax line until both passes finish,
            // then leave it without editing. The cached early finding must
            // appear immediately even if the full pass withheld it.
            assert.ok(await editor.edit(edit => edit.insert(at, 'Debug.Print (1\n')));
            editor.selection = new vscode.Selection(at, at);
            await new Promise(resolve => setTimeout(resolve, 4000));
            const syntaxError = () => vscode.languages.getDiagnostics(document.uri).filter(d =>
                codeOf(d) === 'unbalanced-parens' && d.range.start.line === at.line);
            assert.equal(syntaxError().length, 0, 'hold syntax while the cursor stays on its line');
            const leaveStart = Date.now();
            const nextLine = at.translate(1, 0);
            editor.selection = new vscode.Selection(nextLine, nextLine);
            await waitForDiagnostics(document, () => syntaxError().length > 0);
            const cachedRevealMs = Date.now() - leaveStart;
            assert.ok(cachedRevealMs < 500, `cached line-exit reveal took ${cachedRevealMs} ms`);
            console.log(JSON.stringify({ liveDiagnosticsPerformance: { module: module.name, times, screenshotWarmupMs, screenshotTimes, semanticTimes, nonErrorTimes, surfaceTimes, cachedRevealMs, maximumGap } }));
            await vscode.commands.executeCommand('xlide.copyPerformanceSnapshot');
            console.log(await vscode.env.clipboard.readText());
            assert.ok(times.every(t => t.appearMs < 1000 && t.clearMs < 1000), `trailing-minus squiggle times: ${JSON.stringify(times)}`);
            assert.ok(semanticTimes.every(t => t.appearMs < 1000 && t.clearMs < 1000), `semantic squiggle times: ${JSON.stringify(semanticTimes)}`);
            assert.ok(screenshotTimes.every(t => t.ms < 1000), `screenshot activation/deactivation times: ${JSON.stringify(screenshotTimes)}`);
            assert.ok(surfaceTimes.every(t => t.ms < 1000), `declaration/procedure/closer times: ${JSON.stringify(surfaceTimes)}`);
            assert.ok(nonErrorTimes.every(t => t.appearMs < 1000 && t.clearMs < 1000), `warning/information times: ${JSON.stringify(nonErrorTimes)}`);
            assert.ok(maximumGap < 2000, `extension host stalled for ${maximumGap} ms`);
        } finally {
            clearInterval(heartbeat);
            await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor');
            await closeAllEditors();
        }
    });
});
