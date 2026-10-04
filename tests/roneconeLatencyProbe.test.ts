import { it } from 'vitest';
import { readModule } from '../src/vba/projectService';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { editorModuleSymbols } from '../src/analyzer/symbols/editorModuleSymbols';
import { buildLiveVbaProjectIndex, projectEditorSymbolContextForModule } from '../src/vbaProjectAnalysis';
import { resolveKeywordCompletions } from '../src/analyzer/completion/keywordCompletion';
import { resolveMemberCompletions } from '../src/analyzer/completion/memberAccess';

it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('profiles the actual large class after fresh edits', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const samples = Array.from({ length: 5 }, (_, i) => {
        const source = original + '\r\nPublic Sub XlideLatencyProbe()\r\n ThisWorkbook.Sheets(1).Na\r\nEnd Sub\r\n\' change ' + i;
        const times: Record<string, number> = {};
        const measure = <T>(name: string, run: () => T) => {
            const before = performance.now(); const result = run(); times[name] = performance.now() - before; return result;
        };
        const tokens = measure('tokenize', () => tokenizeCached(source));
        const ast = measure('parse', () => parseModule(source));
        measure('symbols', () => editorModuleSymbols('ROneCOne', 'class', source));
        const project = measure('singleModuleIndex', () => buildLiveVbaProjectIndex([{ moduleName: 'ROneCOne', moduleKind: 'class', source }]));
        measure('editorProjection', () => projectEditorSymbolContextForModule(project, 'ROneCOne'));
        const offset = source.indexOf('ThisWorkbook.Sheets(1).Na', original.length) + 'ThisWorkbook.Sheets(1).Na'.length;
        measure('memberCompletion', () => resolveMemberCompletions(source, offset, { moduleName: 'ROneCOne' }));
        return { ...times, tokens: tokens.length };
    });
    process.stdout.write('Actual large class fresh-edit phases: ' + JSON.stringify({ bytes: original.length, samples }) + '\n');
});

it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('measures keyword classification on ordinary prefixes in the actual class', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const sources = Array.from({ length: 6 }, (_, i) => original + '\r\nPublic Sub XlideKeywordProbe()\r\n Latency' + i);
    for (const source of sources) { resolveKeywordCompletions(source, source.length); }
    const times = Array.from({ length: 21 }, () => {
        const before = performance.now();
        for (const source of sources) { resolveKeywordCompletions(source, source.length); }
        return (performance.now() - before) / sources.length;
    }).sort((a, b) => a-b);
    process.stdout.write('Actual class keyword benchmark: ' + JSON.stringify({ characters: original.length, medianMs: times[10], slowestBatchMs: times[20] }) + '\n');
}, 15000);
