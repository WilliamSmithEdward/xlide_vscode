import { expect, it } from 'vitest';
import { readModule } from '../src/vba/projectService';
import { buildLiveVbaProjectIndex } from '../src/vbaProjectAnalysis';

it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('compares class member surfaces with and without diagnostic value facts', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const facts = [], editor = [];
    for (let i = 0; i < 21; i++) {
        const source = original + "\r\n' surface benchmark " + i;
        const input = [{ moduleName: 'ROneCOne', moduleKind: 'class' as const, source }];
        const full = buildLiveVbaProjectIndex(input);
        const fast = buildLiveVbaProjectIndex(input);
        const measureFull = () => { const start = performance.now(); const result = full.projectMemberSurfaces('ROneCOne'); facts.push(performance.now() - start); return result; };
        const measureEditor = () => { const start = performance.now(); const result = fast.projectMemberSurfaces('ROneCOne', { includeClassValueFacts: false }); editor.push(performance.now() - start); return result; };
        const [fullResult, editorResult] = i % 2 ? [measureFull(), measureEditor()] : (() => { const fastResult = measureEditor(); return [measureFull(), fastResult]; })();
        expect(editorResult).toEqual(fullResult.map(type => ({ ...type, members: type.members.map(({ knownValue, ...member }) => member) })));
    }
    const summarize = (samples: number[]) => { const sorted = samples.sort((a,b) => a-b); return { medianMs: sorted[10], maxMs: sorted[20] }; };
    process.stdout.write('Actual class member surfaces: ' + JSON.stringify({ withValueFacts: summarize(facts), editor: summarize(editor) }) + '\n');
}, 60000);
