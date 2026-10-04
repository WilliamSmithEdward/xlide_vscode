import { expect, it } from 'vitest';
import { readModule } from '../src/vba/projectService';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { collectHostMemberMethodTokens } from '../src/analyzer/semantic/typeSemanticTokens';

it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('measures host member coloring across fresh ROneCOne ASTs', () => {
    const source = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const times = [];
    let previous: ReturnType<typeof collectHostMemberMethodTokens> | undefined;
    for (let i = 0; i < 21; i++) {
        const changed = source + "\r\n' scope benchmark " + i;
        parseModule(changed);
        const start = performance.now();
        const result = collectHostMemberMethodTokens(changed);
        times.push(performance.now() - start);
        if (previous) { expect(result).toEqual(previous); }
        previous = result;
    }
    times.sort((a,b) => a-b);
    process.stdout.write('Actual class scope lookup coloring: ' + JSON.stringify({
        medianMs: times[10], maxMs: times[20], hostTokens: previous?.length,
    }) + '\n');
}, 60000);
