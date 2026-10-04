import { expect, it } from 'vitest';
import { readModule } from '../src/vba/projectService';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { parseModule } from '../src/analyzer/parser/parseModule';

it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('measures prefix parsing at early middle and late class positions', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const results = [];
    for (const fraction of [0.01, 0.5, 0.99]) {
        const offset = Math.floor(original.length * fraction);
        const after: number[] = [];
        for (let i = 0; i < 15; i++) {
            parseModule(original);
            const source = original.slice(0, offset) + String.fromCharCode(65 + i) + original.slice(offset);
            tokenizeCached(source);
            const start = performance.now();
            const result = parseModule(source);
            after.push(performance.now() - start);
            // Evict large snapshots to verify against a cold full parse.
            for (let j = 0; j < 12; j++) parseModule("' eviction " + j);
            expect(result).toEqual(parseModule(source));
        }
        const median = (values: number[]) => values.sort((a,b) => a-b)[7];
        results.push({ fraction, afterMs: median(after) });
    }
    process.stdout.write('Actual class prefix parser: ' + JSON.stringify(results) + '\n');
}, 60000);
