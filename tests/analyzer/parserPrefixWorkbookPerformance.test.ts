import { expect, it } from 'vitest';
import { readModule } from '../../src/vba/projectService';
import { tokenize, tokenizeCached } from '../../src/analyzer/lexer/tokenize';
import { parseModule, parseModuleFreshForTests } from '../../src/analyzer/parser/parseModule';

// Opt-in: XLIDE_PERF_WORKBOOK names a workbook, and XLIDE_PERF_MODULE a module
// in it (ROneCOne by default), so any disposable workbook will do. Timing is
// measured first, then every edited parse is compared against an independent
// fresh parse rather than the cache, which could hand back the very AST under
// test (issue #1442).
it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('measures prefix parsing at early middle and late class positions', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, process.env.XLIDE_PERF_MODULE ?? 'ROneCOne').source;
    const results = [];
    for (const fraction of [0.01, 0.5, 0.99]) {
        const offset = Math.floor(original.length * fraction);
        const after: number[] = [];
        for (let i = 0; i < 15; i++) {
            // Seed the previous snapshot an edit starts from.
            parseModule(original);
            const source = original.slice(0, offset) + String.fromCharCode(65 + i) + original.slice(offset);
            tokenizeCached(source);
            const start = performance.now();
            const result = parseModule(source);
            after.push(performance.now() - start);
            expect(tokenizeCached(source)).toEqual(tokenize(source));
            const fresh = parseModuleFreshForTests(source);
            expect(fresh).not.toBe(result);
            expect(result).toEqual(fresh);
        }
        const median = (values: number[]) => values.sort((a,b) => a-b)[7];
        results.push({ fraction, afterMs: median(after) });
    }
    process.stdout.write('Actual class prefix parser: ' + JSON.stringify(results) + '\n');
}, 120000);
