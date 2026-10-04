import { expect, it } from 'vitest';
import { readModule } from '../src/vba/projectService';
import { tokenize, tokenizeCached } from '../src/analyzer/lexer/tokenize';

it.skipIf(!process.env.XLIDE_PERF_WORKBOOK)('compares actual class lexing at early, middle and late edit positions', () => {
    const original = readModule(process.env.XLIDE_PERF_WORKBOOK!, 'ROneCOne').source;
    const results = [];
    for (const fraction of [0.01, 0.5, 0.99]) {
        const offset = Math.floor(original.length * fraction);
        for (const replacement of [false, true]) {
            const samples = [];
            for (let i = 0; i < 15; i++) {
                tokenizeCached(original);
                const inserted = String.fromCharCode(65 + i);
                const changed = original.slice(0, offset) + inserted + original.slice(offset + (replacement ? 1 : 0));
                let before = performance.now();
                const actual = tokenizeCached(changed);
                const cachedMs = performance.now() - before;
                before = performance.now();
                const expected = tokenize(changed);
                const fullMs = performance.now() - before;
                if (i === 0) { expect(actual).toStrictEqual(expected); }
                samples.push({ cachedMs, fullMs });
            }
            const median = (key: 'cachedMs' | 'fullMs') => samples.map(s => s[key]).sort((a,b) => a-b)[7];
            results.push({ fraction, replacement, cachedMs: median('cachedMs'), fullMs: median('fullMs') });
        }
    }
    process.stdout.write('Actual class incremental lexer: ' + JSON.stringify(results) + '\n');
}, 60000);
