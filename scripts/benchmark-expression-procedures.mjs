import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-expression-procedures-')), path = join(scratch, 'api.cjs');
try {
    const plugins = baseline ? [{ name: 'baseline', setup(b) {
        b.onLoad({ filter: /resolveExpressionType\.ts$/ }, args => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/expression/resolveExpressionType.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) }));
    } }] : [];
    const bundle = await build({ plugins, stdin: { contents: "export {resolveExpressionType} from './src/analyzer/expression/resolveExpressionType';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(path, bundle.outputFiles[0].contents);
    const api = createRequire(import.meta.url)(path), rows = [];
    for (const procedures of [1, 100, 1000]) for (const target of ['first', 'last', 'spread']) for (const queries of [1, 100]) for (const scope of ['warm-binding', 'fresh-binding']) {
        const parts = [], spans = [];
        let offset = 0;
        for (let i = 0; i < procedures; i++) {
            const prefix = 'Function F' + i + '() As Long\nF' + i + ' = ';
            const part = prefix + '1\nEnd Function\n';
            spans.push({ start: offset + prefix.length, end: offset + prefix.length + 1 });
            parts.push(part); offset += part.length;
        }
        const source = parts.join('');
        const selections = Array.from({ length: queries }, (_, i) => spans[target === 'first' ? 0 : target === 'last' ? procedures - 1 : i % procedures]);
        const expected = selections.map(() => ({ type: 'Long', isObject: false, complete: true }));
        const ctx = { moduleName: 'Bench' + procedures + target + queries + scope };
        const times = [];
        for (let round = -3; round < 15; round++) {
            const text = scope === 'fresh-binding' ? source + "' fresh " + round + '\n' : source;
            const start = performance.now();
            const result = selections.map(span => api.resolveExpressionType(text, span, ctx));
            const elapsed = performance.now() - start;
            assert.deepEqual(result, expected);
            if (round >= 0) times.push(elapsed);
        }
        times.sort((a, b) => a - b);
        rows.push({ procedures, target, queries, scope, medianMs: times[7], p95Ms: times[14], resultSha256: createHash('sha256').update(JSON.stringify(expected)).digest('hex') });
    }
    console.log(JSON.stringify({ baseline: baseline ?? 'working-tree', node: process.version, rounds: 15, warmups: 3, scope: 'actual public scalar literal resolver; source/context/span setup and independent output assertions outside timer; fresh-binding changes source identity and includes parsing/binding/index construction; no IO or UI', rows }, null, 2));
} finally {
    if (existsSync(path)) unlinkSync(path);
    rmdirSync(scratch);
}
