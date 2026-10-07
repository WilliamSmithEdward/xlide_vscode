import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-call-parens-'));
const path = join(scratch, 'api.cjs');
try {
    const plugins = baseline ? [{ name: 'baseline', setup(b) {
        b.onLoad({ filter: /callContext\.ts$/ }, args => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/call/callContext.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) }));
    } }] : [];
    const bundle = await build({ plugins, stdin: { contents: "export {standaloneEmptyParenthesizedCallStatement, standaloneMultiArgParenthesizedCallStatement} from './src/analyzer/call/callContext'; export {statementTokensCached} from './src/analyzer/lexer/tokenHelpers';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(path, bundle.outputFiles[0].contents);
    const api = createRequire(import.meta.url)(path), rows = [];
    for (const depth of [0, 5, 100, 1000]) for (const shape of ['nested', 'unclosed', 'receiver', 'multi']) for (const consumer of ['standaloneEmptyParenthesizedCallStatement', 'standaloneMultiArgParenthesizedCallStatement']) for (const scope of ['cached-tokens', 'fresh-tokens']) {
        const nested = 'G('.repeat(depth) + '1' + ')'.repeat(depth);
        const source = shape === 'unclosed' ? 'F(' + 'G('.repeat(depth) + '1' : shape === 'receiver' ? 'F(' + nested + ').Done()' : 'F(' + nested + (shape === 'multi' ? ', 2)' : ')');
        const span = { start: 0, end: source.length };
        const expected = shape === 'receiver' && consumer === 'standaloneEmptyParenthesizedCallStatement'
            ? { name: 'Done', isMember: true, startsWithLeadingDot: false, calleeEndOffset: source.length - 2, emptyParensSpan: { start: source.length - 2, end: source.length }, span: { start: source.length - 6, end: source.length } }
            : shape === 'multi' && consumer === 'standaloneMultiArgParenthesizedCallStatement'
                ? { name: 'F', isMember: false, qualifier: undefined, argumentCount: 2, span: { start: 0, end: source.length } } : undefined;
        const tokens = api.statementTokensCached(source, span);
        for (const token of tokens) Object.freeze(token);
        Object.freeze(tokens);
        const times = [], batch = scope === 'cached-tokens' && depth <= 5 ? 100 : 1;
        for (let round = -3; round < 15; round++) {
            const text = scope === 'fresh-tokens' ? source + "\n' fresh " + round : source;
            let result;
            const start = performance.now();
            for (let i = 0; i < batch; i++) result = api[consumer](text, span);
            const elapsed = (performance.now() - start) / batch;
            assert.deepEqual(result, expected);
            if (round >= 0) times.push(elapsed);
        }
        times.sort((a, b) => a - b);
        rows.push({ depth, shape, consumer, scope, medianMs: times[7], p95Ms: times[14], resultSha256: createHash('sha256').update(JSON.stringify(expected) ?? 'undefined').digest('hex') });
    }
    console.log(JSON.stringify({ baseline: baseline ?? 'working-tree', node: process.version, rounds: 15, warmups: 3, scope: 'Actual public consumers; source/span prepared outside timer, fresh source comment changes cache identity and includes statement lexing; independent full result assertions outside timer; no Office or IO', rows }, null, 2));
} finally {
    if (existsSync(path)) unlinkSync(path);
    rmdirSync(scratch);
}
