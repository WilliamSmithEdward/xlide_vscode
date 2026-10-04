import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname, relative } from 'node:path';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';

const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const root = process.cwd();
const scratch = mkdtempSync(join(tmpdir(), 'xlide-reference-prefix-'));
try {
    const plugins = baseline ? [{ name: 'baseline', setup(builder) {
        builder.onLoad({ filter: /vbaReferenceResolution\.ts$/ }, args => {
            const file = relative(root, args.path).replaceAll('\\', '/');
            return { contents: execFileSync('git', ['show', baseline + ':' + file], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) };
        });
    } }] : [];
    const bundle = await build({ plugins, stdin: { contents: "export {collectSymbolReferences} from './src/vbaReferenceResolution'; export {buildVbaProjectIndex} from './src/vbaProjectAnalysis'; export {tokenizeCached} from './src/analyzer/lexer/tokenize';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    const path = join(scratch, 'api.cjs');
    writeFileSync(path, bundle.outputFiles[0].contents);
    const api = createRequire(import.meta.url)(path);
    const rows = [];
    for (const count of [1, 10, 100, 1000]) {
        for (const style of ['qualified', 'bare', 'continued', 'with', 'typed-with']) {
            const statement = { qualified: '    Library.Greet', bare: '    Greet', continued: '    Library. _\n        Greet', with: '    .Greet', 'typed-with': '    .Greet' }[style];
            const caller = 'Sub UseIt()\n' + (style === 'with' ? '    With Library\n' : style === 'typed-with' ? '    Dim receiver As Library\n    Set receiver = New Library\n    With receiver\n' : '') + Array(count).fill(statement).join('\n') + '\n' + (style.endsWith('with') ? '    End With\n' : '') + 'End Sub\n';
            const modules = [{ moduleName: 'Library', type: style === 'typed-with' ? 'class' : undefined, source: 'Public Sub Greet()\nEnd Sub\n' }, { moduleName: 'Caller', source: caller }];
            const project = api.buildVbaProjectIndex(modules);
            const byModule = new Map(modules.map(mod => [mod.moduleName.toLowerCase(), mod]));
            const at = modules[0].source.indexOf('Greet');
            const times = [];
            let expected;
            for (let i = -3; i < 15; i++) {
                const start = performance.now();
                const result = api.collectSymbolReferences(byModule, project, modules, modules[0].source, 'Library', modules[0], 'Greet', at + 5, at, true);
                const elapsed = performance.now() - start;
                assert.equal(result.hasSymbol, true);
                assert.equal(result.references.length, count + 1);
                assert.equal(result.ambiguous.length, 0);
                if (i === -3) { expected = result; } else { assert.deepEqual(result, expected); }
                if (i >= 0) { times.push(elapsed); }
            }
            times.sort((a, b) => a - b);
            rows.push({ count, style, medianMs: times[7], p95Ms: times[14], resultSha256: createHash('sha256').update(JSON.stringify(expected)).digest('hex') });
        }
    }
    console.log(JSON.stringify({ baseline: baseline ?? 'working-tree', node: process.version, rounds: 15, warmups: 3, scope: 'Complete actual reference collection with warm project and source caches; project/source preparation and complete output assertions outside timer; no Office/UI execution', rows }, null, 2));
} finally {
    // The directory is a fresh task-owned temporary directory above.
    const path = join(scratch, 'api.cjs');
    if (existsSync(path)) { unlinkSync(path); }
    rmdirSync(scratch);
}
