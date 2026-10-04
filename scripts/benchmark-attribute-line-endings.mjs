// Run baseline and fixed in separate processes; --baseline=<commit> replaces only the writer.
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const output = process.argv.find(arg => arg.startsWith('--output='))?.slice(9);
const plugins = baseline ? [{ name: 'baseline-writer', setup(b) {
    b.onLoad({ filter: /attributeRewriter\.ts$/ }, args => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/annotations/attributeRewriter.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) }));
} }] : [];
const bundle = await build({ plugins, stdin: { contents: "export {readAttributeAnnotations} from './src/analyzer/annotations/attributeAnnotations'; export {applyAttributeAnnotations} from './src/analyzer/annotations/attributeRewriter';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
const path = '.audit/attribute-ending-benchmark-' + (baseline ? 'baseline' : 'fixed') + '.cjs';
writeFileSync(path, bundle.outputFiles[0].contents);
const api = createRequire(import.meta.url)('../' + path);
const rows = [];
for (const count of [1, 100, 1000]) for (const eol of ['\n', '\r\n']) for (const shape of ['aligned', 'changed', 'inserted', 'module-only']) {
    const lines = ['Attribute VB_Name = "M"'];
    if (shape === 'module-only') lines.push('Attribute VB_Description = "doc"', "'@ModuleDescription(\"doc\")");
    for (let i = 0; i < count; i++) {
        if (shape !== 'module-only') lines.push("'@Description(\"doc\")");
        lines.push('Sub P' + i + '()');
        if (shape === 'aligned' || shape === 'changed') lines.push('Attribute P' + i + '.VB_Description = "' + (shape === 'aligned' ? 'doc' : 'old') + '"');
        lines.push('End Sub');
    }
    const source = lines.join(eol) + eol;
    const run = () => api.applyAttributeAnnotations(source, api.readAttributeAnnotations(source));
    const expected = run();
    assert.equal(expected.changes.length, shape === 'changed' || shape === 'inserted' ? count : 0);
    assert.deepEqual(expected.skipped, []);
    if (shape === 'aligned' || shape === 'module-only') assert.equal(expected.text, source);
    if (shape === 'changed') assert.equal(expected.text, source.replaceAll('= "old"', '= "doc"'));
    if (shape === 'inserted') assert.equal(expected.text, source.replace(/Sub (P\d+)\(\)/g, (_, name) => 'Sub ' + name + '()' + eol + 'Attribute ' + name + '.VB_Description = "doc"'));
    const times = [];
    const batch = count === 1 ? 200 : count === 100 ? 20 : 2;
    for (let round = 0; round < 18; round++) {
        const start = performance.now();
        let result;
        for (let iteration = 0; iteration < batch; iteration++) result = run();
        const elapsed = (performance.now() - start) / batch;
        assert.deepEqual(result, expected);
        if (round >= 3) times.push(elapsed);
    }
    times.sort((a, b) => a - b);
    rows.push({ count, eol: eol === '\n' ? 'LF' : 'CRLF', shape, medianMs: times[7], p95Ms: times[14], hash: createHash('sha256').update(JSON.stringify(expected)).digest('hex') });
}
const result = { baseline: baseline ?? null, node: process.version, rounds: 15, warmups: 3, scope: 'actual reader plus writer, source prepared outside timer; no IO or Office', rows };
if (output) writeFileSync(output, JSON.stringify(result, null, 2));
else console.log(JSON.stringify(result, null, 2));
