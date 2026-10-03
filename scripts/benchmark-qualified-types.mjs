// Run in the target checkout: node scripts/benchmark-qualified-types.mjs [--rounds=15]
import { buildSync } from 'esbuild';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-qualified-types-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { resolveTypeName, resolveTypeCompletions } from './src/analyzer/completion/typeCompletion';
        export { resolveTypeSemanticTokens } from './src/analyzer/semantic/typeSemanticTokens';
        export { getExcelObjectModel } from './src/analyzer/host/excelObjectModel';
        `, resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(bundle, result.outputFiles[0].contents);
    api = createRequire(import.meta.url)(bundle);
} finally {
    try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    rmdirSync(scratch);
}
const model = api.getExcelObjectModel();
const rows = [];
function measure(name, run) {
    for (let i = 0; i < 3; i++) run();
    const samples = [];
    for (let i = 0; i < rounds; i++) {
        const start = performance.now(); run(); samples.push(performance.now() - start);
    }
    samples.sort((a, b) => a - b);
    rows.push({ name, medianMs: +samples[Math.floor(samples.length / 2)].toFixed(3),
        p95Ms: +samples[Math.ceil(samples.length * .95) - 1].toFixed(3) });
}
measure('qualified-lookups/1000', () => {
    for (let i = 0; i < 1000; i++) {
        if (api.resolveTypeName('Excel.Range', { model })?.kind !== 'host') throw new Error('Missing Range');
    }
});
measure('qualified-completions/100', () => {
    for (let i = 0; i < 100; i++) {
        if (!api.resolveTypeCompletions('Dim x As Excel.R', 16, { model }).some(c => c.name === 'Range')) throw new Error('Missing Range');
    }
});
for (const count of [100, 500, 2000]) {
    const source = 'Option Explicit\n' + Array.from({ length: count }, (_, i) =>
        'Public Sub P' + i + '()\nDim r As Excel.Range\nEnd Sub\n').join('\n');
    measure('semantic-types/' + count, () => {
        if (api.resolveTypeSemanticTokens(source, { model }).length !== count) throw new Error('Missing semantic types');
    });
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, rounds, rows }, null, 2));
