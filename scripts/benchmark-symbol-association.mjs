// Reproducible pure-analyzer timings; no VS Code or Office application required.
// Run: node scripts/benchmark-symbol-association.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-analyzer-benchmark-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { parseModule } from './src/analyzer/parser/parseModule';
        export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';
        `, resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(bundle, result.outputFiles[0].contents);
    api = createRequire(import.meta.url)(bundle);
} finally {
    try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    rmdirSync(scratch);
}
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
for (const count of [500, 2000, 5000]) {
    const defSource = 'DefLng A-Z\n' + Array.from({ length: count }, (_, i) =>
        'Public Sub P' + i + '(ByVal x As Long)\nDim declared As Long\ndeclared = x\nimplicit = x\nEnd Sub\n').join('\n');
    const attrSource = 'Option Explicit\n' + Array.from({ length: count }, (_, i) =>
        'Attribute P' + i + '.VB_Description = "Work"\n').join('')
        + Array.from({ length: count }, (_, i) => 'Public Sub P' + i + '()\nEnd Sub\n').join('\n');
    const defModule = api.parseModule(defSource), attrModule = api.parseModule(attrSource);
    measure('implicit-symbols/' + count, () => {
        const result = api.buildModuleSymbols('M', 'standard', defSource, { parsedModule: defModule });
        if (result.implicitLocals?.size !== count) throw new Error('Missing implicit locals');
    });
    measure('attributed-symbols/' + count, () => {
        const result = api.buildModuleSymbols('M', 'standard', attrSource, { parsedModule: attrModule });
        if (!result.root.children?.every(s => s.attributes?.length === 1)) throw new Error('Missing attributes');
    });
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, rounds, rows }, null, 2));
