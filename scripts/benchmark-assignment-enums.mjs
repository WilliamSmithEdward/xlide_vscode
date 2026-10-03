// Reproducible pure-analyzer timings; no VS Code or Office application required.
// Run: node scripts/benchmark-assignment-enums.mjs [--rounds=15]
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
        export { checkAssignmentTypes } from './src/analyzer/diagnostics/rules/assignments';
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
for (const [members, assignments] of [[100, 1000], [1000, 1000], [1000, 5000]]) {
    const source = 'Option Explicit\n'+Array.from({length: members}, (_,i) => 'Private v'+i+' As Long\n').join('')
        +'Public Sub Main(ByVal x As Long)\nDim total As Long\n'
        +Array.from({length:assignments},(_,i) => 'total = x + '+(i%10)+'\n').join('')+'Debug.Print total\nEnd Sub';
    const module = api.parseModule(source);
    const symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: module });
    measure('assignment-types/' + members + '-members/' + assignments + '-assignments', () => {
        const diagnostics=[];
        api.checkAssignmentTypes(source,module,symbols,undefined,{},undefined,(...args)=>diagnostics.push(args));
        if(diagnostics.length) throw new Error(JSON.stringify(diagnostics));
    });
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, rounds, rows }, null, 2));
