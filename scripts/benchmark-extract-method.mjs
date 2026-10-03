// Run in the target checkout: node scripts/benchmark-extract-method.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-extract-method-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { extractMethod } from './src/analyzer/refactor/extractMethod';
        export { findIdentifierOccurrences } from './src/vbaSourceScan';
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
    samples.sort((a,b)=>a-b);
    rows.push({ name, medianMs: +samples[Math.floor(samples.length/2)].toFixed(3),
        p95Ms: +samples[Math.ceil(samples.length*.95)-1].toFixed(3) });
}
for(const [locals, statements] of [[100,1000],[300,3000],[1000,2000]]) {
    const prefix='Option Explicit\nPublic Sub Main()\n'+Array.from({length:locals},(_,i)=>'Dim v'+i+' As Long\n').join('');
    const selected=Array.from({length:5},(_,i)=>'v'+i+' = v'+i+' + 1\n').join('');
    const source=prefix+selected+Array.from({length:statements},(_,i)=>'Debug.Print v'+(i%locals)+'\n').join('')+'End Sub\n';
    const input={source,span:{start:prefix.length,end:prefix.length+selected.length-1},name:'Extracted'};
    measure(locals+'-locals/'+statements+'-statements',()=>{const result=api.extractMethod(input);if(!result.ok)throw new Error(result.reason);});
}
const referenceSource=Array.from({length:10000},(_,i)=>'v'+(i%100)+' = v'+((i+1)%100)+' + 1\n').join('');
measure('single-name-scan/10000-statements',()=>{if(api.findIdentifierOccurrences(referenceSource,'v0').length!==200)throw new Error('Missing occurrences');});
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, rounds, rows }, null, 2));
