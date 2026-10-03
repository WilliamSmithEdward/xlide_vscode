// Run in the target checkout: node scripts/benchmark-formatter.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-formatter-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { formatVbaModule, tokenStreamDifference } from './src/analyzer/format/formatModule';
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
const body = (n, continuation=false) => Array.from({length:n},(_,i)=>'total= x + '+i+(continuation?' + _\n  1':'')+'\n').join('');
const proc = (i,n,continuation=false) => 'public sub P'+i+'(byval x as long)\ndim total as long\n'+body(n,continuation)+'debug.print total\nend sub\n';
for(const [name,source] of [
    ['many-100',Array.from({length:100},(_,i)=>proc(i,12)).join('\n')],
    ['many-2000',Array.from({length:2000},(_,i)=>proc(i,2)).join('\n')],
    ['body-10000',proc(0,10000)],
    ['continuations-2000',proc(0,2000,true)],
]) {
    measure(name+'/format',()=>{const result=api.formatVbaModule(source,{tabSize:4,insertSpaces:true});if(result.text===undefined)throw new Error(result.refusal);});
    measure(name+'/validate',()=>{const difference=api.tokenStreamDifference(source,source);if(difference)throw new Error(difference);});
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, rounds, rows }, null, 2));
