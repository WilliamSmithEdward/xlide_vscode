// Run in the target checkout: node scripts/benchmark-canonical-casing.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-canonical-casing-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { resolveCanonicalCaseEdit, resolveCanonicalCaseEdits } from './src/analyzer/completion/canonicalCasing';
        
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
for (const [procedures, statements] of [[10,10],[100,10],[1,1000]]) {
    const source = Array.from({length:procedures},(_,p)=>'Sub Main'+p+'()\nDim Value As Long\n'+Array.from({length:statements},()=> 'value = value + 1\n').join('')+'End Sub\n').join('\n');
    const ctx = {identifier:{includeGlobals:false,includeRuntime:false}};
    measure(procedures+'-procedures/'+statements+'-statements',()=>{
        const result=api.resolveCanonicalCaseEdits(source,{start:0,end:source.length},ctx);
        if(result.length!==procedures*statements*2)throw new Error('Unexpected edits: '+result.length);
    });
}
const singleSource = Array.from({length:100},(_,p)=>'Sub Main'+p+'()\nDim Value As Long\nvalue = value + 1\nEnd Sub\n').join('\n');
const singleOffset = singleSource.lastIndexOf('value') + 5;
measure('single-position/100-procedures',()=>{
    const result=api.resolveCanonicalCaseEdit(singleSource,singleOffset,{identifier:{includeGlobals:false,includeRuntime:false}});
    if(result?.text!=='Value')throw new Error('Unexpected single edit');
});
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
