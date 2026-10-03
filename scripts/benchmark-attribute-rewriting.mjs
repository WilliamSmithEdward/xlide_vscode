// Run in the target checkout: node scripts/benchmark-attribute-rewriting.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-attribute-rewriting-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { applyAttributeAnnotations } from './src/analyzer/annotations/attributeRewriter';
        export { readAttributeAnnotations } from './src/analyzer/annotations/attributeAnnotations';
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
for(const count of [100,1000,3000]) {
 const source='Attribute VB_Name = "Module"\n'+Array.from({length:count},(_,i)=> "'@Description(\"Doc "+i+"\")\n'@ExcelHotkey(\"A\")\nSub P"+i+'()\n'+Array.from({length:10},()=> 'Debug.Print 1').join('\n')+'\nEnd Sub\n').join('');
 const annotations=api.readAttributeAnnotations(source);
 if(annotations.annotations.length!==count*2||annotations.problems.length)throw new Error('Invalid annotation fixture');
 measure(count+'-procedures/insert',()=>{const out=api.applyAttributeAnnotations(source,annotations);if(out.changes.length!==count*2||out.skipped.length)throw new Error('Wrong rewrite result');});
 const rewritten=api.applyAttributeAnnotations(source,annotations).text;
 measure(count+'-procedures/unchanged',()=>{const out=api.applyAttributeAnnotations(rewritten,annotations);if(out.changes.length||out.skipped.length||out.text!==rewritten)throw new Error('Wrong no-op result');});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
