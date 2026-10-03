// Run in the target checkout: node scripts/benchmark-split-case-fold.mjs [--rounds=15]
import { build } from 'esbuild';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-split-case-fold-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
    const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/rules[\\/]arrays\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/arrays.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
    const result = await build({ plugins, stdin: { contents:
        "export { arrayValueShape } from './src/analyzer/diagnostics/rules/arrays'; export { tokenizeCached } from './src/analyzer/lexer/tokenize';",
        resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
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
for(const count of [1000,10000,30000]) {
 for(const mode of ['text','binary','text-limit-one']) {
  const text='AaX'.repeat(count),limit=mode==='text-limit-one'?1:-1;
  const tokens=api.tokenizeCached('Split("'+text+'", "'+(mode==='binary'?'X':'x')+'", '+limit+', '+(mode==='binary'?'vbBinaryCompare':'vbTextCompare')+')');
  measure(count+'-'+mode,()=>{const shape=api.arrayValueShape(tokens,'v',0);if(shape?.values?.length!==(limit===1?1:count+1))throw new Error('Wrong shape');});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
