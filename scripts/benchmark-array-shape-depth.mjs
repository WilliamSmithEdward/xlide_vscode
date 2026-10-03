// Run in the target checkout: node scripts/benchmark-array-shape-depth.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-array-shape-depth-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
    const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/rules[\\/]arrays\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/arrays.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
    const result = await build({ plugins, stdin: { contents:
        "export { arrayValueShape } from './src/analyzer/diagnostics/rules/arrays'; export { rawExpressionTokens } from './src/analyzer/diagnostics/walker';",
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
for(const depth of [10,100,1000,5000]) {
 for(const kind of ['Array','Filter']) {
  const source=kind==='Array'?'Array('.repeat(depth)+'1'+')'.repeat(depth):'Filter('.repeat(depth)+'Array("a", "b")'+', "a")'.repeat(depth);
  const tokens=api.rawExpressionTokens(source);let errors=0,levels=0;
  measure(depth+'-'+kind,()=>{try{let shape=api.arrayValueShape(tokens,'a',0,undefined,'binary');levels=0;while(shape){levels++;shape=shape.elements?.[0];}}catch(error){if(!(error instanceof RangeError))throw error;errors++;}});
  rows.at(-1).stackErrors=errors;rows.at(-1).returnedLevels=levels;
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
