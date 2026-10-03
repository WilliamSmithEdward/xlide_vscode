// Run in the target checkout: node scripts/benchmark-redim-option-base.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-redim-option-base-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
    const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/rules[\\/]arrays\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/arrays.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
    const result = await build({ plugins, stdin: { contents:
        "export { checkRedimPreserveDimensions } from './src/analyzer/diagnostics/rules/arrays'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';",
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
for(const count of [100,1000,3000]) {
 for(const option of ['', 'Option Base 1\n']) {
  const source=option+Array.from({length:count},(_,i)=>'Sub P'+i+'()\nDim a() As Long\nReDim a(1 To 3)\nReDim Preserve a(4)\nEnd Sub').join('\n');
  const mod=api.parseModule(source);
  measure(count+'-'+(option?'base1':'default'),()=>{let hits=0;api.checkRedimPreserveDimensions(source,mod,undefined,()=>hits++);if(hits!==(option?0:count))throw new Error('Wrong findings: '+hits);});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
