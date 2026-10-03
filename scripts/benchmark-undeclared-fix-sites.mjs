// Run in the target checkout: node scripts/benchmark-undeclared-fix-sites.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-undeclared-sites-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
    const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/rules[\\/]undeclared\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/undeclared.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
    const result = await build({ plugins, stdin: { contents:
        "export { checkUndeclaredVariables } from './src/analyzer/diagnostics/rules/undeclared'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';",
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
 for(const mode of ['missing','declared']) {
  const source='Option Explicit\nSub P()\n'+Array.from({length:count},(_,i)=>'Dim D'+i+' As Object').join('\n')+'\n'+Array.from({length:count},(_,i)=>'Set '+(mode==='missing'?'Missing':'D')+i+' = Nothing').join('\n')+'\nEnd Sub';
  const mod=api.parseModule(source), symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod}), known=new Set();
  measure(count+'-'+mode,()=>{let fixes=0;api.checkUndeclaredVariables(source,mod,symbols,undefined,known,undefined,undefined,undefined,undefined,'standard',undefined,undefined,undefined,(_key,_message,_span,data)=>{if(data?.declareVariable)fixes++;});if(fixes!==(mode==='missing'?count:0))throw new Error('Wrong fix count: '+fixes);});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
