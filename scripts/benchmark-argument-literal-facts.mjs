// Run in the target checkout: node scripts/benchmark-argument-literal-facts.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-argument-literal-facts-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
    const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/rules[\\/]argumentTypes\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/argumentTypes.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
    const result = await build({ plugins, stdin: { contents:
        "export { checkArgumentTypes } from './src/analyzer/diagnostics/rules/argumentTypes'; export { forEachStatementWithHeaders } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';",
        resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(bundle, result.outputFiles[0].contents);
    api = createRequire(import.meta.url)(bundle);
} finally {
    try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    rmdirSync(scratch);
}
const rows=[];
function measure(name,prepare){
 for(let i=0;i<3;i++)prepare()();
 const samples=[];
 for(let i=0;i<rounds;i++){const run=prepare(),start=performance.now();run();samples.push(performance.now()-start);}
 samples.sort((a,b)=>a-b);rows.push({name,medianMs:+samples[Math.floor(samples.length/2)].toFixed(3),p95Ms:+samples[Math.ceil(samples.length*.95)-1].toFixed(3)});
}
for(const count of [100,1000,3000]) {
 for(const mode of ['no-calls','named-argument','literal-argument']) {
  const source='Sub TakeLong(ByVal n As Long)\nEnd Sub\nSub P()\n'+Array.from({length:25},(_,i)=>'Dim v'+i+' As Long').join('\n')+'\nDim value As Variant\nvalue = 300\n'+Array.from({length:count},(_,i)=>'v'+(i%25)+' = '+(i%10)).join('\n')+'\n'+(mode==='named-argument'?'TakeLong value\n':mode==='literal-argument'?'TakeLong "bad"\n':'')+'End Sub';
  const mod=api.parseModule(source);let previousRoot;
  measure(count+'-'+mode,()=>{const symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod});if(symbols.root===previousRoot)throw Error('Reused binding context');previousRoot=symbols.root;return ()=>{let hits=0;const factory=api.checkArgumentTypes(source,symbols,undefined,undefined,{},()=>hits++);for(const proc of mod.members){if(proc.kind==='Procedure'){const visit=factory(proc);if(visit)api.forEachStatementWithHeaders(source,proc.body,visit);}}if(hits!==(mode==='literal-argument'?1:0))throw Error('Wrong findings: '+hits);};});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
