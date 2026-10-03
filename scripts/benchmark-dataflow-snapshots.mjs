// Run: node scripts/benchmark-dataflow-snapshots.mjs [--baseline=<commit>] [--rounds=15]
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.slice(9) ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw Error('rounds must be 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-dataflow-snapshots-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/straightLineValues\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/straightLineValues.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics')}));}}]:[];
 const result = await build({ plugins, stdin: { contents: "export { createConditionalActivityTracker } from './src/analyzer/conditional/conditionalCompilation'; export { checkAssignmentTypes } from './src/analyzer/diagnostics/rules/assignments'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(bundle, result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; } rmdirSync(scratch); }
const rows=[];
for(const count of [10,100,1000,3000])for(const mode of ['cold-default','cold-conditional','warm-default','scalar-control']){
 const options=mode==='cold-conditional'?'#If VBA7 Then\nOption Base 1\n#Else\nOption Base 0\n#End If\n':'';
 const constants=Array.from({length:count},(_,i)=>'Const K'+i+' = 1\n').join('');
 const source=options+constants+'Sub P()\nDim v As Variant\nDim n As Long\nv = Array("bad", "1")\n'+(mode==='scalar-control'?'n = 1\n':'n = v('+(mode==='cold-conditional'?1:0)+')\n').repeat(1000)+'End Sub';
 const environment=mode==='cold-conditional'?{compilerConstants:{VBA7:true}}:undefined;let previousRoot, previousBody, serial = 0;
 function prepare(){const sampleSource=mode.startsWith('cold-')?source+"\n' sample "+serial++:source;const sampleMod=api.parseModule(sampleSource);const body=sampleMod.members.find(m=>m.kind==='Procedure').body;if(mode.startsWith('cold-')&&body===previousBody)throw Error('Reused parse body');previousBody=body;const symbols=api.buildModuleSymbols('M','standard',sampleSource,{parsedModule:sampleMod,conditionalCompilation:environment});if(symbols.root===previousRoot)throw Error('Reused root');previousRoot=symbols.root;const activity=environment?api.createConditionalActivityTracker(sampleMod,environment):undefined;return()=>{let hits=0;api.checkAssignmentTypes(sampleSource,sampleMod,symbols,undefined,{},activity,()=>hits++);if(hits!==(mode==='scalar-control'?0:1000))throw Error('Wrong '+mode+' hits '+hits);};}
 for(let i=0;i<3;i++)prepare()();const samples=[];for(let i=0;i<rounds;i++){const run=prepare(),start=performance.now();run();samples.push(performance.now()-start);}samples.sort((a,b)=>a-b);
 rows.push({name:count+'-'+mode,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
