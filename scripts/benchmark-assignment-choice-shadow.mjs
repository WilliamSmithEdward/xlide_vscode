import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??15);
if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3 to 100');
const dir=mkdtempSync(join(tmpdir(),'xlide-choice-shadow-')),bundle=join(dir,'analyzer.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/rules[\\/]assignments\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/assignments.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export { checkAssignmentTypes } from './src/analyzer/diagnostics/rules/assignments'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(bundle,result.outputFiles[0].contents);api=createRequire(import.meta.url)(bundle);
}finally{try{unlinkSync(bundle);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(dir);}
const rows=[];
for(const count of [10,100,1000,3000])for(const mode of ['choose','switch','iif','qualified','scalar']){
 const expr={choose:'Choose(0, 1)',switch:'Switch(False, 1)',iif:'IIf(False, 1, Null)',qualified:'VBA.Choose(0, 1)',scalar:'1'}[mode];
 const source=Array.from({length:count},(_,i)=>'Dim k'+i+' As Long\n').join('')+'Sub P()\nDim n As Long\n'+('n = '+expr+'\n').repeat(1000)+'End Sub';
 const mod=api.parseModule(source);let previous;
 function prepare(){const symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod});if(symbols.root===previous)throw Error('Reused root');previous=symbols.root;return()=>{let hits=0;api.checkAssignmentTypes(source,mod,symbols,undefined,{},undefined,()=>hits++);if(hits!==(mode==='scalar'?0:1000))throw Error('Wrong hits '+mode+': '+hits);};}
 for(let i=0;i<3;i++)prepare()();const samples=[];
 for(let i=0;i<rounds;i++){const run=prepare(),start=performance.now();run();samples.push(performance.now()-start);}samples.sort((a,b)=>a-b);
 rows.push({name:count+'-'+mode,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
