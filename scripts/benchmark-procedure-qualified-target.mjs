import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-procedure-qualified-target.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(),'xlide-qualified-target-'));
const path = join(scratch,'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(build){build.onLoad({filter:/[\\/]refactor[\\/]procedureCallBinding\.ts$/},args=>({contents:execFileSync('git',['show',`${baseline}:src/analyzer/refactor/procedureCallBinding.ts`],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}] : [];
 const built = await build({stdin:{contents:"export {introduceParameter} from './src/analyzer/refactor/introduceParameter'; export {parseModule} from './src/analyzer/parser/parseModule'; export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 writeFileSync(path,built.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally { if(existsSync(path))unlinkSync(path);rmdirSync(scratch); }
const rows=[];
for(const unrelated of [0,1000])for(const calls of [1,1000])for(const mode of ['qualified','bare','private']){
 const samples=[];
 for(let round=-3;round<9;round++){
  const prefix=Array.from({length:unrelated},(_,i)=>'Sub Earlier'+i+'()\nEnd Sub\n').join('');
  const access=mode==='private'?'Private':'Public';
  const source=prefix+access+' Sub Report()\nDim limit As Long\nlimit = 3\nDebug.Print limit\nEnd Sub\n';
  const call=(mode==='bare'?'':'Module1.')+'Report';
  const caller='Sub Caller()\n'+Array(calls).fill(call).join('\n')+'\nEnd Sub\n'+"' round "+round;
  const input={source,offset:source.indexOf('limit'),moduleName:'Module1',otherModuleSources:{Caller:caller}};
  api.parseModule(source);api.parseModule(caller);
  const begin=performance.now();const results=Array.from({length:5},()=>api.introduceParameter(input));const elapsed=(performance.now()-begin)/5;
  for(const result of results){
   assert.equal(result.ok,true);assert.equal(result.title,"Introduce 'limit' as a parameter");
   assert.equal(api.applyVbaTextEdits(source,result.edits),prefix+access+' Sub Report(ByVal limit As Long)\nDebug.Print limit\nEnd Sub\n');
   assert.equal(result.otherModules?.length??0,mode==='private'?0:1);
   if(mode!=='private'){assert.equal(result.otherModules[0].moduleName,'Caller');assert.equal(api.applyVbaTextEdits(caller,result.otherModules[0].edits),'Sub Caller()\n'+Array(calls).fill(call+' 3').join('\n')+'\nEnd Sub\n'+"' round "+round);}
  }
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({unrelated,calls,mode,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeResultsChecked:true});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete warmed Introduce Parameter calls. AST/source construction outside clock; query project construction, binding and edits included. Complete result titles/module edits/applied output independently checked. No cold-parser, heap-byte or editor-latency claim.',rows},null,2));
