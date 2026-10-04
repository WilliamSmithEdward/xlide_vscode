import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {mkdtempSync,writeFileSync,existsSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {dirname,join} from 'node:path';
import {tmpdir,cpus} from 'node:os';
import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-conditional-compiler-lookup.mjs [--baseline=COMMIT]
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const scratch=mkdtempSync(join(tmpdir(),'xlide-conditional-compiler-')),path=join(scratch,'bundle.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]conditional[\\/]conditionalCompilation\.ts$/},a=>({contents:execFileSync('git',['show',baseline+':src/analyzer/conditional/conditionalCompilation.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(a.path)}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export {createConditionalActivityTracker,indexConditionalCompilation} from './src/analyzer/conditional/conditionalCompilation';export {parseModule} from './src/analyzer/parser/parseModule';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,result.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally {if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
const rows=[];
for(const count of [1,1000])for(const additionalCompilerFlags of [0,1000]){
 const source='Option Explicit\n'+Array.from({length:count},(_,i)=>'#Const LOCAL'+i+' = '+(i===0?'VBA7':'LOCAL'+(i-1)+' + 1')+'\n').join('')+'#If LOCAL'+(count-1)+' = '+count+' Then\nPublic TakenValue As Long\n#End If\n';
 const module=api.parseModule(source),start=source.indexOf('Public Taken'),span={start,end:start+6},env={compilerConstants:Object.fromEntries(Array.from({length:additionalCompilerFlags},(_,i)=>['HOSTFLAG'+i,i+1]))};
 for(const scope of ['tracker-construction','constant-index','complete-module-diagnostics']){const samples=[];for(let round=-3;round<9;round++){
  const failures=[],begin=performance.now(),actual=scope==='tracker-construction'?api.createConditionalActivityTracker(module,env):scope==='constant-index'?api.indexConditionalCompilation(module,env):api.analyzeModule(source,{conditionalCompilation:env,onInternalError:e=>failures.push(String(e))}),elapsed=performance.now()-begin;
  assert.deepEqual(failures,[]);if(scope==='tracker-construction')assert.equal(actual.activityForSpan(span),'active');else if(scope==='constant-index')assert.deepEqual(actual.constants.map(c=>[c.name,c.value]),Array.from({length:count},(_,i)=>['LOCAL'+i,i+1]));else assert.deepEqual(actual,[]);
  if(round>=0)samples.push(elapsed);
 }samples.sort((a,b)=>a-b);rows.push({count,additionalCompilerFlags,scope,medianMs:+samples[4].toFixed(5),completeOutputsChecked:true});}
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,rows,scope:'Source/environment creation and AST warming outside clock. Complete tracker construction, constant indexing or analyzeModule diagnostics inside. Independent expected constant sequences, activity and complete diagnostics/failures checked every call. Zero additional flags uses the normal default compiler constants; 1000 extra compiler flags is a custom-environment stress workload, not a typical VBA project-constant count. No cold-parser, retained-heap or renderer claim.'},null,2));
