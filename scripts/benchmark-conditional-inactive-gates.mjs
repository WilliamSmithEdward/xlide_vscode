import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {mkdtempSync,writeFileSync,existsSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {dirname,join} from 'node:path';
import {tmpdir,cpus} from 'node:os';
import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-conditional-inactive-gates.mjs [--baseline=COMMIT]
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const scratch=mkdtempSync(join(tmpdir(),'xlide-conditional-gates-')),path=join(scratch,'bundle.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]conditional[\\/]conditionalCompilation\.ts$/},a=>({contents:execFileSync('git',['show',baseline+':src/analyzer/conditional/conditionalCompilation.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(a.path)}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export {createConditionalActivityTracker} from './src/analyzer/conditional/conditionalCompilation';export {parseModule} from './src/analyzer/parser/parseModule';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,result.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally {if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
const rows=[],expression='Not '.repeat(100)+'0';
for(const count of [1,1000])for(const kind of ['settled-elseif','inactive-parent','uncertain-control']){
 const source='Option Explicit\n'+(kind==='inactive-parent'?'#If False Then\n'+Array(count).fill('#If '+expression+' Then\n').join('')+'Public SkippedValue As Long\n'+'#End If\n'.repeat(count)+'#Else\nPublic TakenValue As Long\n#End If\n':'#If '+(kind==='settled-elseif'?'True':'Missing')+' Then\nPublic TakenValue As Long\n'+Array(count).fill('#ElseIf '+expression+' Then\n').join('')+'Public SkippedValue As Long\n#Else\nPublic LastValue As Long\n#End If\n');
 const module=api.parseModule(source),span=name=>{const start=source.indexOf(name);return {start,end:start+name.length};};
 for(const scope of ['tracker-construction','complete-module-diagnostics']){const samples=[];for(let round=-3;round<9;round++){
  const failures=[],start=performance.now(),result=scope==='tracker-construction'?api.createConditionalActivityTracker(module):api.analyzeModule(source,{onInternalError:e=>failures.push(String(e))}),elapsed=performance.now()-start;
  assert.deepEqual(failures,[]);if(scope==='complete-module-diagnostics')assert.deepEqual(result,[]);else {assert.equal(result.activityForSpan(span('TakenValue')),kind==='uncertain-control'?'unknown':'active');assert.equal(result.activityForSpan(span('SkippedValue')),'inactive');if(kind!=='inactive-parent')assert.equal(result.activityForSpan(span('LastValue')),kind==='uncertain-control'?'unknown':'inactive');}
  if(round>=0)samples.push(elapsed);
 }samples.sort((a,b)=>a-b);rows.push({count,kind,scope,medianMs:+samples[4].toFixed(5),completeOutputsChecked:true});}
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,rows,scope:'100-Not expression per arm. Source creation/AST warming outside clock. Complete tracker construction or complete analyzeModule diagnostics inside; independent activity/empty-diagnostic/failure checks every call. Uncertain-arm control requires all condition evaluations. No cold-parser, renderer or retained-heap claim.'},null,2));
