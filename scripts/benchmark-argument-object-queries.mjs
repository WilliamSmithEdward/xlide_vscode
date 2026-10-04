import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {dirname,relative} from 'node:path';
import {performance} from 'node:perf_hooks';
import {writeFileSync,unlinkSync,mkdtempSync,rmdirSync} from 'node:fs';
import {tmpdir,cpus} from 'node:os';
import {join} from 'node:path';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
const baseline=process.argv.find(s=>s.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-argument-queries-')),file=join(scratch,'api.cjs');let a;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/](?:rules[\\/])?(?:argumentTypes|typeInference)\.ts$/},p=>({contents:execFileSync('git',['show',baseline+':'+relative(process.cwd(),p.path).replace(/\\/g,'/')],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(p.path)}));}}]:[];
 const b=await build({plugins,stdin:{contents:"export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,b.outputFiles[0].contents);a=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file)}finally{rmdirSync(scratch)}}
const rows=[],workCounts=[];
for(const mode of ['declared','held','new'])for(const n of [1,100,1000]){let reads=0,names=0;const implemented=new Proxy(Array.from({length:n},(_,i)=>i===n-1?'Class1':'Other'+i),{get(t,k,r){if(typeof k==='string'&&/^(0|[1-9]\d*)$/.test(k)){reads++;}return Reflect.get(t,k,r);}});const surface=(name,impl=[])=>({get name(){names++;return name;},moduleName:name,kind:'class',exhaustive:true,members:[],implements:impl});const classes=[surface('Class1'),surface('Class2',implemented),...Array.from({length:n},(_,i)=>surface('Extra'+i))];const source=['Option Explicit','Sub Take(ByVal target As Class1)','End Sub','Sub Go(ByVal actual As '+(mode==='held'?'Object':'Class2')+')',...(mode==='held'?['Set actual = New Class2']:[]),...Array.from({length:n},()=>mode==='new'?'Take New Class2':'Take actual'),'End Sub',''].join('\n');const run=()=>{const errors=[];const diagnostics=a.analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});return {diagnostics,errors}};const expected={diagnostics:[],errors:[]};assert.deepEqual(run(),expected);workCounts.push({mode,n,names,reads});if(!baseline){assert.ok(names<=n*20+100);assert.ok(reads<=n*3);}for(const c of classes){Object.defineProperty(c,'name',{value:c.moduleName,writable:true,enumerable:true,configurable:true});}classes[1].implements=Array.from({length:n},(_,i)=>i===n-1?'Class1':'Other'+i);const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=run(),elapsed=performance.now()-begin;assert.deepEqual(actual,expected);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({mode,n,scope:'complete-module-diagnostics',medianMs:+samples[4].toFixed(5)});}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,warmups:3,rounds:9,workCounts,rows,scope:'Warm complete-module setup included; compatible calls independently no diagnostics/errors. Plain metadata timed, getters and array proxies untimed only. Equality outside clock. No editor/cold/heap claim.'},null,2));
