import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {dirname} from 'node:path';
import {performance} from 'node:perf_hooks';
import {writeFileSync,unlinkSync,mkdtempSync,rmdirSync} from 'node:fs';
import {tmpdir,cpus} from 'node:os';
import {join} from 'node:path';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
const baseline=process.argv.find(s=>s.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-assignment-queries-')),file=join(scratch,'api.cjs');let a;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]rules[\\/]assignments\.ts$/},p=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/assignments.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(p.path)}));}}]:[];
 const b=await build({plugins,stdin:{contents:"export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,b.outputFiles[0].contents);a=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file)}finally{rmdirSync(scratch)}}
const rows=[],workCounts=[];
for(const mode of ['variable','held','property'])for(const n of [1,100,1000]){
 let reads=0,nameReads=0;const names=Array.from({length:n},(_,i)=>i===n-1?'Class1':'Other'+i),list=new Proxy(names,{get(t,k,r){if(typeof k==='string'&&/^(0|[1-9]\d*)$/.test(k))reads++;return Reflect.get(t,k,r)}});
 const surface=(name,impl=[])=>({get name(){nameReads++;return name},moduleName:name,kind:'class',exhaustive:true,members:[],implements:impl}),classes=[surface('Class1'),surface('Class2',list),...Array.from({length:n},(_,i)=>surface('Extra'+i))];
 if(mode==='property'){const h=surface('Holder');h.members=[{name:'Item',kind:'property',moduleName:'Holder',returns:'Class1',writeType:'Class1',writable:true,setAccessor:true}];classes.push(h);}
 const source=['Option Explicit','Sub Go(ByVal actual As '+(mode==='held'?'Object':'Class2')+', ByVal target As Class1'+(mode==='property'?', ByVal holder As Holder':'')+')',...(mode==='held'?['Set actual = New Class2']:[]),...Array.from({length:n},()=>mode==='property'?'Set holder.Item = actual':'Set target = actual'),'End Sub',''].join('\n');
 const run=()=>{const errors=[];const diagnostics=a.analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});return {diagnostics,errors}};
 const expected={diagnostics:[],errors:[]};assert.deepEqual(run(),expected);workCounts.push({mode,n,nameReads,reads});if(!baseline){assert.ok(reads<=n*3);assert.ok(nameReads<=n*25+100);}
 // All name getters and index-counting proxies are removed before timing.
 for(const c of classes){const name=c.moduleName;Object.defineProperty(c,'name',{value:name,enumerable:true,configurable:true,writable:true});}classes[1].implements=names;
 const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=run(),elapsed=performance.now()-begin;assert.deepEqual(actual,expected);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({mode,n,scope:'complete-module-diagnostics',medianMs:+samples[4].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,workCounts,rows,scope:'Complete-module diagnostics and internal errors independently empty for compatible assignments. Warm calls include all analyzer setup. Plain metadata timed; property getters and array-index proxy only untimed. Equality outside clock. No editor/cold/heap claim.'},null,2));
