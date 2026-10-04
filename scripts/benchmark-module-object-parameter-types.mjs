import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {dirname} from 'node:path';
import {performance} from 'node:perf_hooks';
import {writeFileSync,unlinkSync,mkdtempSync,rmdirSync} from 'node:fs';
import {tmpdir,cpus} from 'node:os';
import {join} from 'node:path';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
const baseline=process.argv.find(s=>s.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-module-parameters-')),file=join(scratch,'api.cjs');let a;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]rules[\\/]objectState\.ts$/},p=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/objectState.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(p.path)}));}}]:[];
 const b=await build({plugins,stdin:{contents:"export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,b.outputFiles[0].contents);a=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file)}finally{rmdirSync(scratch)}}
const rows=[],workCounts=[];
for(const distinct of [false,true])for(const n of [1,100,1000]){let names=0;const classes=Array.from({length:n+1},(_,i)=>({get name(){names++;return 'Class'+i},moduleName:'Class'+i,kind:'class',exhaustive:true,members:[]}));const source=['Option Explicit',...Array.from({length:n},(_,i)=>'Sub P'+i+'(ByVal actor As Class'+(distinct?i:0)+')\nEnd Sub'),''].join('\n');const run=()=>{const errors=[];const diagnostics=a.analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});return {diagnostics,errors}};const expected={diagnostics:[],errors:[]};assert.deepEqual(run(),expected);workCounts.push({distinct,n,names});if(!baseline)assert.ok(names<=n*8+30);for(const c of classes){Object.defineProperty(c,'name',{value:c.moduleName,writable:true,enumerable:true,configurable:true});}const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=run(),elapsed=performance.now()-begin;assert.deepEqual(actual,expected);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({distinct,n,scope:'complete-module-diagnostics',medianMs:+samples[4].toFixed(5)});}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,warmups:3,rounds:9,workCounts,rows,scope:'Warm complete-module setup included. Independently empty diagnostics/errors for repeated and distinct object parameters. Plain metadata timed, name getters untimed only. Equality outside clock. No editor/cold/heap claim.'},null,2));
