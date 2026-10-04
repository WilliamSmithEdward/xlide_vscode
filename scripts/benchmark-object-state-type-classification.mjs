import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {dirname,join} from 'node:path';
import {performance} from 'node:perf_hooks';
import {writeFileSync,unlinkSync,mkdtempSync,rmdirSync} from 'node:fs';
import {tmpdir,cpus} from 'node:os';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
const baseline=process.argv.find(s=>s.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-object-classification-')),file=join(scratch,'api.cjs');let a;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]rules[\\/]objectState\.ts$/},p=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/objectState.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(p.path)}));}}]:[];
 const b=await build({plugins,stdin:{contents:"export {checkObjectVariableNotSet} from './src/analyzer/diagnostics/rules/objectState';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,b.outputFiles[0].contents);a=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file)}finally{rmdirSync(scratch)}}
const rows=[],workCounts=[];
for(const family of ['functions','module-variable','locals','arrays'])for(const distinct of family==='module-variable'?[false]:[false,true])for(const n of [1,100,1000]){
 let names=0;const classes=Array.from({length:n+1},(_,i)=>({get name(){names++;return 'Class'+i},moduleName:'Class'+i,kind:'class',exhaustive:true,members:[]}));const lines=['Option Explicit'];if(family==='module-variable')lines.push('Private actor As Class0');for(let i=0;i<n;i++){const type='Class'+(distinct?i:0);if(family==='functions'){lines.push('Function P'+i+'() As '+type,'End Function');}else{lines.push('Sub P'+i+'()');if(family==='locals')lines.push('Dim actor As '+type);if(family==='arrays')lines.push('Dim actor(0) As '+type);lines.push('End Sub');}}const source=lines.join('\n')+'\n',mod=a.parseModule(source),symbols=a.buildModuleSymbols('M','standard',source,{parsedModule:mod});
 const rule=()=>{const out=[];a.checkObjectVariableNotSet(source,mod,symbols,{projectClassMembers:classes},undefined,(...v)=>out.push(v));return out;};
 const full=()=>{const errors=[],diagnostics=a.analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});return {diagnostics,errors}};
 names=0;assert.deepEqual(rule(),[]);workCounts.push({family,distinct,n,scope:'public-rule',names});if(!baseline)assert.ok(names<=n*8+40);
 names=0;const expected=full();assert.deepEqual(expected.errors,[]);assert.equal(expected.diagnostics.length,family==='module-variable'?1:n);let from=0;for(let i=0;i<expected.diagnostics.length;i++){const d=expected.diagnostics[i],fn=family==='functions',needle=fn?'Function P'+i:family==='module-variable'?'Private actor':'Dim actor',at=source.indexOf(needle,from)+(fn?9:family==='module-variable'?8:4),name=fn?'P'+i:'actor';from=at+name.length;assert.equal(d.code,fn?'missing-return-assignment':'unused-variable');assert.equal(d.message,fn?"Function '"+name+"' has no return assignment; VBA will return the default value. Assign to '"+name+"' before exit if a value is intended.":(family==='module-variable'?'Module-level':'Local')+" variable 'actor' is declared but never used.");assert.deepEqual(d.span,{start:at,end:at+name.length});}workCounts.push({family,distinct,n,scope:'complete-module-diagnostics',names});
 for(const c of classes)Object.defineProperty(c,'name',{value:c.moduleName,writable:true,enumerable:true,configurable:true});
 for(const [scope,run,want] of [['public-rule',rule,[]],['complete-module-diagnostics',full,expected]]){const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=run(),elapsed=performance.now()-begin;assert.deepEqual(actual,want);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({family,distinct,n,scope,medianMs:+samples[4].toFixed(5)});}
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,warmups:3,rounds:9,workCounts,rows,scope:'Public-rule AST/symbols outside clock; full-module setup included. Getter counters untimed, plain properties timed. Independently expected empty rule findings and exact full warning messages/spans/counts; errors empty. Equality outside clock; no editor/cold/heap claim.'},null,2));
