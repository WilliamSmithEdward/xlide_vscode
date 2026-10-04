import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {dirname,relative} from 'node:path';
import {performance} from 'node:perf_hooks';
import {writeFileSync,unlinkSync,mkdtempSync,rmdirSync} from 'node:fs';
import {tmpdir,cpus} from 'node:os';
import {join} from 'node:path';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
const baseline=process.argv.find(s=>s.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-default-member-queries-')),file=join(scratch,'api.cjs');let a;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/](?:rules[\\/])?(?:objectValues|statementForms|typeInference)\.ts$/},p=>({contents:execFileSync('git',['show',baseline+':'+relative(process.cwd(),p.path).replace(/\\/g,'/')],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(p.path)}));}}]:[];
 const b=await build({plugins,stdin:{contents:"export {checkObjectDefaultValues} from './src/analyzer/diagnostics/rules/objectValues';export {checkStatementForms} from './src/analyzer/diagnostics/rules/statementForms';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {forEachStatement} from './src/analyzer/parser/statementWalk';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,b.outputFiles[0].contents);a=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file)}finally{rmdirSync(scratch)}}
const rows=[],workCounts=[];
for(const distinct of [false,true])for(const n of [1,100,1000]){
 let names=0;const classes=Array.from({length:n+1},(_,i)=>({get name(){names++;return 'Class'+i},moduleName:'Class'+i,kind:'class',exhaustive:true,members:[]}));const source=['Option Explicit',...Array.from({length:n},(_,i)=>'Sub P'+i+'(ByVal actor As Class'+(distinct?i:0)+')\nDebug.Print actor\nEnd Sub'),''].join('\n'),mod=a.parseModule(source),symbols=a.buildModuleSymbols('M','standard',source,{parsedModule:mod});
 let from=0;const expected=Array.from({length:n},(_,i)=>{const start=source.indexOf('Debug.Print actor',from)+12;from=start+5;return ['objectDefaultValue',"'actor' is a Class"+(distinct?i:0)+", which has no default member, so it has no value to read here. This will raise Run-time error '438': Object doesn't support this property or method, or '91' while it is Nothing.",{start,end:start+5}];});
 const values=()=>{const out=[],factory=a.checkObjectDefaultValues(source,symbols,{projectClassMembers:classes},(...v)=>out.push(v));for(const p of mod.members)if(p.kind==='Procedure'){const visit=factory(p);if(visit)a.forEachStatement(p.body,visit);}return out;};const statements=()=>{const out=[];a.checkStatementForms(source,mod,symbols,undefined,undefined,(...v)=>out.push(v),{projectClassMembers:classes});return out;};const full=()=>{const errors=[],diagnostics=a.analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});return {diagnostics,errors}};
 for(const [scope,run,want] of [['values-rule',values,expected],['statements-rule',statements,[]]]){names=0;assert.deepEqual(run(),want);workCounts.push({distinct,n,scope,names});if(!baseline)assert.ok(names<=n*15+100);}
 names=0;const fullExpected=full();assert.deepEqual(fullExpected.errors,[]);assert.equal(fullExpected.diagnostics.length,n);for(let i=0;i<n;i++){assert.equal(fullExpected.diagnostics[i].code,'object-default-value');assert.equal(fullExpected.diagnostics[i].message,expected[i][1]);assert.deepEqual(fullExpected.diagnostics[i].span,expected[i][2]);}workCounts.push({distinct,n,scope:'complete-module-diagnostics',names});
 // Counter getters are exclusively untimed.
 for(const c of classes)Object.defineProperty(c,'name',{value:c.moduleName,writable:true,enumerable:true,configurable:true});
 for(const [scope,run,want] of [['values-rule',values,expected],['statements-rule',statements,[]],['complete-module-diagnostics',full,fullExpected]]){const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=run(),elapsed=performance.now()-begin;assert.deepEqual(actual,want);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({distinct,n,scope,medianMs:+samples[4].toFixed(5)});}
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,warmups:3,rounds:9,workCounts,rows,scope:'Warm public-rule AST/symbols outside timing; complete-module setup included. Plain metadata timed, name getters untimed only. Independent exact N default-read messages/spans and empty statement findings; full errors empty; complete equality outside clock. No editor/cold/heap claim.'},null,2));
