import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {dirname,join} from 'node:path';
import {performance} from 'node:perf_hooks';
import {writeFileSync,unlinkSync,mkdtempSync,rmdirSync} from 'node:fs';
import {tmpdir,cpus} from 'node:os';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
const baseline=process.argv.find(s=>s.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-object-state-default-')),file=join(scratch,'api.cjs');let a;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]rules[\\/]objectState\.ts$/},p=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/objectState.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(p.path)}));}}]:[];
 const b=await build({plugins,stdin:{contents:"export {checkObjectVariableNotSet} from './src/analyzer/diagnostics/rules/objectState';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,b.outputFiles[0].contents);a=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file)}finally{rmdirSync(scratch)}}
const rows=[],workCounts=[],lineFor={let:'actor = 1',condition:'If actor Then Debug.Print 1',indexed:'Debug.Print actor(1)'};
for(const path of ['let','condition','indexed'])for(const layout of ['one-procedure','repeated-types','distinct-types'])for(const n of [1,100,1000]){
 let names=0;const classes=Array.from({length:n+1},(_,i)=>({get name(){names++;return 'Class'+i},moduleName:'Class'+i,kind:'class',exhaustive:true,members:[]}));const source=layout==='one-procedure'?['Option Explicit','Sub Go()','Dim actor As Class0',...Array(n).fill(lineFor[path]),'End Sub',''].join('\n'):['Option Explicit',...Array.from({length:n},(_,i)=>'Sub P'+i+'()\nDim actor As Class'+(layout==='distinct-types'?i:0)+'\n'+lineFor[path]+'\nEnd Sub'),''].join('\n'),mod=a.parseModule(source),symbols=a.buildModuleSymbols('M','standard',source,{parsedModule:mod});
 const rule=()=>{const out=[];a.checkObjectVariableNotSet(source,mod,symbols,{projectClassMembers:classes},undefined,(...v)=>out.push(v));return out;};const full=()=>{const errors=[],diagnostics=a.analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});return {diagnostics,errors}};
 const expected=[];if(path!=='condition'){let from=0;for(let i=0;i<n;i++){const start=source.indexOf(lineFor[path],from)+(path==='indexed'?12:0),type='Class'+(layout==='distinct-types'?i:0);from=start+5;expected.push([path==='let'?'set-required':'object-default-value',path==='let'?"Assignment to 'actor' requires Set: "+type+" has no default member for a Let to reach. It is still Nothing here, so this will raise Run-time error '91': Object variable or With block variable not set.":"'actor' is a "+type+", which has no default member to take an index. This will raise Run-time error '438': Object doesn't support this property or method.",{start,end:start+5}]);}if(path==='let'){let from=0;for(let i=0;i<(layout==='one-procedure'?1:n);i++){const start=source.indexOf('Dim actor',from)+4;from=start+5;expected.push(['variable-never-read',"Variable 'actor' is assigned but its value is never read.",{start,end:start+5}]);}}}
 const sort=xs=>xs.slice().sort((x,y)=>x[2].start-y[2].start||x[0].localeCompare(y[0]));
 names=0;assert.deepEqual(rule(),[]);workCounts.push({path,layout,n,scope:'public-rule',names});if(!baseline)assert.ok(names<=n*8+30);
 names=0;const fullExpected=full();assert.deepEqual(fullExpected.errors,[]);assert.deepEqual(sort(fullExpected.diagnostics.map(d=>[d.code,d.message,d.span])),sort(expected));workCounts.push({path,layout,n,scope:'complete-module-diagnostics',names});
 for(const c of classes)Object.defineProperty(c,'name',{value:c.moduleName,writable:true,enumerable:true,configurable:true});
 for(const [scope,run,want] of [['public-rule',rule,[]],['complete-module-diagnostics',full,fullExpected]]){const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=run(),elapsed=performance.now()-begin;assert.deepEqual(actual,want);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({path,layout,n,scope,medianMs:+samples[4].toFixed(5)});}
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,warmups:3,rounds:9,workCounts,rows,scope:'Public-rule parse/symbol setup excluded; complete-module setup included. Getter counters untimed, plain metadata timed; independent empty rule findings and full exact message/code/span tuples, errors empty. Equality outside clock. No editor/cold/heap claim.'},null,2));
