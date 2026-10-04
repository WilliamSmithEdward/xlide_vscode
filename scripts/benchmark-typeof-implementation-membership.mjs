import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {dirname} from 'node:path';
import {performance} from 'node:perf_hooks';
import {writeFileSync,unlinkSync,mkdtempSync,rmdirSync} from 'node:fs';
import {tmpdir,cpus} from 'node:os';
import {join} from 'node:path';
import {createRequire} from 'node:module';
import assert from 'node:assert/strict';
const baseline=process.argv.find(s=>s.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-typeof-membership-')),file=join(scratch,'api.cjs');let a;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]rules[\\/]typeOfIs\.ts$/},p=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/typeOfIs.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(p.path)}));}}]:[];
 const b=await build({plugins,stdin:{contents:"export {checkTypeOfIsCompatibility} from './src/analyzer/diagnostics/rules/typeOfIs';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,b.outputFiles[0].contents);a=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file)}finally{rmdirSync(scratch)}}
const rows=[],workCounts=[];
for(const kind of ['repeated-missing','distinct-missing','host-direct','host-reverse'])for(const n of [1,100,1000]){
 let reads=0;const names=Array.from({length:n},(_,i)=>i===n-1&&kind.startsWith('host-')?'Worksheet':'Other'+i),list=new Proxy(names,{get(t,k,r){if(typeof k==='string'&&/^(0|[1-9]\d*)$/.test(k))reads++;return Reflect.get(t,k,r)}});
 const surface=(name,impl=[])=>({name,moduleName:name,kind:'class',members:[],implements:impl}),classes=[surface('Class1',list),...Array.from({length:n},(_,i)=>surface('Class'+(i+2)))];
 const operand=kind==='host-reverse'?'Worksheet':'Class1',targets=Array.from({length:n},(_,i)=>kind==='host-direct'?'Worksheet':kind==='host-reverse'?'Class1':'Class'+(kind==='distinct-missing'?i+2:2));
 const source=['Option Explicit',`Sub Go(ByVal actor As ${operand})`,...targets.map(t=>`If TypeOf actor Is ${t} Then\nDebug.Print 1\nEnd If`),'End Sub',''].join('\n');
 const m=a.parseModule(source),symbols=a.buildModuleSymbols('M','standard',source,{parsedModule:m});
 const rule=()=>{const out=[],factory=a.checkTypeOfIsCompatibility(symbols,{projectClassMembers:classes},(...v)=>out.push(v));for(const p of m.members)if(p.kind==='Procedure'){const visit=factory(p);if(!visit)continue;for(const node of p.body)if(node.kind==='IfBlock')for(const branch of node.branches)if(branch.condition)visit(branch.condition)}return out};
 let from=0;const expected=kind.startsWith('host-')?[]:targets.map(t=>{const text='TypeOf actor Is '+t,start=source.indexOf(text,from);from=start+text.length;return ['typeOfIsAlwaysFalse',`'TypeOf ... Is ${t}' is always False: 'actor' is declared As Class1, which is never ${t}.`,{start,end:from}]});
 assert.deepEqual(rule(),expected);workCounts.push({kind,n,reads});if(!baseline)assert.ok(reads<=n*3);
 // Counter proxy is used only in the untimed check; timed lists are plain arrays.
 classes[0].implements=names;
 const module=()=>{const errors=[];const diagnostics=a.analyzeModule(source,{projectClassMembers:classes,onInternalError:e=>errors.push(String(e))});return {diagnostics,errors}};
 const moduleExpected=module();assert.deepEqual(moduleExpected.errors,[]);assert.equal(moduleExpected.diagnostics.length,expected.length);
 for(let i=0;i<expected.length;i++){assert.equal(moduleExpected.diagnostics[i].code,'typeof-is-always-false');assert.equal(moduleExpected.diagnostics[i].message,expected[i][1]);assert.deepEqual(moduleExpected.diagnostics[i].span,expected[i][2]);}
 for(const scope of ['rule','complete-module-diagnostics']){
  const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=scope==='rule'?rule():module(),elapsed=performance.now()-begin;assert.deepEqual(actual,scope==='rule'?expected:moduleExpected);if(round>=0)samples.push(elapsed)}
  samples.sort((a,b)=>a-b);rows.push({kind,n,scope,medianMs:+samples[4].toFixed(5)});
 }
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,workCounts,rows,scope:'Warm direct-rule AST/symbol setup outside timing. Full-module includes analyzer setup. Plain metadata timed; array-index proxy only untimed. Exact independently expected diagnostics/messages/spans, full-output equality outside clock. No editor/cold/heap claim.'},null,2));
