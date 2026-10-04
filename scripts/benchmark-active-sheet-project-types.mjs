import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??15);
if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3 to 100');
const dir=mkdtempSync(join(tmpdir(),'xlide-activesheet-')),bundle=join(dir,'analyzer.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/rules[\\/]assignments\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/assignments.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export { checkSetAssignments } from './src/analyzer/diagnostics/rules/assignments'; export { tokenizeCached } from './src/analyzer/lexer/tokenize'; export { forEachStatementWithHeaders } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(bundle,result.outputFiles[0].contents);api=createRequire(import.meta.url)(bundle);
}finally{try{unlinkSync(bundle);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(dir);}


const rows=[];
for(const count of [10,100,1000,3000])for(const mode of ['object','worksheet','class','collection','shadow','other-value']){
 const surfaces=Array.from({length:count},(_,i)=>({name:'K'+i,kind:'class',moduleName:'K'+i,members:[],exhaustive:true}));surfaces.push({name:'Box',kind:'class',moduleName:'Box',members:[],exhaustive:true});
 const type=mode==='worksheet'?'Worksheet':mode==='class'?'Box':mode==='collection'?'Collection':'Object';
 const source='Sub P()\nDim c As '+type+'\n'+(mode==='shadow'?'Dim ActiveSheet As Object\n':'')+('Set c = '+(mode==='other-value'?'Nothing':'ActiveSheet')+'\n').repeat(1000)+'End Sub';const mod=api.parseModule(source),tokens=api.tokenizeCached(source).filter(t=>t.kind!=='comment');let previous;
 function prepare(){const symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod});if(symbols.root===previous)throw Error('Reused root');previous=symbols.root;const ctx={projectClassMembers:surfaces,parsedModule:mod,sourceTokens:tokens,withScanCache:new Map(),receiverTypeCache:new Map(),receiverChainCache:new Map(),memberSurfaceCache:new Map(),allowSetAssignmentRefinement:false};return()=>{let hits=0;const visitor=api.checkSetAssignments(source,symbols,undefined,ctx,()=>hits++);for(const proc of mod.members)if(proc.kind==='Procedure'){const visit=visitor(proc);if(visit)api.forEachStatementWithHeaders(source,proc.body,visit);}if(hits!==((mode==='class'||mode==='collection')?1000:0))throw Error('Wrong hits '+mode+': '+hits);return hits;};}
 for(let i=0;i<3;i++)prepare()();const samples=[];for(let i=0;i<rounds;i++){const run=prepare(),start=performance.now();run();samples.push(performance.now()-start);}samples.sort((a,b)=>a-b);
 rows.push({name:count+'-'+mode,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
