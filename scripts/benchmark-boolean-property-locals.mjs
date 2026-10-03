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
const dir=mkdtempSync(join(tmpdir(),'xlide-boolean-property-')),bundle=join(dir,'analyzer.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/rules[\\/]assignments\.ts$/},()=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/assignments.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/diagnostics/rules')}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export { checkAssignmentTypes } from './src/analyzer/diagnostics/rules/assignments'; export { tokenizeCached } from './src/analyzer/lexer/tokenize'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(bundle,result.outputFiles[0].contents);api=createRequire(import.meta.url)(bundle);
}finally{try{unlinkSync(bundle);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(dir);}


const rows=[];
for(const count of [10,100,1000,3000])for(const mode of ['long','true','false','integer','literal','distinct','true-first']){
 const declarations=Array.from({length:count},(_,i)=>'Dim k'+i+' As Long\n').join('');
 const type=mode==='true'||mode==='false'||mode==='true-first'?'Boolean':mode==='integer'?'Integer':'Long',initial=mode==='long'?'500':mode==='false'?'False':'True',expr=mode==='literal'?'12':'v';
 const declaration='Dim v As '+type+'\n';const assignments=Array.from({length:1000},(_,i)=>'Range("A1").Font.Size = '+(mode==='distinct'?'k'+(i%count):expr)+'\n').join('');const source='Sub P()\n'+(mode==='true-first'?declaration:'')+declarations+(mode==='true-first'?'':declaration)+'v = '+initial+'\n'+assignments+'End Sub';const mod=api.parseModule(source),tokens=api.tokenizeCached(source).filter(t=>t.kind!=='comment');let previous;
 function prepare(){const symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod});if(symbols.root===previous)throw Error('Reused root');previous=symbols.root;
 const ctx={parsedModule:mod,sourceTokens:tokens,withScanCache:new Map(),receiverTypeCache:new Map(),receiverChainCache:new Map(),memberSurfaceCache:new Map(),allowSetAssignmentRefinement:false};return()=>{let hits=0;api.checkAssignmentTypes(source,mod,symbols,undefined,ctx,undefined,()=>hits++);if(hits!==((mode==='long'||mode==='false'||mode==='integer'||mode==='distinct')?1000:0))throw Error('Wrong hits '+mode+': '+hits);return hits;};}
 for(let i=0;i<3;i++)prepare()();const samples=[];for(let i=0;i<rounds;i++){const run=prepare(),start=performance.now();run();samples.push(performance.now()-start);}samples.sort((a,b)=>a-b);
 rows.push({name:count+'-'+mode,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
