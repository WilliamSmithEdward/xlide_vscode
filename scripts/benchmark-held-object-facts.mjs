// Run: node scripts/benchmark-held-object-facts.mjs [--baseline=COMMIT] [--rounds=9]
import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {writeFileSync,mkdtempSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {tmpdir,cpus} from 'node:os';
import {dirname,join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const rounds=Number(process.argv.find(arg=>arg.startsWith('--rounds='))?.slice(9)??9);
if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3..100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-held-facts-')),file=join(scratch,'api.cjs');let api;
try{
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/heldObjects\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/heldObjects.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];
 const built=await build({plugins,stdin:{contents:"export {heldObjectsAt} from './src/analyzer/diagnostics/heldObjects';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}
const rows=[];
for(const count of [100,1000,10000]){
 const source='Sub Work()\nDim c As New Collection\nc.Add 1\n'+'x = c.Count\n'.repeat(count)+'End Sub\n';
 const module=api.parseModule(source),proc=module.members.find(member=>member.kind==='Procedure');assert(proc);assert.equal(proc.body.length,count+2);
 const samples=[];
 for(let round=-3;round<rounds;round++){
  // A fresh symbol identity makes every round a cold facts-cache query; lexer,
  // parsing and symbol construction are outside the measured component.
  const symbols={...api.buildModuleSymbols('M','standard',source,{parsedModule:module})};
  const start=performance.now();const facts=Array.from({length:3},()=>api.heldObjectsAt(source,proc,symbols,undefined));const elapsed=performance.now()-start;
  if(round>=0)samples.push(elapsed);
  for(const at of facts)for(let i=0;i<proc.body.length;i++){
   assert.deepEqual([...at(proc.body[i]).classes],i===0?[]:[['c','Collection']]);
   assert.deepEqual([...at(proc.body[i]).items],i===0?[]:[['c',i===1?[]:['(value)']]]);
  }
 }
 samples.sort((a,b)=>a-b);rows.push({statements:count,consumers:3,allStatementFactsCorrect:true,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,scope:'three default fact consumers, fresh symbol snapshot; all statement facts checked outside timing',rows},null,2));
