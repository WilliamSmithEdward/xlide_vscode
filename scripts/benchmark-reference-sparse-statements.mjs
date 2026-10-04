// Run: node scripts/benchmark-reference-sparse-statements.mjs [--baseline=COMMIT]
import {build} from 'esbuild';import assert from 'node:assert/strict';import {execFileSync} from 'node:child_process';import {mkdtempSync,writeFileSync,unlinkSync,rmdirSync} from 'node:fs';import {createRequire} from 'node:module';import {cpus,tmpdir} from 'node:os';import {dirname,join} from 'node:path';import {performance} from 'node:perf_hooks';import {fileURLToPath} from 'node:url';
const root=dirname(dirname(fileURLToPath(import.meta.url))),baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-reference-sparse-')),file=join(scratch,'api.cjs');let api;
try{
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/referenceKinds\.ts$/},args=>({contents:execFileSync('git',['show',`${baseline}:src/analyzer/references/referenceKinds.ts`],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];
 const built=await build({plugins,stdin:{contents:"export {classifyReferenceKinds} from './src/analyzer/references/referenceKinds';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}
const rows=[];
for(const count of [1,100,10000])for(const mode of ['first','middle','last','separated','dense','long-statement']){
 let source,offsets,expected;
 if(mode==='long-statement'){
  source='value = f('+Array(count).fill('arg').join(',')+')';offsets=[source.lastIndexOf('arg')];expected=[[offsets[0],'read']];
 }else{
  const line='value = value + 1\n';source=line.repeat(count);
  const selected=mode==='dense'?Array.from({length:count},(_,i)=>i):[...new Set(mode==='first'?[0]:mode==='middle'?[Math.floor(count/2)]:mode==='last'?[count-1]:[0,count-1])];
  const writes=selected.map(i=>i*line.length),reads=writes.map(i=>i+8);
  offsets=[...writes,...reads].toReversed();expected=[...writes.map(i=>[i,'write']),...reads.toReversed().map(i=>[i,'read'])];
 }
 const samples=[];
 for(let round=-3;round<9;round++){
  const start=performance.now();let result;
  for(let i=0;i<20;i++)result=api.classifyReferenceKinds(source,offsets);
  const elapsed=(performance.now()-start)/20;
  assert.deepEqual([...result],expected);
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({count,mode,sourceCharacters:source.length,queriedOffsets:offsets.length,medianMs:samples[4],maxBatchAverageMs:samples[8]});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,warmups:3,rounds:9,batchSize:20,scope:'Warm complete reference classification including token-cache retrieval. Lexing primed in warmups; fixture construction and independent complete Map assertions excluded. Cold lexing still reads the whole module; no whole-refactor/editor latency claim.',rows},null,2));
