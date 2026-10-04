// Run: node scripts/benchmark-refactor-local-occurrences.mjs [--baseline=COMMIT]
import {build} from 'esbuild';
import assert from 'node:assert/strict';
import {execFileSync} from 'node:child_process';
import {mkdtempSync,writeFileSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {cpus,tmpdir} from 'node:os';
import {dirname,join,relative} from 'node:path';
import {performance} from 'node:perf_hooks';
import {fileURLToPath} from 'node:url';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const scratch=mkdtempSync(join(tmpdir(),'xlide-local-occurrences-'));
const file=join(scratch,'api.cjs');
let api;
try {
 const plugins=baseline?[{name:'baseline',setup(builder){
  builder.onLoad({filter:/(?:vbaSourceScan|shared)\.ts$/},args=>{
   const path=relative(root,args.path).replace(/\\/g,'/');
   if(!['src/vbaSourceScan.ts','src/analyzer/refactor/shared.ts'].includes(path))return;
   return {contents:execFileSync('git',['show',`${baseline}:${path}`],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)};
  });
 }}]:[];
 const built=await build({plugins,stdin:{contents:"export {parseModule} from './src/analyzer/parser/parseModule';export {localUsesIn} from './src/analyzer/refactor/shared';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,built.outputFiles[0].contents);
 api=createRequire(import.meta.url)(file);
} finally {
 try {unlinkSync(file);} catch(error){if(error.code!=='ENOENT')throw error;}
 rmdirSync(scratch);
}
const rows=[];
for(const count of [1,100,1000])for(const position of ['first','middle','last']){
 const at=position==='first'?0:position==='last'?count:Math.floor(count/2);
 const others=Array.from({length:count},(_,i)=>`Sub Other${i}()\nDim target As Long\ntarget = 1\nDebug.Print target\nEnd Sub\n`);
 const before=others.slice(0,at).join('');
 const source=before+'Sub Main()\nDim target As Long\ntarget = 7\nDebug.Print target\nEnd Sub\n'+others.slice(at).join('');
 const module=api.parseModule(source);
 const procedure=module.members.find(m=>m.kind==='Procedure'&&m.name==='Main');
 const declaration=procedure.body.find(n=>n.kind==='VariableGroup');
 const write={line:at*5+2,column:0,offset:before.length+'Sub Main()\nDim target As Long\n'.length,text:'target'};
 const read={line:at*5+3,column:12,offset:write.offset+'target = 7\nDebug.Print '.length,text:'target'};
 const expected={uses:[write,read],writes:[write]},samples=[];
 for(let round=-3;round<9;round++){
  const start=performance.now();let result;
  for(let i=0;i<10;i++)result=api.localUsesIn(source,procedure,declaration.span,'target');
  const elapsed=(performance.now()-start)/10;
  assert.deepEqual(result,expected);
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);
 rows.push({count,position,sourceCharacters:source.length,medianMs:samples[4],maxBatchAverageMs:samples[8]});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,warmups:3,rounds:9,batchSize:10,scope:'Warm localUsesIn query, including occurrence matching and full reference classification. Parsing outside timer, source stripping and token caches primed in warmups. Input construction and independent complete result assertions excluded. Cold stripping and reference classification still scan the whole source; no whole-refactor/editor latency claim.',rows},null,2));
