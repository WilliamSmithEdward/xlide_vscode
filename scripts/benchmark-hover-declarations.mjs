// Run: node scripts/benchmark-hover-declarations.mjs [--baseline=COMMIT] [--rounds=9]
import {build} from 'esbuild';import {execFileSync} from 'node:child_process';import {writeFileSync,mkdtempSync,unlinkSync,rmdirSync} from 'node:fs';import {createRequire} from 'node:module';import {tmpdir,cpus} from 'node:os';import {dirname,join} from 'node:path';import {fileURLToPath} from 'node:url';import {performance} from 'node:perf_hooks';import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url))),baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??9);if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3..100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-hover-declarations-')),file=join(scratch,'api.cjs');let api;
try{const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/resolveHover\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/hover/resolveHover.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];const built=await build({plugins,stdin:{contents:"export {resolveHover} from './src/analyzer/hover/resolveHover';export {editorModuleSymbols} from './src/analyzer/symbols/editorModuleSymbols';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}

const rows=[];
for(const count of [1,100,1000])for(const kind of ['procedures','locals']) {
 const samples=[];
 for(let round=-3;round<rounds;round++) {
  const prefix=kind==='procedures'?Array.from({length:count},(_,i)=>'Sub P'+i+'()\nEnd Sub\n').join(''):'',locals=kind==='locals'?Array.from({length:count},(_,i)=>'Dim Local'+i+' As Long\n').join(''):'';
  const source=prefix+'Sub Target()\n'+locals+'Dim HoverValue As Long\nDebug.Print hovervalue\nEnd Sub\n'+"' round "+round,start=source.lastIndexOf('hovervalue'),offset=start+2,expected={signature:'HoverValue As Long',details:['Local in Target'],span:{start,end:start+10}};
  assert.deepEqual(api.resolveHover(source,offset),expected);
  const startTime=performance.now(),results=Array.from({length:100},()=>api.resolveHover(source,offset));const elapsed=performance.now()-startTime;
  if(round>=0)samples.push(elapsed);for(const result of results)assert.deepEqual(result,expected);
 }
 samples.sort((a,b)=>a-b);rows.push({count,kind,requests:100,completeResultsCorrect:true,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
for(const count of [1,100,1000]) {
 const samples=[];
 for(let round=-3;round<rounds;round++) {
  const source='Sub Target()\nDim HoverValue As Long\nDebug.Print hovervalue\nEnd Sub\n'+Array.from({length:count},(_,i)=>'Sub P'+i+'()\nEnd Sub\n').join('')+"' cold round "+round;
  const start=source.indexOf('hovervalue'),offset=start+2,expected={signature:'HoverValue As Long',details:['Local in Target'],span:{start,end:start+10}};
  api.editorModuleSymbols('Module','standard',source);
  const startTime=performance.now(),result=api.resolveHover(source,offset),elapsed=performance.now()-startTime;
  if(round>=0)samples.push(elapsed);assert.deepEqual(result,expected);
 }
 samples.sort((a,b)=>a-b);rows.push({count,kind:'firstScopeColdIndex',requests:1,completeResultsCorrect:true,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,scope:'Warm rows: 100 complete hovers with parse/symbol/index warmed. Cold-index rows: one first-scope hover after symbols warmed, index build included. Construction excluded.',rows},null,2));
