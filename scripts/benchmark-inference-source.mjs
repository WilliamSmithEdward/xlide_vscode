// Run: node scripts/benchmark-inference-source.mjs [--baseline=COMMIT] [--rounds=9]
import {build} from 'esbuild';import {execFileSync} from 'node:child_process';import {writeFileSync,mkdtempSync,unlinkSync,rmdirSync} from 'node:fs';import {createRequire} from 'node:module';import {tmpdir,cpus} from 'node:os';import {dirname,join} from 'node:path';import {fileURLToPath} from 'node:url';import {performance} from 'node:perf_hooks';import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url))),baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??9);if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3..100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-inference-source-')),file=join(scratch,'api.cjs');let api;
try{const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/typeInference\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/typeInference.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];const built=await build({plugins,stdin:{contents:"export {knownLocalLiteralValuesAt,unreachableStatementsIn,functionResultFor} from './src/analyzer/diagnostics/typeInference';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}

const rows=[];
for(const bytes of [1000,100000,1000000])for(const mode of ['same','copied'])for(const consumer of ['localValues','unreachable','functionCall']) {
 const samples=[];
 for(let round=-3;round<rounds;round++) {
  const source='Sub Run()\nDim n As Long\nn = 3\nDebug.Print n\nGoTo Done\nDebug.Print 0\nDone:\nEnd Sub\nFunction Work() As Long\nWork = 3\nEnd Function\n'+"' unique round "+round+'\n'+"' filler\n".repeat(Math.ceil(bytes/9));
  const caller=mode==='same'?source:('x'+source).slice(1),mod=api.parseModule(source),symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod});
  const proc=mod.members.find(m=>m.kind==='Procedure'&&m.name==='Run'),fn=mod.members.find(m=>m.kind==='Procedure'&&m.name==='Work');
  const query=s=>consumer==='localValues'?api.knownLocalLiteralValuesAt(s,proc,symbols,undefined):consumer==='unreachable'?api.unreachableStatementsIn(s,proc,symbols,undefined):api.functionResultFor(s,fn,symbols,undefined,[]);
  const expected=query(source);
  if(consumer==='localValues') {const stmt=proc.body.find(n=>source.slice(n.span.start,n.span.end).trim()==='Debug.Print n');assert.ok(stmt);assert.deepEqual([...expected(stmt)],[['n',{kind:'number',value:3,origin:'literal'}]]);}
  else if(consumer==='unreachable')assert.deepEqual([...expected].map(n=>source.slice(n.span.start,n.span.end).trim()),['Debug.Print 0']);
  else assert.deepEqual(expected.map(t=>t.rawText),['3']);
  const start=performance.now();for(let i=0;i<1000;i++)assert.equal(query(caller),expected);const elapsed=performance.now()-start;
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({bytes,mode,consumer,calls:1000,allAnswersCorrect:true,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,scope:'1000 warm inference cache hits; construction, parse, symbols and priming excluded; independently expected values/dead statements/call result and identity checks',rows},null,2));
