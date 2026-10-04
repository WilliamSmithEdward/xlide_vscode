// Run: node scripts/benchmark-inference-reader-source.mjs [--baseline=COMMIT] [--rounds=9]
import {build} from 'esbuild';import {execFileSync} from 'node:child_process';import {readFileSync,writeFileSync,mkdtempSync,unlinkSync,rmdirSync} from 'node:fs';import {createRequire} from 'node:module';import {tmpdir,cpus} from 'node:os';import {dirname,join} from 'node:path';import {fileURLToPath} from 'node:url';import {performance} from 'node:perf_hooks';import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url))),baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??9);if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3..100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-inference-reader-source-')),file=join(scratch,'api.cjs');let api;
try{const plugins=[{name:'inference-reader-query',setup(builder){builder.onLoad({filter:/typeInference\.ts$/},args=>({contents:(baseline?execFileSync('git',['show',baseline+':src/analyzer/diagnostics/typeInference.ts'],{cwd:root,encoding:'utf8'}):readFileSync(args.path,'utf8'))+'\nexport {callEffectsFor,declaredFactsFor};',loader:'ts',resolveDir:dirname(args.path)}));}}];const built=await build({plugins,stdin:{contents:"export {callEffectsFor,declaredFactsFor} from './src/analyzer/diagnostics/typeInference';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {tokenizeCached} from './src/analyzer/lexer/tokenize';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}

const rows=[];
for(const bytes of [1000,100000,1000000])for(const mode of ['same','copied'])for(const consumer of ['effects','declarations']) {
 const samples=[];
 for(let round=-3;round<rounds;round++) {
  const source='Const Limit As Long = 3\nSub Zero(ByRef value As Long)\nvalue = 0\nEnd Sub\nSub Main()\nDim n As Long\nDim fixed(3) As Long\nZero n\nEnd Sub\n'+"' unique round "+round+'\n'+"' filler\n".repeat(Math.ceil(bytes/9));
  const caller=mode==='same'?source:('x'+source).slice(1),mod=api.parseModule(source),symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod}),proc=mod.members.find(m=>m.kind==='Procedure'&&m.name==='Main');
  const query=s=>consumer==='effects'?api.callEffectsFor(s,symbols,undefined):api.declaredFactsFor(s,symbols,proc),expected=query(source);
  if(consumer==='effects')assert.deepEqual([...expected(api.tokenizeCached('Zero n').filter(t=>t.kind!=='eof'))].map(([name,toks])=>[name,toks.map(t=>t.rawText)]),[['n',['0']]]);
  else {assert.equal(expected.type('n'),'long');assert.equal(expected.type('fixed'),'long()');assert.deepEqual(expected.bounds('fixed'),[0,3]);assert.equal(expected.constant('limit'),3);assert.equal(expected.constant('n'),undefined);}
  const start=performance.now();for(let i=0;i<1000;i++)assert.equal(query(caller),expected);const elapsed=performance.now()-start;if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({bytes,mode,consumer,calls:1000,allAnswersCorrect:true,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,scope:'1000 warm private inference reader lookups; independent effects/declaration assertions and identity checks; construction/parse/binding/priming/callback evaluation excluded; temporary benchmark exports only',rows},null,2));
