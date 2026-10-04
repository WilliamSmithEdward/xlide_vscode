// Run: node scripts/benchmark-module-state-source.mjs [--baseline=COMMIT] [--rounds=9]
import {build} from 'esbuild';import {execFileSync} from 'node:child_process';import {writeFileSync,mkdtempSync,unlinkSync,rmdirSync} from 'node:fs';import {createRequire} from 'node:module';import {tmpdir,cpus} from 'node:os';import {dirname,join} from 'node:path';import {fileURLToPath} from 'node:url';import {performance} from 'node:perf_hooks';import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url))),baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??9);if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3..100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-module-state-source-')),file=join(scratch,'api.cjs');let api;
try{const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/moduleState\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/moduleState.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];const built=await build({plugins,stdin:{contents:"export {untouchedModuleVariables,untouchedModuleVariablesIn} from './src/analyzer/diagnostics/moduleState'; export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols'; export {parseModule} from './src/analyzer/parser/parseModule';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}

const rows=[];
for(const bytes of [1000,100000,1000000])for(const mode of ['same','copied'])for(const scoped of [false,true]) {
 const samples=[];
 for(let round=-3;round<rounds;round++) {
  const source='Private untouched As Long\nSub Run()\nDebug.Print untouched\nEnd Sub\n'+"' unique round "+round+'\n'+"' filler\n".repeat(Math.ceil(bytes/9));
  const caller=mode==='same'?source:('x'+source).slice(1);
  const mod=api.parseModule(source),symbols=api.buildModuleSymbols('Module1','standard',source,{parsedModule:mod}),proc=mod.members.find(m=>m.kind==='Procedure');
  const query=s=>scoped?api.untouchedModuleVariablesIn(s,symbols,proc):api.untouchedModuleVariables(s,symbols);
  const expected=query(source);assert.deepEqual([...expected.keys()],['untouched']);assert.equal(expected.get('untouched').name,'untouched');
  const start=performance.now();for(let i=0;i<1000;i++)assert.equal(query(caller),expected);const elapsed=performance.now()-start;
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({bytes,mode,scoped,calls:1000,allAnswersCorrect:true,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,scope:'1000 warm module-state cache hits; source construction, parse, symbols and priming excluded; independent expected variable and identity checks',rows},null,2));
