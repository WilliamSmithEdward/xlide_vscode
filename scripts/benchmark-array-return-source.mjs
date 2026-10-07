// Run: node scripts/benchmark-array-return-source.mjs [--baseline=COMMIT] [--rounds=9]
import {build} from 'esbuild';import {execFileSync} from 'node:child_process';import {readFileSync,writeFileSync,mkdtempSync,unlinkSync,rmdirSync} from 'node:fs';import {createRequire} from 'node:module';import {tmpdir,cpus} from 'node:os';import {dirname,join} from 'node:path';import {fileURLToPath} from 'node:url';import {performance} from 'node:perf_hooks';import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url))),baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??9);if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3..100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-array-return-source-')),file=join(scratch,'api.cjs');let api;
try{const plugins=[{name:'array-return-query',setup(builder){builder.onLoad({filter:/[\\/]rules[\\/]arrays\.ts$/},args=>({contents:(baseline?execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/arrays.ts'],{cwd:root,encoding:'utf8'}):readFileSync(args.path,'utf8'))+'\nexport {functionReturnShapes};',loader:'ts',resolveDir:dirname(args.path)}));}}];const built=await build({plugins,stdin:{contents:"export {functionReturnShapes} from './src/analyzer/diagnostics/rules/arrays';export {parseModule} from './src/analyzer/parser/parseModule';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}

const rows=[];
for(const bytes of [1000,100000,1000000])for(const mode of ['same','copied'])for(const optionBase of [0,1]) {
 const samples=[];
 for(let round=-3;round<rounds;round++) {
  const source='Function F() As Variant\nF = Array(1, 2)\nEnd Function\n'+"' unique round "+round+'\n'+"' filler\n".repeat(Math.ceil(bytes/9));
  const caller=mode==='same'?source:('x'+source).slice(1),mod=api.parseModule(source),query=s=>api.functionReturnShapes(s,mod,undefined,optionBase),expected=query(source);
  assert.deepEqual([...expected],[['f',{name:'F()',dims:[{lower:optionBase,upper:optionBase+1,explicitLower:true}],origin:'returned by F',values:[1,2]}]]);
  const start=performance.now();for(let i=0;i<1000;i++)assert.equal(query(caller),expected);const elapsed=performance.now()-start;if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({bytes,mode,optionBase,calls:1000,allAnswersCorrect:true,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,scope:'1000 warm returned-array shape hits; independent complete facts and identity checks; construction/parse/priming excluded; private helper exposed only in benchmark bundle',rows},null,2));
