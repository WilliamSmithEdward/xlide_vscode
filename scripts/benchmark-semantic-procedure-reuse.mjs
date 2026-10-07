// Run: node scripts/benchmark-semantic-procedure-reuse.mjs [--baseline=COMMIT]
// Optional: XLIDE_PERF_WORKBOOK points to a read-only ROneCOne fixture.
import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {mkdtempSync,writeFileSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {tmpdir,cpus} from 'node:os';
import {dirname,join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url))),baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-semantic-procedure-')),file=join(scratch,'api.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/typeSemanticTokens\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/semantic/typeSemanticTokens.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];
 const built=await build({plugins,stdin:{contents:"export {collectTypeNameReferences} from './src/analyzer/semantic/typeSemanticTokens';export {parseModule} from './src/analyzer/parser/parseModule';export {readModule} from './src/vba/projectService';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}
const fixtures=[1,100,1000].map(count=>({name:'synthetic-'+count,source:Array.from({length:count},(_,i)=>'Sub Ref'+i+'()\nDim item As Collection\n'+'Set item = New Collection\n'.repeat(10)+'End Sub\n').join('\n')}));
if(process.env.XLIDE_PERF_WORKBOOK)fixtures.push({name:'actual-large-class',source:api.readModule(process.env.XLIDE_PERF_WORKBOOK,'ROneCOne').source});
const rows=[];
for(const fixture of fixtures)for(const scenario of ['retained-procedures','cold-procedures']){
 const eol=fixture.source.includes('\r\n')?'\r\n':'\n',prefix=fixture.source+eol+'Sub XLIDE_SemanticProbe()'+eol+'Dim n As Long'+eol+'n = 100'+eol+'End Sub'+eol;
 const original=api.parseModule(prefix),expected=api.collectTypeNameReferences(prefix),samples=[];let reusedProcedures=0;
 for(let round=-3;round<9;round++){
  const changed=(scenario==='cold-procedures'?"' cold snapshot "+round+eol+prefix:prefix).replace('n = 100'+eol,'n = '+(200+round)+eol);
  const parsed=api.parseModule(changed);
  reusedProcedures=parsed.members.filter(m=>m.kind==='Procedure'&&original.members.includes(m)).length;
  const start=performance.now(),result=api.collectTypeNameReferences(changed),elapsed=performance.now()-start;
  if(round>=0)samples.push(elapsed);
  if(scenario==='retained-procedures')assert.deepEqual(result,expected);
  else{const delta=changed.length-prefix.length;assert.deepEqual(result,expected.map(ref=>({...ref,span:{start:ref.span.start+delta,end:ref.span.end+delta},...(ref.qualifierSpan?{qualifierSpan:{start:ref.qualifierSpan.start+delta,end:ref.qualifierSpan.end+delta}}:{}),...(ref.fullSpan?{fullSpan:{start:ref.fullSpan.start+delta,end:ref.fullSpan.end+delta}}:{})})));}
  for(const ref of result)assert.equal(changed.slice(ref.span.start,ref.span.end).toLowerCase(),ref.name.toLowerCase());
 }
 samples.sort((a,b)=>a-b);rows.push({name:fixture.name,scenario,characters:fixture.source.length,reusedProcedures,references:expected.length,medianMs:samples[4],p95Ms:samples[8]});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Type-reference collection once per edited snapshot, parsing outside clock, complete outputs checked; cold-procedures adds a leading comment so AST identity is new. Actual fixture read only, no editor latency claim.',rows},null,2));
