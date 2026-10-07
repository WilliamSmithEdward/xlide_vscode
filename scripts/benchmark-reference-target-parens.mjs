import { build } from 'esbuild';import { execFileSync } from 'node:child_process';import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';import { tmpdir } from 'node:os';import { join, dirname, relative } from 'node:path';import { createRequire } from 'node:module';import { performance } from 'node:perf_hooks';import { createHash } from 'node:crypto';import assert from 'node:assert/strict';
const shape=process.argv.find(a=>a.startsWith('--shape='))?.slice(8)??'paren-target';
if(!['paren-target','inline-if'].includes(shape))throw Error('Unknown benchmark shape: '+shape);
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),root=process.cwd(),scratch=mkdtempSync(join(tmpdir(),'xlide-target-bench-')),path=join(scratch,'api.cjs'),require=createRequire(import.meta.url),rounds=15,warmups=3;
try{
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/referenceKinds\.ts$/},args=>{const file=relative(root,args.path).replaceAll('\\','/');if(file!=='src/analyzer/references/referenceKinds.ts')return;return{contents:execFileSync('git',['show',baseline+':'+file],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)};});}}]:[];
 const bundle=await build({plugins,stdin:{contents:"export { classifyReferenceKinds } from './src/analyzer/references/referenceKinds';export { tokenizeCached } from './src/analyzer/lexer/tokenize';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,bundle.outputFiles[0].contents);const api=require(path),rows=[];
 for(const depth of [0,5,100,1000])for(const statements of depth===1000?[1]:[1,100])for(const mode of ['terminal','all'])for(const scope of ['cached-tokens','fresh-tokens']){
  const statement=depth===0?'a(i) = 1':shape==='inline-if'?'If c Then '.repeat(depth)+'x = 1':'a('+'f('.repeat(depth)+'1'+')'.repeat(depth)+').Value = 1',source=Array.from({length:statements},()=>statement).join('\n')+'\n';
  const tokens=api.tokenizeCached(source),offsets=tokens.filter(t=>t.kind==='identifier'&&(mode==='all'||t.rawText===(depth===0?'a':shape==='inline-if'?'x':'Value'))).map(t=>t.start);
  const times=[];let expected,sha;
  for(let i=-warmups;i<rounds;i++){
   const text=scope==='fresh-tokens'?source+"' fresh "+i+'\n':source;
   const start=performance.now(),result=api.classifyReferenceKinds(text,offsets),elapsed=performance.now()-start,value=[...result];
   if(i===-warmups){expected=value;sha=createHash('sha256').update(JSON.stringify(value)).digest('hex');}else assert.deepEqual(value,expected);
   assert.equal(result.size,offsets.length);assert.equal(value.filter(v=>v[1]==='write').length,statements);
   if(i>=0)times.push(elapsed);
  }
  times.sort((a,b)=>a-b);rows.push({depth,statements,mode,scope,medianMs:times[7],p95Ms:times[14],resultSha256:sha});
 }
 console.log(JSON.stringify({baseline:baseline??'working-tree',shape,node:process.version,rounds,warmups,scope:'Actual syntactic classifier; stress target strings, no VBE execution claim; source/offset construction and complete ordered Map assertions outside timer; fresh-tokens changes trailing comment to include lexing',rows},null,2));
}finally{if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
