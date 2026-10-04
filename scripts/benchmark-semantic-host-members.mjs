import { build } from 'esbuild';import { execFileSync } from 'node:child_process';import { readFileSync, mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';import { tmpdir } from 'node:os';import { join, relative, dirname } from 'node:path';import { createRequire } from 'node:module';import { createHash } from 'node:crypto';import { performance } from 'node:perf_hooks';import assert from 'node:assert/strict';
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),rounds=15,warmups=3,root=process.cwd(),scratch=mkdtempSync(join(tmpdir(),'xlide-semantic-bench-')),path=join(scratch,'api.cjs'),require=createRequire(import.meta.url);
try{
 const files=new Set(['src/analyzer/completion/memberAccess.ts','src/analyzer/semantic/typeSemanticTokens.ts']);
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/\.ts$/},args=>{const file=relative(root,args.path).replaceAll('\\','/');if(!files.has(file))return;return{contents:execFileSync('git',['show',baseline+':'+file],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)};});}}]:[];
 const bundle=await build({plugins,stdin:{contents:"export { collectHostMemberMethodTokens } from './src/analyzer/semantic/typeSemanticTokens';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,bundle.outputFiles[0].contents);const api=require(path),rows=[];
 const model=count=>({source:'Synthetic timing model',types:{'Excel.Range':{displayName:'Range',members:Array.from({length:count},(_,i)=>({name:'Member'+i,kind:i%2?'property':'method'}))},'Excel.Other':{displayName:'Other',members:[{name:'OtherKnown',kind:'method'}]}},aliases:{range:'Excel.Range'},globals:{}});
 for(const members of [5,1000])for(const scenario of ['first','last','missing','combined-last','unknown','model-cold-first'])for(const references of scenario==='model-cold-first'?[1]:[1,1000])for(const scope of scenario==='model-cold-first'?['cached-source']:['cached-source','fresh-source']){
  const name=scenario==='first'||scenario==='model-cold-first'?'Member0':scenario==='missing'?'OtherKnown':scenario==='unknown'?'NotKnown':'Member'+(members-1),receiver=scenario==='combined-last'?'Sheet1':'r';
  const source='Sub Main()\nDim r As Range\n'+Array.from({length:references},()=>receiver+'.'+name).join('\n')+'\nEnd Sub\n',sharedModel=model(members),context=scenario==='combined-last'?{codeNames:{sheet1:'Excel.Range'},projectTypes:[{name:'Sheet1',kind:'document'}]}:{};
  const times=[];let expected,sha;
  for(let i=-warmups;i<rounds;i++){
   const ctx={...context,model:scenario==='model-cold-first'?model(members):sharedModel},text=scope==='fresh-source'?source+"' fresh "+i+'\n':source;
   const start=performance.now(),tokens=api.collectHostMemberMethodTokens(text,ctx),elapsed=performance.now()-start;
   if(i===-warmups){expected=tokens;sha=createHash('sha256').update(JSON.stringify(tokens)).digest('hex');}else assert.deepEqual(tokens,expected);
   const expectedCount=scenario==='missing'||scenario==='unknown'?0:references;assert.equal(tokens.length,expectedCount);
   if(i>=0)times.push(elapsed);
  }
  times.sort((a,b)=>a-b);rows.push({members,scenario,references,scope,medianMs:times[7],p95Ms:times[14],resultSha256:sha});
 }
 console.log(JSON.stringify({baseline:baseline??'working-tree',node:process.version,rounds,warmups,scope:'Actual host member semantic collector; synthetic ordinary host metadata; metadata construction/assertions outside clock; cached/fresh sources; model-cold-first uses a fresh model each round; no editor/Office timing',rows},null,2));
}finally{if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
