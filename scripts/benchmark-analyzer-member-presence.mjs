import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { readFileSync, mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname, relative } from 'node:path';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11), rounds=15,warmups=3;
const root=process.cwd(),scratch=mkdtempSync(join(tmpdir(),'xlide-member-presence-')),path=join(scratch,'api.cjs'),require=createRequire(import.meta.url);
const hash=value=>createHash('sha256').update(JSON.stringify(value)).digest('hex');
try{
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/(undeclared|memberAccess)\.ts$/},args=>{
  const file=relative(root,args.path).replaceAll('\\','/');if(!['src/analyzer/diagnostics/rules/undeclared.ts','src/analyzer/completion/memberAccess.ts'].includes(file))return;
  return{contents:execFileSync('git',['show',baseline+':'+file],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)};
 });}}]:[];
 const result=await build({plugins,stdin:{contents:"export { checkMemberNotFound } from './src/analyzer/diagnostics/rules/undeclared'; export { walkProcedureStatements } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { tokenizeCached } from './src/analyzer/lexer/tokenize'; export { analyzeModule } from './src/analyzer/diagnostics/analyzeModule';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(path,result.outputFiles[0].contents);const api=require(path),rows=[];
 const makeType=(name,privateMembers)=>({name,moduleName:name,kind:'document',members:[],privateMembers});
 const scenarios=['no-private','unrelated','positive-first','positive-last','large-cold-negative'];
 for(const scenario of scenarios)for(const references of scenario==='large-cold-negative'?[1]:[1,1000]){
  const member=scenario.startsWith('positive')?'Hidden':'Value';
  const source='Option Explicit\nSub Main()\nDim r As Range\nDim c As C\nDim result As Variant\n'+Array.from({length:references},()=> 'result = '+(member==='Hidden'?'c.Hidden':'r.Value')).join('\n')+'\nEnd Sub\n';
  const parsed=api.parseModule(source),tokens=api.tokenizeCached(source).filter(t=>t.kind!=='comment');
  const metadata=()=>{
   const count=scenario==='large-cold-negative'?1000:10;
   const types=Array.from({length:count},(_,i)=>makeType('Other'+i,scenario==='no-private'?[]:Array.from({length:scenario==='large-cold-negative'?100:10},(_,j)=>'Unused'+i+'_'+j)));
   const c=makeType('C',scenario.startsWith('positive')?['Hidden']:[]);
   return scenario==='positive-first'?[c,...types]:[...types,c];
  };
  for(const scope of ['rule','full','full-fresh']){
   let expected,expectedHash;const timings=[];
   for(let i=-warmups;i<rounds;i++){
    const projectClassMembers=metadata(),text=scope==='full-fresh'?source+"' uncached "+i+'\n':source;
    const ctx={projectClassMembers,parsedModule:parsed,sourceTokens:tokens,receiverTypeCache:new Map(),receiverChainCache:new Map(),memberSurfaceCache:new Map(),withScanCache:new Map()};
    const diagnostics=[],push=(...args)=>diagnostics.push(args);const start=performance.now();
    const value=scope==='rule'?(api.walkProcedureStatements(parsed,undefined,[api.checkMemberNotFound(source,ctx,push)],undefined,{source,takes:[true]}),diagnostics):api.analyzeModule(text,{projectClassMembers,onInternalError(e){throw e;}});
    const elapsed=performance.now()-start;
    if(i===-warmups){expected=value;expectedHash=hash(value);}else assert.deepEqual(value,expected);
    if(i>=0)timings.push(elapsed);
   }
   timings.sort((a,b)=>a-b);rows.push({scenario,references,scope,medianMs:timings[7],p95Ms:timings[14],resultSha256:expectedHash});
  }
 }
 console.log(JSON.stringify({baseline:baseline??'working-tree',node:process.version,rounds,warmups,scope:'Actual diagnostic rule and complete analyzer; mock in-memory project metadata; fresh lists per round; metadata generation and assertions outside clock; full-fresh changes trailing comment to bypass parser cache; no Office/UI timing',rows},null,2));
}finally{if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
