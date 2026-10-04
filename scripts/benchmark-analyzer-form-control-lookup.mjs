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
const root=process.cwd(),scratch=mkdtempSync(join(tmpdir(),'xlide-form-control-')),path=join(scratch,'api.cjs'),require=createRequire(import.meta.url);
const hash=value=>createHash('sha256').update(JSON.stringify(value)).digest('hex');
try{
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/undeclared\.ts$/},args=>{
  const file=relative(root,args.path).replaceAll('\\','/');if(file!=='src/analyzer/diagnostics/rules/undeclared.ts')return;
  return{contents:execFileSync('git',['show',baseline+':'+file],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)};
 });}}]:[];
 const result=await build({plugins,stdin:{contents:"export { checkMemberNotFound } from './src/analyzer/diagnostics/rules/undeclared'; export { walkProcedureStatements } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { tokenizeCached } from './src/analyzer/lexer/tokenize'; export { analyzeModule } from './src/analyzer/diagnostics/analyzeModule';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(path,result.outputFiles[0].contents);const api=require(path),rows=[];
 const controls=[{name:'T1',type:'MSForms.TextBox'},{name:'T2',type:'MSForms.TextBox'}];
 const makeType=name=>({name,moduleName:name,kind:'document',members:[]});
 const scenarios=['local-form','host-no-form','bare-control-first','bare-control-last','qualified-control'];
 for(const scenario of scenarios)for(const projectTypes of scenario.includes('control')?[1000]:[5,1000])for(const references of [1,1000]){
  const inForm=scenario!=='host-no-form'&&scenario!=='qualified-control',expression=scenario==='local-form'?'r.Value':scenario==='host-no-form'?'ActiveSheet.Name':scenario==='qualified-control'?'f.T1.Nope':'T1.Nope';
  const source='Option Explicit\nSub Main()\nDim r As Range\nDim f As F1\nDim result As Variant\n'+Array.from({length:references},()=> 'result = '+expression).join('\n')+'\nEnd Sub\n';
  const parsed=api.parseModule(source),tokens=api.tokenizeCached(source).filter(t=>t.kind!=='comment');
  const metadata=()=>{
   const types=Array.from({length:projectTypes},(_,i)=>makeType('Other'+i));
   const form={name:'F1',moduleName:'F1',kind:'userform',exhaustive:true,members:controls.map(c=>({name:c.name,moduleName:'F1',kind:'property',returns:c.type}))};
   return scenario==='bare-control-first'?[form,...types]:[...types,form];
  };
  for(const scope of ['rule','full','full-fresh']){
   let expected,expectedHash;const timings=[];
   for(let i=-warmups;i<rounds;i++){
    const projectClassMembers=metadata(),text=scope==='full-fresh'?source+"' uncached "+i+'\n':source;
    const ctx={projectClassMembers,...(inForm?{meProjectType:'F1',meType:'VBA.UserForm',implicitMembers:controls}:{}),parsedModule:parsed,sourceTokens:tokens,receiverTypeCache:new Map(),receiverChainCache:new Map(),memberSurfaceCache:new Map(),withScanCache:new Map()};
    const diagnostics=[],push=(...args)=>diagnostics.push(args);const start=performance.now();
    const value=scope==='rule'?(api.walkProcedureStatements(parsed,undefined,[api.checkMemberNotFound(source,ctx,push)],undefined,{source,takes:[true]}),diagnostics):api.analyzeModule(text,{projectClassMembers,moduleName:inForm?'F1':'M',moduleKind:inForm?'userform':'standard',...(inForm?{meProjectType:'F1',implicitMembers:controls}:{}),onInternalError(e){throw e;}});
    const elapsed=performance.now()-start;
    if(i===-warmups){expected=value;expectedHash=hash(value);}else assert.deepEqual(value,expected);
    if(i>=0)timings.push(elapsed);
   }
   timings.sort((a,b)=>a-b);rows.push({scenario,projectTypes,references,scope,medianMs:timings[7],p95Ms:timings[14],resultSha256:expectedHash});
  }
 }
 console.log(JSON.stringify({baseline:baseline??'working-tree',node:process.version,rounds,warmups,scope:'Actual diagnostic rule and complete analyzer; mock in-memory project metadata; fresh lists per round; metadata generation and assertions outside clock; full-fresh changes trailing comment to bypass parser cache; no Office/UI timing',rows},null,2));
}finally{if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
