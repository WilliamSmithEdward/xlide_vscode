import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??15);
if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3 to 100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-folder-expand-')),file=join(scratch,'provider.cjs');let api;
try{
 const plugins=[{name:'provider-host',setup(b){
  b.onResolve({filter:/^vscode$/},()=>({path:'vscode',namespace:'bench'}));
  b.onResolve({filter:/^vitest$/},()=>({path:'vitest',namespace:'bench'}));
  b.onLoad({filter:/.*/,namespace:'bench'},args=>args.path==='vitest'?{contents:'export const vi={fn:(impl)=>(...args)=>impl?.(...args)};',loader:'js'}:{contents:"import { vscodeMock } from "+JSON.stringify(join(root,'tests/helpers/vscodeMock.ts').replaceAll('\\','/'))+";const base=vscodeMock();export const {CancellationError,Disposable,EventEmitter,LanguageModelTextPart,LanguageModelToolResult,Location,MarkdownString,Position,Range,RelativePattern,TabInputCustom,TabInputText,TabInputTextDiff,ThemeIcon,ThemeColor,TreeItem,TextEdit,FileChangeType,FileType,TreeItemCollapsibleState,ViewColumn,FileSystemError,Uri,commands,env,languages,lm,window,workspace}=base;",loader:'js',resolveDir:root});
  if(baseline)b.onLoad({filter:/projectExplorer\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/projectExplorer.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));
 }}];
 const result=await build({plugins,stdin:{contents:"export { ProjectExplorer } from './src/projectExplorer';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,result.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(scratch);}

const rows=[];
for(const count of [5,100,1000,3000]){
 const subs=Object.freeze(Array.from({length:count},(_,i)=>Object.freeze({name:'P'+i,kind:'Sub',line:i+1})));
 for(const mode of ['cold','first','last','missing','jumps']){
  const samples=[];
  for(let round=-3;round<rounds;round++){
   const methods=[],book='C:/work/App.vbp';
   const tree=new api.ProjectExplorer({call:async(method)=>{methods.push(method);return method==='listModules'?[{name:'M',type:'standard'}]:method==='listSubs'?subs:{isPasswordProtected:false,isSigned:false};}});
   try{
    // Load the module through the provider before measuring its procedure rows.
    const project={kind:'project',label:'App.vbp',filePath:book};
    const [module]=await tree.getChildren(project);
    let nodes;if(mode!=='cold')nodes=await tree.getChildren(module);
    const queries=Array.from({length:mode==='cold'?1:200},(_,i)=>mode==='first'?'SUB P0':mode==='missing'?'Sub Absent':mode==='jumps'?'Sub P'+((i*317)%count):'Sub P'+(count-1));
    const results=[];const start=performance.now();
    if(mode==='cold')nodes=await tree.getChildren(module);
    else for(const query of queries)results.push(await tree.resolveProcedureNode(book,'M',query));
    const elapsed=performance.now()-start;
    assert.equal(nodes.length,count);
    if(mode!=='cold')for(let i=0;i<queries.length;i++)assert.equal(results[i],nodes.find(n=>n.kind==='sub'&&n.label.toLowerCase()===queries[i].toLowerCase()));
    assert.equal(methods.filter(m=>m==='listModules').length,1);
    assert.equal(methods.filter(m=>m==='listSubs').length,1);
    assert.equal(Object.isFrozen(subs),true);
    if(round>=0)samples.push(elapsed);
   }finally{tree.dispose();}
  }
  samples.sort((a,b)=>a-b);rows.push({count,mode,iterations:mode==='cold'?1:200,medianMs:samples[Math.floor(samples.length/2)],p95Ms:samples[Math.ceil(samples.length*.95)-1]});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0].model,baseline:baseline??null,rounds,warmups:3,scope:'Actual ProjectExplorer procedure expansion and resolveProcedureNode; mock bridge and VS Code, assertions outside clock; no Office or UI timing',rows},null,2));
