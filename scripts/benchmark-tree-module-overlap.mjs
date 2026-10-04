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
 const result=await build({plugins,stdin:{contents:"export { ProjectExplorer } from './src/projectExplorer';export { workspace,window } from 'vscode';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,result.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(scratch);}
const book='C:/work/App.vbp';
api.workspace.workspaceFolders=[{uri:{fsPath:'C:/work'}}];
api.workspace.findFiles=async()=>[{scheme:'file',fsPath:book}];
api.window.showErrorMessage=(message)=>{throw Error(message);};
const rows=[];
for(const count of [50,1000,3000])for(const [mode,callers] of [['cold',1],['cold',2],['cold',8],['warm',2]]){
 const modules=Object.freeze(Array.from({length:count},(_,i)=>Object.freeze({name:'M'+String((i*317)%count).padStart(5,'0'),type:['standard','class','userform'][i%3]})));
 // Sorting outside the clock defines the expected output independently of the
 // provider under comparison. No comparator counters run inside timing samples.
 const order={userform:1,standard:2,class:3};
 const expected=[...modules].sort((a,b)=>order[a.type]-order[b.type]||a.name.localeCompare(b.name)).map(m=>m.name);
 const samples=[];
 for(let round=0;round<rounds+3;round++){
  let release;let calls=0;const pending=new Promise(r=>{release=r;});
  const explorer=new api.ProjectExplorer({call:method=>{assert.equal(method,'listModules');calls++;return pending;}});
  try{
   const [project]=await explorer.getChildren();
   if(mode==='warm'){const ready=explorer.getChildren(project);release(modules);await ready;}
   const start=performance.now();
   const expansions=[explorer.getChildren(project)];
   const follows=Array.from({length:callers-1},()=>explorer.resolveModuleNode(book,modules[0].name));
   // Let tab-follow traverse the cached project root before releasing the RPC.
   // This is JavaScript work only: no Office, VS Code renderer or network delay.
   await Promise.resolve();await Promise.resolve();release(modules);
   const [children,nodes]=await Promise.all([Promise.all(expansions),Promise.all(follows)]);
   const elapsed=performance.now()-start;
   assert.deepEqual(children[0].map(n=>n.moduleName),expected);
   for(const node of nodes)assert.equal(node,children[0].find(n=>n.moduleName===modules[0].name));
   assert.equal(calls,1);assert.deepEqual(modules.map(m=>m.name),Array.from({length:count},(_,i)=>'M'+String((i*317)%count).padStart(5,'0')));
   if(round>=3)samples.push(elapsed);
  }finally{explorer.dispose();}
 }
 samples.sort((a,b)=>a-b);
 rows.push({modules:count,mode,callers,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,baseline:baseline??null,rounds,rows},null,2));
