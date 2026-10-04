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
 const result=await build({plugins,stdin:{contents:"export { ProjectExplorer } from './src/projectExplorer';export { workspace,window } from 'vscode';export {presentAgentModuleWrite,keepAgentChange} from './src/xlideAgentDiff';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,result.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(scratch);}
const book='C:/work/App.vbp';
api.workspace.workspaceFolders=[{uri:{fsPath:'C:/work'}}];
api.workspace.findFiles=async()=>[{scheme:'file',fsPath:book}];
const renders=200,rows=[];
const cases=[...[50,1000].flatMap(count=>['none','first','last','all'].map(pending=>({count,pending}))),{count:3000,pending:'none'},{count:3000,pending:'last'}];
for(const {count,pending} of cases){
 const modules=Array.from({length:count},(_,i)=>({name:'M'+String(i).padStart(5,'0'),type:'standard',folder:'F'+String(i).padStart(5,'0')}));
 const samples=[];
 for(let round=0;round<rounds+3;round++){
  const explorer=new api.ProjectExplorer({call:async()=>modules});explorer.setView('folders');
  const registered=pending==='none'?[]:pending==='all'?modules.map(m=>m.name):[modules[pending==='first'?0:count-1].name];
  try{
   const [project]=await explorer.getChildren(),folders=await explorer.getChildren(project);
   for(const name of registered)await api.presentAgentModuleWrite(book,name.toLowerCase(),{before:'Sub Old()\nEnd Sub',beforeExisted:true,after:'Sub New()\nEnd Sub'});
   const start=performance.now(),items=[];
   for(let i=0;i<renders;i++)items.push(explorer.getTreeItem(folders[i%count]));
   const elapsed=performance.now()-start;
   for(let i=0;i<items.length;i++){
    const expected=registered.includes(modules[i%count].name);
    assert.equal(items[i].description.includes('agent edit'),expected);
    assert.equal(items[i].iconPath.color?.id,expected?'xlide.agentEdit.foreground':undefined);
    assert.equal(items[i].resourceUri,undefined);
   }
   if(round>=3)samples.push(elapsed);
  }finally{for(const name of registered)api.keepAgentChange(book,name);explorer.dispose();}
 }
 samples.sort((a,b)=>a-b);rows.push({modules:count,pending,renders,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,baseline:baseline??null,rounds,rows},null,2));
