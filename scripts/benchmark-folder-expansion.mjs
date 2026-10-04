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
const book='C:/work/App.vbp';api.workspace.workspaceFolders=[{uri:{fsPath:'C:/work'}}];api.workspace.findFiles=async()=>[{scheme:'file',fsPath:book}];api.window.showErrorMessage=(message)=>{throw Error(message);};
const cases=[['small',10,1,'last'],['wide',1000,1,'last'],['wider',3000,1,'last'],['deep-wide',1000,8,'last'],['deep',1,64,'last'],['wide-first',3000,1,'first'],['wide-missing',3000,1,'missing']];
const expansions=200,rows=[];
function prepare(width,depth,target){
 const prefix=Array.from({length:depth-1},(_,i)=>'Level'+i).join('.');
 const modules=Array.from({length:width},(_,i)=>({name:'M'+String(i).padStart(5,'0'),type:'standard',folder:(prefix?prefix+'.':'')+'F'+String(i).padStart(5,'0')}));
 const bridge={call:async(method)=>method==='listModules'?modules:method==='listSubs'?[]:{isPasswordProtected:false,isSigned:false}};
 const explorer=new api.ProjectExplorer(bridge);explorer.setView('folders');
 let project,folder;
 const expected=target==='missing'?[]:[modules[target==='first'?0:width-1].name];
 async function load(){[project]=await explorer.getChildren();await explorer.getChildren(project);folder=target==='missing'?{kind:'folder',label:'Absent',filePath:book,folder:'Absent'}:explorer.getFolderNode(book,modules[target==='first'?0:width-1].folder);assert.ok(folder,'Folder must be registered');}
 return{load,async warm(){const out=[];for(let i=0;i<expansions;i++)out.push(await explorer.getChildren(folder));return out;},async cold(){await load();return [await explorer.getChildren(folder)];},verify(out){for(const group of out)assert.deepEqual(group.map(node=>node.moduleName),expected);},dispose:()=>explorer.dispose()};
}
for(const [name,width,depth,target]of cases)for(const mode of name.startsWith('wide-')?['warm']:['warm','build']){
 const samples=[];for(let i=0;i<rounds+3;i++){const sample=prepare(width,depth,target);if(mode==='warm')await sample.load();const start=performance.now(),out=await(mode==='warm'?sample.warm():sample.cold());const elapsed=performance.now()-start;sample.verify(out);sample.dispose();if(i>=3)samples.push(elapsed);}samples.sort((a,b)=>a-b);
 rows.push({name:name+'-'+mode,width,depth,expansions:mode==='warm'?expansions:1,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,baseline:baseline??null,rounds,rows},null,2));
