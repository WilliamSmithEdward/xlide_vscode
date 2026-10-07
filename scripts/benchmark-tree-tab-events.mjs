import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { readFileSync, mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
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
const normalized=p=>(baseline?execFileSync('git',['show',baseline+':'+p],{cwd:root,encoding:'utf8'}):readFileSync(join(root,p),'utf8')).replace(/\r\n/g,'\n');
const extension=normalized('src/extension.ts');
const callback=extension.match(/modulesClosedBy: (\(event\) => modulesWithNoTabLeft\([\s\S]*?\n\s+\)),/)[1];
const location=normalized('src/vbaDocumentLocation.ts');
function declaration(name){const start=location.indexOf('export function '+name+'(');if(start<0)throw Error('Missing '+name);const end=location.indexOf('\n}',start);return location.slice(start,end+2);}
// The exact production resolver/callback run; ownership decoding is excluded.
const facade="import * as vscode from 'vscode';\nconst moduleLocationOfUri=uri=>uri.toString().startsWith('xlide-vba:')?{projectPath:'Book.xlsm',moduleName:'M',native:false}:undefined; const projectIdentityKey=x=>x;const moduleIdentityKey=x=>x;\n"+declaration('tabUris')+'\n'+declaration('modulesWithNoTabLeft')+'\nexport const productionCallback='+callback+';';
const host=String.raw`
const disposable=()=>({dispose(){}}),noopEvent=()=>disposable();
export class TabInputText {constructor(uri){this.uri=uri;}}
export class TabInputCustom {constructor(uri){this.uri=uri;}}
export class TabInputTextDiff {constructor(original,modified){this.original=original;this.modified=modified;}}
export let changed;
export const window={visibleTextEditors:[],onDidChangeActiveTextEditor:noopEvent,tabGroups:{all:[],onDidChangeTabs(callback){changed=callback;return disposable();}}};
export const workspace={onDidChangeConfiguration:noopEvent};
`;
const scratch=mkdtempSync(join(tmpdir(),'xlide-tree-tabs-')),file=join(scratch,'follow.cjs');let api;
try{
 const plugins=[{name:'tab-event-host',setup(b){
  b.onResolve({filter:/^vscode$/},()=>({path:'vscode',namespace:'bench'}));
  b.onResolve({filter:/^tab-resolver$/},()=>({path:'resolver',namespace:'bench'}));
  b.onLoad({filter:/.*/,namespace:'bench'},args=>({contents:args.path==='vscode'?host:facade,loader:args.path==='vscode'?'js':'ts'}));
  if(baseline)b.onLoad({filter:/explorerFollow\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/explorerFollow.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));
 }}];
 const result=await build({plugins,stdin:{contents:"export { ExplorerFollow } from './src/explorerFollow'; export * as host from 'vscode'; export { productionCallback } from 'tab-resolver';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,result.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(scratch);}
const events=200,rows=[];
function prepare(count,n,kind){
 let groupReads=0,uriReads=0,resolverCalls=0,folds=0;const outcomes=[];
 const tabs=Array.from({length:n},(_,i)=>({input:new api.host.TabInputText({toString:count?()=>{uriReads++;return 'file:/notes'+i+'.txt';}:()=> 'file:/notes'+i+'.txt'})}));
 const groups=Array.from({length:Math.min(n,4)},(_,i)=>({tabs:tabs.filter((_,j)=>j%Math.min(n,4)===i)}));
 Object.defineProperty(api.host.window.tabGroups,'all',{configurable:true,get:count?()=>{groupReads++;return groups;}:()=>groups});
 const noop=()=>({dispose(){}});
 const explorer={onDidReplaceRows:noop,clearActiveModule(project,module){assert.equal(project,'Book.xlsm');assert.equal(module,'M');folds++;}};
 const treeView={visible:true,selection:[],onDidChangeVisibility:noop,onDidChangeSelection:noop,onDidExpandElement:noop,onDidCollapseElement:noop};
 const callback=count?(event)=>{resolverCalls++;const result=api.productionCallback(event);outcomes.push(result);return result;}:(event)=>{const result=api.productionCallback(event);outcomes.push(result);return result;};
 const follow=new api.ExplorerFollow({explorer,treeView,caret:{current:undefined,onDidChange:noop},enabled:()=>true,modulesClosedBy:callback});
 const tab={input:new api.host.TabInputText({toString:()=> kind==='module-close'?'xlide-vba:/Book.xlsm/M':'file:/closed-notes.txt'})};
 const event={closed:kind.endsWith('close')?[tab]:[],opened:kind==='opened'?[tab]:[],changed:kind==='changed'?[tabs[0]]:[]};
 return{run(){for(let i=0;i<events;i++)api.host.changed(event);},verify(){for(const out of outcomes)assert.deepEqual(out,kind==='module-close'?[{projectPath:'Book.xlsm',moduleName:'M',native:false}]:[]);assert.equal(folds,kind==='module-close'?events:0);},counts:()=>({resolverCalls,groupReads,uriReads,folds}),dispose:()=>follow.dispose()};
}
for(const kind of ['changed','opened','empty','nonmodule-close','module-close'])for(const n of [1,50,1000]){
 const counted=prepare(true,n,kind);counted.run();counted.verify();const counts=counted.counts();assert.equal(counts.uriReads,counts.groupReads*n,'Every enumerated text tab must contribute its URI');if(kind.endsWith('close'))assert.equal(counts.resolverCalls,events,'Actual closure must still run');counted.dispose();
 const samples=[];for(let i=0;i<rounds+3;i++){const sample=prepare(false,n,kind),start=performance.now();sample.run();const elapsed=performance.now()-start;sample.verify();sample.dispose();if(i>=3)samples.push(elapsed);}samples.sort((a,b)=>a-b);
 rows.push({name:kind+'-'+n,tabs:n,events,...counts,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,baseline:baseline??null,rounds,rows},null,2));
