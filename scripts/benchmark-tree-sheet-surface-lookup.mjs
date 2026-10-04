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
const warmExpansions=Number(process.argv.find(a=>a.startsWith('--warm-expansions='))?.slice(18)??1);
if(!Number.isInteger(warmExpansions)||warmExpansions<1||warmExpansions>500)throw Error('warm-expansions must be 1 to 500');
const scratch=mkdtempSync(join(tmpdir(),'xlide-folder-expand-')),file=join(scratch,'provider.cjs');let api;
try{
 const plugins=[{name:'provider-host',setup(b){
  b.onResolve({filter:/^vscode$/},()=>({path:'vscode',namespace:'bench'}));
  b.onResolve({filter:/^vitest$/},()=>({path:'vitest',namespace:'bench'}));
  b.onLoad({filter:/.*/,namespace:'bench'},args=>args.path==='vitest'?{contents:'export const vi={fn:(impl)=>(...args)=>impl?.(...args)};',loader:'js'}:{contents:"import { vscodeMock } from "+JSON.stringify(join(root,'tests/helpers/vscodeMock.ts').replaceAll('\\','/'))+";const base=vscodeMock();export const {CancellationError,Disposable,EventEmitter,LanguageModelTextPart,LanguageModelToolResult,Location,MarkdownString,Position,Range,RelativePattern,TabInputCustom,TabInputText,TabInputTextDiff,ThemeIcon,ThemeColor,TreeItem,TextEdit,FileChangeType,FileType,TreeItemCollapsibleState,ViewColumn,FileSystemError,Uri,commands,env,languages,lm,window,workspace}=base;",loader:'js',resolveDir:root});
  if(baseline)b.onLoad({filter:/shapeRows\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/shapeRows.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));
 }}];
 const result=await build({plugins,stdin:{contents:"export { ShapeRows } from './src/shapeRows';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(file,result.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(scratch);}

const rows=[];
for(const [name,count,mode] of [['small',5,'ordered'],['medium',100,'ordered'],['large',1000,'ordered'],['larger',3000,'ordered'],['missing',1000,'missing'],['reversed',1000,'reversed']]){
 const sheets=Array.from({length:count},(_,i)=>Object.freeze({name:'Sheet'+i,kind:'worksheet'}));
 const surfaces=sheets.map((sheet,i)=>Object.freeze({surface:mode==='missing'?'Absent'+i:sheet.name,shapes:Object.freeze([{name:'Box',kind:'shape'}])}));
 if(mode==='reversed')surfaces.reverse();Object.freeze(surfaces);Object.freeze(sheets);
 const project={kind:'project',label:'Book',filePath:'C:/work/Book.xlsm'};
 for(const phase of ['cold','warm']){
  const samples=[];let signature;
  for(let round=-3;round<rounds;round++){
   const calls=[];const tree=new api.ShapeRows({call:async(method)=>{calls.push(method);return method==='listShapes'?{surfaces}:{sheets};}},()=>{});
   const {folders:[folder]}=await tree.projectRows(project,[]);
   if(phase==='warm')await tree.children(folder,async()=>[]);
   let result;const iterations=phase==='warm'?20:1;const start=performance.now();
   for(let i=0;i<iterations;i++)result=await tree.children(folder,async()=>[]);
   const elapsed=performance.now()-start;
   assert.deepEqual(calls,['listWorkbookSheets','listShapes']);
   const actual=mode==='missing'?await tree.children(result[0],async()=>[]):result;
   assert.deepEqual(actual.map(n=>n.label),sheets.map(s=>s.name));
   assert.equal(actual.every(n=>tree.parentOf(n)===(mode==='missing'?result[0]:folder)),true);
   const current=actual.map(n=>[n.kind,n.label,n.itemCount??null]);
   if(signature)assert.deepEqual(current,signature);else signature=current;
   if(round>=0)samples.push(elapsed);
  }
  samples.sort((a,b)=>a-b);rows.push({name,count,phase,expansions:phase==='warm'?20:1,medianMs:samples[Math.floor(samples.length/2)],p95Ms:samples[Math.ceil(samples.length*.95)-1],signature});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0].model,baseline:baseline??null,rounds,warmups:3,scope:'Actual ShapeRows.children, mock bridge and VS Code host; no Office or UI timing; assertions outside clock',rows},null,2));
