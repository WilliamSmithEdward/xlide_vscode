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
const scratch=mkdtempSync(join(tmpdir(),'xlide-caret-switch-')),file=join(scratch,'tracker.cjs');
const vscodeSource=String.raw`
export class EventEmitter {
 listeners=new Set();
 event=(callback)=>{this.listeners.add(callback);return {dispose:()=>this.listeners.delete(callback)};};
 fire(value){for(const callback of this.listeners)callback(value);}
 dispose(){this.listeners.clear();}
}
export const editorChanged=new EventEmitter(),selectionChanged=new EventEmitter();
export const window={activeTextEditor:undefined,onDidChangeActiveTextEditor:editorChanged.event,onDidChangeTextEditorSelection:selectionChanged.event};
`;
let api;
try {
 const plugins=[{name:'tracker-host',setup(b){
  b.onResolve({filter:/^vscode$/},()=>({path:'vscode',namespace:'bench'}));
  b.onLoad({filter:/.*/,namespace:'bench'},()=>({contents:vscodeSource,loader:'js'}));
  // Stub only document ownership. The real tracker, event handling and range scanner run.
  b.onLoad({filter:/vbaDocumentLocation.ts$/},()=>({contents:'export function moduleLocationOfDocument(doc){return doc.location;}',loader:'ts'}));
  if(baseline)b.onLoad({filter:/vbaCaretProcedure.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/vbaCaretProcedure.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));
 }}];
 const built=await build({plugins,stdin:{contents:"export { VbaCaretProcedureTracker } from './src/vbaCaretProcedure';export {vbaProcedureRanges,vbaProcedureAtLine,vbaProcedureLabel} from './src/vbaProcedureAtLine'; export { window, editorChanged, selectionChanged } from 'vscode';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
} finally {try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(scratch);}
const rows=[];
const source=n=>['Option Explicit',...Array.from({length:n},(_,i)=>'Sub P'+i+'()\nDebug.Print 1\nEnd Sub')].join('\n');
const cases=[...[2,1000,5000].flatMap(procedures=>['last','jumps','first','declarations'].map(mode=>({procedures,mode,steps:1000}))),{procedures:1000,mode:'edited-last',steps:200}];
for(const {procedures,mode,steps} of cases){
 const text=source(procedures),ranges=api.vbaProcedureRanges(text);
 const lines=Array.from({length:steps},(_,i)=>mode==='declarations'?0:mode==='first'?2:mode==='jumps'?((i*317)%procedures)*3+2:(procedures-1)*3+1+i%3);
 const expected=lines.map(line=>api.vbaProcedureLabel(api.vbaProcedureAtLine(ranges,line)));
 const samples=[];
 for(let round=0;round<rounds+3;round++){
  const document={version:1,location:{projectPath:'Book.xlsm',moduleName:'M',native:false},getText:()=>text};
  const editor={document,selection:{active:{line:lines[0]}}};api.window.activeTextEditor=editor;
  const tracker=new api.VbaCaretProcedureTracker();
  try{
   const out=[],start=performance.now();
   for(const line of lines){editor.selection.active.line=line;if(mode==='edited-last')document.version++;api.selectionChanged.fire({textEditor:editor});out.push(tracker.current.label);}
   const elapsed=performance.now()-start;assert.deepEqual(out,expected);
   if(round>=3)samples.push(elapsed);
  }finally{tracker.dispose();api.window.activeTextEditor=undefined;}
 }
 samples.sort((a,b)=>a-b);rows.push({procedures,mode,selectionEvents:steps,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,baseline:baseline??null,rounds,rows},null,2));
