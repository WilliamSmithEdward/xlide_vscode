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
 const built=await build({plugins,stdin:{contents:"export { VbaCaretProcedureTracker } from './src/vbaCaretProcedure'; export { window, editorChanged } from 'vscode';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
} finally {try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(scratch);}
const source=(procedures)=>['Option Explicit',...Array.from({length:procedures},(_,i)=>['Sub P'+i+'()',...Array.from({length:20},()=> '    Debug.Print 12345'),'End Sub'].join('\r\n'))].join('\r\n');
const small=source(2),large=source(1000);
const fixtures=[['single-small',1,small,false],['two-small',2,small,false],['two-large',2,large,false],['eight-large',8,large,false],['edited-two-large',2,large,true]];
const steps=200,rows=[];
function prepare(count,text,n,edited){
 let reads=0;
 const line=text.split(/\r\n/).length-1;
 const docs=Array.from({length:n},(_,i)=>({uri:{toString:()=> 'xlide-vba:/Book.xlsm/M'+i+'.bas'},version:1,location:{projectPath:'Book.xlsm',moduleName:'M'+i,native:false},getText:count?()=>{reads++;return text;}:()=>text}));
 const editors=docs.map(document=>({document,selection:{active:{line}}}));
 api.window.activeTextEditor=undefined;
 const tracker=new api.VbaCaretProcedureTracker();
 const activate=(editor)=>{api.window.activeTextEditor=editor;api.editorChanged.fire(editor);};
 // Both implementations see each document before measurement; only a version change forces a rescan.
 for(const editor of editors)activate(editor);
 return {run(){const out=[];for(let i=0;i<steps;i++){const editor=editors[i%n];if(edited)editor.document.version++;activate(editor);out.push([tracker.current.moduleName,tracker.current.label]);}return out;},verify(out){assert.equal(out.length,steps);for(let i=0;i<steps;i++)assert.deepEqual(out[i],['M'+i%n,'Sub P'+(text===large?999:1)]);},reads:()=>reads,dispose:()=>tracker.dispose()};
}
for(const [name,n,text,edited]of fixtures){
 const counted=prepare(true,text,n,edited);counted.verify(counted.run());const reads=counted.reads();counted.dispose();
 const samples=[];
 for(let i=0;i<rounds+3;i++){const sample=prepare(false,text,n,edited),start=performance.now(),out=sample.run(),elapsed=performance.now()-start;sample.verify(out);sample.dispose();if(i>=3)samples.push(elapsed);}
 samples.sort((a,b)=>a-b);rows.push({name,documents:n,sourceLines:text.split(/\r\n/).length,sourceBytes:Buffer.byteLength(text),switchEvents:steps,readsIncludingPreparation:reads,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,baseline:baseline??null,rounds,rows},null,2));
