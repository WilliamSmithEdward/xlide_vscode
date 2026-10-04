import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??15);
if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3 to 100');
const changedFiles=new Set([
 "src/analyzer/diagnostics/calleeArguments.ts",
 "src/analyzer/diagnostics/dataflow.ts",
 "src/analyzer/diagnostics/heldObjects.ts",
 "src/analyzer/diagnostics/functionResults.ts",
 "src/analyzer/diagnostics/loopCounters.ts",
 "src/analyzer/diagnostics/straightLineValues.ts",
 "src/analyzer/diagnostics/rules/accessData.ts",
 "src/analyzer/diagnostics/rules/addressOfUse.ts",
 "src/analyzer/diagnostics/rules/arrays.ts",
 "src/analyzer/diagnostics/rules/byNameCalls.ts",
 "src/analyzer/diagnostics/rules/classInstanceValues.ts",
 "src/analyzer/diagnostics/rules/collectionState.ts",
 "src/analyzer/diagnostics/rules/conditionValues.ts",
 "src/analyzer/diagnostics/rules/callArity.ts",
 "src/analyzer/diagnostics/rules/controlFlow.ts",
 "src/analyzer/diagnostics/rules/deletedSettings.ts",
 "src/analyzer/diagnostics/rules/deletedObjects.ts",
 "src/analyzer/diagnostics/rules/declares.ts",
 "src/analyzer/diagnostics/rules/declarations.ts",
 "src/analyzer/diagnostics/rules/documentNames.ts",
 "src/analyzer/diagnostics/rules/excelSessionState.ts",
 "src/analyzer/diagnostics/rules/errorValues.ts",
 "src/analyzer/diagnostics/rules/dictionaryState.ts",
 "src/analyzer/diagnostics/rules/filePaths.ts",
 "src/analyzer/diagnostics/rules/handlerFlow.ts",
 "src/analyzer/diagnostics/rules/fileStatements.ts",
 "src/analyzer/diagnostics/rules/expressions.ts",
 "src/analyzer/diagnostics/rules/formContents.ts",
 "src/analyzer/diagnostics/rules/hostArguments.ts",
 "src/analyzer/diagnostics/rules/lateBoundMembers.ts",
 "src/analyzer/diagnostics/rules/lockedArrays.ts",
 "src/analyzer/diagnostics/rules/lateBoundObjects.ts",
 "src/analyzer/diagnostics/rules/malformedLines.ts",
 "src/analyzer/diagnostics/rules/moduleMembers.ts",
 "src/analyzer/diagnostics/rules/objectState.ts",
 "src/analyzer/diagnostics/rules/omittedArguments.ts",
 "src/analyzer/diagnostics/rules/objectValues.ts",
 "src/analyzer/diagnostics/rules/parentheses.ts",
 "src/analyzer/diagnostics/rules/overflow.ts",
 "src/analyzer/diagnostics/rules/refusedDeclarations.ts",
 "src/analyzer/diagnostics/rules/paramArrayUse.ts",
 "src/analyzer/diagnostics/rules/runtimeValues.ts",
 "src/analyzer/diagnostics/rules/propertyUse.ts",
 "src/analyzer/diagnostics/rules/statementTypes.ts",
 "src/analyzer/diagnostics/rules/typeMembers.ts",
 "src/analyzer/diagnostics/rules/variantValues.ts"
]);
const dir=mkdtempSync(join(tmpdir(),'xlide-significant-consumers-')),bundle=join(dir,'api.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/\.ts$/},args=>{const file=relative(root,args.path).replaceAll('\\','/');if(!changedFiles.has(file))return;return {contents:execFileSync('git',['show',baseline+':'+file],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)};});}}]:[];
 const result=await build({plugins,stdin:{contents:"export { analyzeModule } from './src/analyzer/diagnostics/analyzeModule'; export { parseModule } from './src/analyzer/parser/parseModule'; export { tokenizeCached } from './src/analyzer/lexer/tokenize';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(bundle,result.outputFiles[0].contents);api=createRequire(import.meta.url)(bundle);
}finally{try{unlinkSync(bundle);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(dir);}
const procs=(n,body)=>'Option Explicit\n'+Array.from({length:n},(_,i)=>'Sub P'+i+'(ByVal x As Long)\nDim total As Long\n'+body+'\nEnd Sub').join('\n');
const fixtures=[['many',procs(150,Array.from({length:12},(_,i)=>'total=x+'+i).join('\n')),'excel'],['body',procs(1,Array.from({length:5000},(_,i)=>'total=x+'+i).join('\n')),'excel'],['branches',procs(100,'Dim a() As Long\nDim c As Collection\nSet c=New Collection\nIf x>0 Then\nReDim a(2)\na(1)=x\nElse\ntotal=1\nEnd If\nFor total=0 To 2\nc.Add total\nNext\ntotal=c.Count'),'excel'],['types','Option Explicit\n'+Array.from({length:500},(_,i)=>'Sub P'+i+'()\nDim r As Excel.Range\nEnd Sub').join('\n'),'excel'],['loops',procs(100,'Dim a(2) As Long\nDim i As Long\nFor i=0 To 2\na(i)=i\nNext\nDo While total<3\ntotal=total+1\nLoop\nSelect Case total\nCase 1\ntotal=2\nCase Else\ntotal=3\nEnd Select'),'excel'],['short','Option Explicit\nSub P()\nDim n As Long\nn=1\nEnd Sub','excel']];
const rows=[];let salt=0;
for(const [name,base,host]of fixtures)for(const mode of ['warm','fresh']){
 let previousBody;
 const reference=api.analyzeModule(base,{host}).map(d=>[d.code,d.message,d.span]);
 function prepare(){const source=mode==='fresh'?base+"\n' sample "+(++salt):base;const mod=api.parseModule(source),body=mod.members.find(m=>m.kind==='Procedure')?.body;if(mode==='fresh'&&body===previousBody)throw Error('Reused parsed body');previousBody=body;api.tokenizeCached(source);const errors=[],options={host,onInternalError:(e,w)=>errors.push([String(e),w])};return {run:()=>api.analyzeModule(source,options),verify:(diagnostics)=>{if(errors.length)throw Error(JSON.stringify(errors));if(JSON.stringify(diagnostics.map(d=>[d.code,d.message,d.span]))!==JSON.stringify(reference))throw Error('Changed diagnostics '+name);}};}
 for(let i=0;i<3;i++){const sample=prepare();sample.verify(sample.run());}const samples=[];for(let i=0;i<rounds;i++){const sample=prepare(),start=performance.now(),diagnostics=sample.run();samples.push(performance.now()-start);sample.verify(diagnostics);}samples.sort((a,b)=>a-b);rows.push({name:name+'-'+mode,diagnostics:reference.length,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
