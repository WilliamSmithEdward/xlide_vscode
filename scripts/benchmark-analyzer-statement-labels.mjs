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
const changedFiles=new Set(['src/analyzer/diagnostics/rules/statementForms.ts']);
const dir=mkdtempSync(join(tmpdir(),'xlide-significant-consumers-')),bundle=join(dir,'api.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/\.ts$/},args=>{const file=relative(root,args.path).replaceAll('\\','/');if(!changedFiles.has(file))return;return {contents:execFileSync('git',['show',baseline+':'+file],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)};});}}]:[];
 const result=await build({plugins,stdin:{contents:"export { analyzeModule } from './src/analyzer/diagnostics/analyzeModule'; export { parseModule } from './src/analyzer/parser/parseModule'; export { tokenizeCached } from './src/analyzer/lexer/tokenize';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {checkStatementForms} from './src/analyzer/diagnostics/rules/statementForms';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(bundle,result.outputFiles[0].contents);api=createRequire(import.meta.url)(bundle);
}finally{try{unlinkSync(bundle);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(dir);}
if(process.argv.includes('--rule-only')) {
const src=(n,body)=>'Option Explicit\nSub P()\nDim total As Long\n'+Array.from({length:n},(_,i)=>body(i)).join('\n')+'\nEnd Sub';
const project={projectClassMembers:[{name:'Foo',moduleName:'Foo',kind:'standardModule',members:[]}]};
const ordinaryCases=[['assignments',src(1000,i=>'total='+i),{}],['long-body',src(5000,i=>'total='+i),{}],['project-no-hit',src(1000,i=>'total='+i),project],['qualified',src(1000,()=> 'total=Foo.Bar()'),project],['label-hit',src(200,()=> 'Foo: total=1\nGoTo Foo\ntotal=Foo()+Foo()'),project],['short',src(1,()=> 'total=1'),{}]];
const classes=process.argv.includes('--sub-member-controls')?Array.from({length:1000},(_,i)=>({name:i===0?'C':'Other'+i,moduleName:i===0?'C':'Other'+i,kind:'class',members:Array.from({length:100},(_,j)=>({name:j===0?'DoIt':'Value'+j,moduleName:i===0?'C':'Other'+i,kind:'method',sub:j===0,returns:j===0?undefined:'Long'}))})):[];
const subContext={projectClassMembers:classes};
const classSource=(n,name)=>'Option Explicit\nSub P()\nDim c As C\nDim result As Long\n'+Array.from({length:n},()=> 'result=c.'+name+'()').join('\n')+'\nEnd Sub';
const cases=process.argv.includes('--sub-member-controls')?[['large-project-negative-short',classSource(1,'Missing'),subContext],['large-project-negative-many',classSource(1000,'Missing'),subContext],['large-project-sub-short',classSource(1,'DoIt'),subContext],['large-project-sub-many',classSource(1000,'DoIt'),subContext]]:ordinaryCases;
let salt=0;const rows=[];
for(const [name,base,context] of cases)for(const mode of ['warm','fresh']){
 function prepare(){const source=mode==='fresh'?base+"\n' sample "+(++salt):base,mod=api.parseModule(source),symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod});return()=>{const diagnostics=[];api.checkStatementForms(source,mod,symbols,undefined,undefined,(code,message,span)=>diagnostics.push({code,message,span}),context);return diagnostics;};}
 const expected=JSON.stringify(prepare()());const samples=[];
 for(let i=0;i<rounds+3;i++){const run=prepare(),start=performance.now(),out=run(),ms=performance.now()-start;if(JSON.stringify(out)!==expected)throw Error('Changed rule output');if(i>=3)samples.push(ms);}
 samples.sort((a,b)=>a-b);rows.push({name:name+'-'+mode,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,rounds,rows},null,2));

} else {
const procs=(n,body)=>'Option Explicit\n'+Array.from({length:n},(_,i)=>'Sub P'+i+'(ByVal x As Long)\nDim total As Long\n'+body+'\nEnd Sub').join('\n');
const fixtures=[['many',procs(150,Array.from({length:12},(_,i)=>'total=x+'+i).join('\n')),'excel'],['body',procs(1,Array.from({length:5000},(_,i)=>'total=x+'+i).join('\n')),'excel'],['branches',procs(100,'Dim a() As Long\nDim c As Collection\nSet c=New Collection\nIf x>0 Then\nReDim a(2)\na(1)=x\nElse\ntotal=1\nEnd If\nFor total=0 To 2\nc.Add total\nNext\ntotal=c.Count'),'excel'],['types','Option Explicit\n'+Array.from({length:500},(_,i)=>'Sub P'+i+'()\nDim r As Excel.Range\nEnd Sub').join('\n'),'excel'],['loops',procs(100,'Dim a(2) As Long\nDim i As Long\nFor i=0 To 2\na(i)=i\nNext\nDo While total<3\ntotal=total+1\nLoop\nSelect Case total\nCase 1\ntotal=2\nCase Else\ntotal=3\nEnd Select'),'excel'],['short','Option Explicit\nSub P()\nDim n As Long\nn=1\nEnd Sub','excel'],['qualified',procs(100,'total=Foo.Bar()'),'excel'],['label-hit',procs(100,'Foo: total=1\nIf x>0 Then GoTo Foo\ntotal=Foo()+Foo()'),'excel']];
const rows=[];let salt=0;
for(const [name,base,host]of fixtures)for(const mode of ['warm','fresh']){
 let previousBody;
 const context=name==='qualified'||name==='label-hit'?{projectClassMembers:[{name:'Foo',moduleName:'Foo',kind:'standardModule',members:[]}]}:{};
 const reference=api.analyzeModule(base,{host,...context}).map(d=>[d.code,d.message,d.span]);
 function prepare(){const source=mode==='fresh'?base+"\n' sample "+(++salt):base;const mod=api.parseModule(source),body=mod.members.find(m=>m.kind==='Procedure')?.body;if(mode==='fresh'&&body===previousBody)throw Error('Reused parsed body');previousBody=body;api.tokenizeCached(source);const errors=[],options={host,...context,onInternalError:(e,w)=>errors.push([String(e),w])};return {run:()=>api.analyzeModule(source,options),verify:(diagnostics)=>{if(errors.length)throw Error(JSON.stringify(errors));if(JSON.stringify(diagnostics.map(d=>[d.code,d.message,d.span]))!==JSON.stringify(reference))throw Error('Changed diagnostics '+name);}};}
 for(let i=0;i<3;i++){const sample=prepare();sample.verify(sample.run());}const samples=[];for(let i=0;i<rounds;i++){const sample=prepare(),start=performance.now(),diagnostics=sample.run();samples.push(performance.now()-start);sample.verify(diagnostics);}samples.sort((a,b)=>a-b);rows.push({name:name+'-'+mode,diagnostics:reference.length,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));

}
