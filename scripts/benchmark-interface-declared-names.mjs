import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-interface-declared-names.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(),'xlide-interface-names-'));
const path = join(scratch,'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(build){build.onLoad({filter:/[\\/]refactor[\\/]implementInterface\.ts$/},args=>({contents:execFileSync('git',['show',`${baseline}:src/analyzer/refactor/implementInterface.ts`],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}] : [];
 const built = await build({stdin:{contents:"export {implementInterface} from './src/analyzer/refactor/implementInterface'; export {parseModule} from './src/analyzer/parser/parseModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 writeFileSync(path,built.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally { if(existsSync(path))unlinkSync(path);rmdirSync(scratch); }
const stub=(header,closer)=>'Private '+header+"\n    Err.Raise 5 'TODO: implement this interface member\nEnd "+closer;
const rows=[];
for(const count of [1,1000])for(const mode of ['ascii','unicode','bracketed','fields']){
 const samples=[];
 for(let round=-3;round<9;round++){
  const iface=[],expectedStubs=[];
  for(let i=0;i<count;i++){
   const name=mode==='unicode'?`Δοκιμή${i}ζ`:mode==='bracketed'?`[Value ${i}]`:`Work${i}`;
   const renamed=baseline&&(mode==='unicode'||mode==='bracketed')?name:mode==='bracketed'?`[IJob_Value ${i}]`:'IJob_'+name;
   if(mode==='fields'){
    iface.push(`Public ${name} As Long`);
    expectedStubs.push(stub(`Property Get IJob_${name}() As Long`,'Property'),stub(`Property Let IJob_${name}(ByVal RHS As Long)`,'Property'));
   }else{
    iface.push(`Public Sub ${name}(ByRef Arg As Long)`,'End Sub');
    expectedStubs.push(stub(`Sub ${renamed}(ByRef Arg As Long)`,'Sub'));
   }
  }
  const source='Implements IJob\n',interfaceSource=iface.join('\n')+"\n' round "+round;
  api.parseModule(source);api.parseModule(interfaceSource);
  const members=mode==='fields'?count*2:count;
  const expected={ok:true,title:`Implement ${members} member${members===1?'':'s'} of 'IJob'`,edits:[{span:{start:source.length,end:source.length},newText:'\n'+expectedStubs.join('\n\n')+'\n'}]};
  const begin=performance.now();const results=Array.from({length:10},()=>api.implementInterface({source,moduleSources:{IJob:interfaceSource}}));const elapsed=(performance.now()-begin)/10;
  for(const result of results)assert.deepEqual(result,expected);
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({count,mode,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeResultsChecked:true,baselineKnownUnprefixed:!!baseline&&['unicode','bracketed'].includes(mode)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete warmed refactor calls, AST/source construction outside clock. Unicode/bracketed baseline outputs intentionally check the known wrong names; ASCII and field controls have equal outputs. No cold-parser, heap-byte or editor-latency claim.',rows},null,2));
