import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-refactor-edit-chunks.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(),'xlide-edit-chunks-'));
const path = join(scratch,'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(build){build.onLoad({filter:/[\\/]refactor[\\/]refactorTypes\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/refactor/refactorTypes.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}] : [];
 const built = await build({stdin:{contents:"export {extractMethod} from './src/analyzer/refactor/extractMethod'; export {parseModule} from './src/analyzer/parser/parseModule'; export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 writeFileSync(path,built.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally { if(existsSync(path))unlinkSync(path);rmdirSync(scratch); }
const rows=[];
for(const count of [1,1000,10000]){
 const source='x'.repeat(count*64),edits=Array.from({length:count},(_,i)=>({span:{start:i*64,end:i*64+1},newText:'Y'})),expected=('Y'+'x'.repeat(63)).repeat(count),samples=[];
 for(let round=-3;round<9;round++){const begin=performance.now();const actual=api.applyVbaTextEdits(source,edits),elapsed=performance.now()-begin;assert.equal(actual,expected);if(round>=0)samples.push(elapsed);}
 samples.sort((a,b)=>a-b);rows.push({kind:'applyEdits',count,sourceChars:source.length,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeOutputChecked:true});
}
for(const count of [1,100,1000]){
 const samples=[];
 for(let round=-3;round<9;round++){
  const prefix='Option Explicit\nSub Go()\n',selection=Array.from({length:count},(_,i)=>'Dim local'+i+' As Long, unused'+i+' As Long\nlocal'+i+' = '+i).join('\n');
  const source=prefix+selection+'\nDebug.Print "after"\nEnd Sub\n',input={source,span:{start:prefix.length,end:prefix.length+selection.length},name:'Work'};
  api.parseModule(source);
  const begin=performance.now();const result=api.extractMethod(input),elapsed=performance.now()-begin;assert.equal(result.ok,true);assert.equal(result.title,"Extract 'Work'");
  const expected=prefix+Array.from({length:count},(_,i)=>'Dim unused'+i+' As Long').join('\n')+'\nWork\nDebug.Print "after"\nEnd Sub\n\nPrivate Sub Work()\n'+Array.from({length:count},(_,i)=>'Dim local'+i+' As Long\nlocal'+i+' = '+i).join('\n')+'\nEnd Sub\n\n';
  assert.equal(api.applyVbaTextEdits(source,result.edits),expected);
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({kind:'extractMethod',count,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeOutputChecked:true});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Direct complete edit application and complete warmed Extract Method with grouped local declarations. All outputs independently checked after timing. Source/AST construction outside refactor clock. No cold-parser, heap-byte or editor-latency claim.',rows},null,2));
