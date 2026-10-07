import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-extract-declared-names.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(),'xlide-extract-declared-'));
const path = join(scratch,'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(build){build.onLoad({filter:/[\\/]refactor[\\/]extractVariable\.ts$/},args=>({contents:execFileSync('git',['show',`${baseline}:src/analyzer/refactor/extractVariable.ts`],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}] : [];
 const built = await build({stdin:{contents:"export {extractVariable} from './src/analyzer/refactor/extractVariable'; export {parseModule} from './src/analyzer/parser/parseModule'; export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 writeFileSync(path,built.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally { if(existsSync(path))unlinkSync(path);rmdirSync(scratch); }
const rows=[];
for(const count of [1,1000])for(const mode of ['bare','single','grouped','trivia','redim']){
 const samples=[];
 for(let round=-3;round<9;round++){
  const body=Array.from({length:count},(_,i)=>mode==='single'?'    Dim prior'+i+' As Long':mode==='grouped'?'    Dim prior'+i+' As Long, extra'+i+' As Long':mode==='trivia'?"    ' Dim value As Long "+'x'.repeat(100):mode==='redim'?'    ReDim prior'+i+'(1, 2)':'    Debug.Print '+i).join('\n')+'\n';
  const collision=['single','grouped','redim'].includes(mode)?'    '+(mode==='redim'?'ReDim value(3)':mode==='grouped'?'Dim another As Long, value As Long':'Dim value As Long')+'\n':'';
  const prefix='Sub Go()\n'+body+collision,source=prefix+'    Debug.Print 2 * 3\nEnd Sub\n'+"' round "+round;
  const input={source,span:{start:source.indexOf('2 * 3'),end:source.indexOf('2 * 3')+5}};
  const name=mode==='single'||mode==='redim'||mode==='trivia'&&baseline||mode==='grouped'&&!baseline?'value2':'value';
  api.parseModule(source);
  const begin=performance.now();const results=Array.from({length:10},()=>api.extractVariable(input));const elapsed=(performance.now()-begin)/10;
  for(const result of results){assert.equal(result.ok,true);assert.equal(result.title,"Extract '"+name+"'");assert.equal(api.applyVbaTextEdits(source,result.edits),prefix+'    Dim '+name+' As Double\n    '+name+' = 2 * 3\n    Debug.Print '+name+'\nEnd Sub\n'+"' round "+round);const applied=api.applyVbaTextEdits(source,result.edits);assert.equal(applied.slice(result.renameSpan.start,result.renameSpan.end),name);}
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({count,mode,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeOutputsChecked:true,baselineKnownWrongName:!!baseline&&['grouped','trivia'].includes(mode)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete warmed Extract Variable calls; source and AST warming outside clock. Independent output/title/rename controls after clock. Grouped/trivia baseline names intentionally wrong. No cold-parser, heap-byte or editor-latency claim.',rows},null,2));
