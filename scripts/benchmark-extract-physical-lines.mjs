import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-extract-physical-lines.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(),'xlide-extract-lines-'));
const path = join(scratch,'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(build){build.onLoad({filter:/(?:[\\/]refactor[\\/]extractVariable|[\\/]vbaSourceScan)\.ts$/},args=>({contents:execFileSync('git',['show',`${baseline}:${args.path.endsWith('vbaSourceScan.ts')?'src/vbaSourceScan.ts':'src/analyzer/refactor/extractVariable.ts'}`],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}] : [];
 const built = await build({stdin:{contents:"export {extractVariable} from './src/analyzer/refactor/extractVariable'; export {parseModule} from './src/analyzer/parser/parseModule'; export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 writeFileSync(path,built.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally { if(existsSync(path))unlinkSync(path);rmdirSync(scratch); }
const rows=[];
for(const prefixLines of [0,10000])for(const eol of ['\n','\r\n','\r']){
 const samples=[];
 for(let round=-3;round<9;round++){
  const prefix=Array(prefixLines).fill("' earlier module line").join(eol)+(prefixLines?eol:'')+'Sub Go()'+eol;
  const source=prefix+'    Debug.Print 2 * 3'+eol+'End Sub'+eol+"' round "+round,start=source.indexOf('2 * 3');
  const input={source,name:'extracted',span:{start,end:start+5}};api.parseModule(source);
  const expected=baseline&&eol==='\r'?'Dim extracted As Double\nextracted = 2 * 3\n'+source.replace('2 * 3','extracted'):prefix+'    Dim extracted As Double'+eol+'    extracted = 2 * 3'+eol+'    Debug.Print extracted'+eol+'End Sub'+eol+"' round "+round;
  const begin=performance.now();const results=Array.from({length:10},()=>api.extractVariable(input));const elapsed=(performance.now()-begin)/10;
  for(const result of results){assert.equal(result.ok,true);assert.equal(result.title,"Extract 'extracted'");const applied=api.applyVbaTextEdits(source,result.edits);assert.equal(applied,expected);assert.equal(applied.slice(result.renameSpan.start,result.renameSpan.end),'extracted');assert.equal(result.edits[0].span.start,baseline&&eol==='\r'?0:prefix.length);}
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({prefixLines,eol,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeOutputsChecked:true,baselineWrongPlacement:!!baseline&&eol==='\r'});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete warmed extraction with explicit supplied name; source and AST warming outside clock. CR baseline independently checks known wrong module-start insertion/LF. LF/CRLF equal-output controls. No cold-parser, heap-byte or editor-latency claim.',rows},null,2));
