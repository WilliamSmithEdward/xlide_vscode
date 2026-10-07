import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir, cpus } from 'node:os';
import { join } from 'node:path';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';

// Run independently in baseline/current/current/baseline order.
// node scripts/benchmark-hover-procedure-signatures.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const root = process.cwd();
const scratch = mkdtempSync(join(tmpdir(), 'xlide-hover-signature-'));
let api;
try {
 const plugins = baseline ? [{name:'baseline-hover',setup(build) {
  build.onLoad({filter:/[\\/]hover[\\/]resolveHover\.ts$/},()=>({contents:execFileSync('git',['show',`${baseline}:src/analyzer/hover/resolveHover.ts`],{encoding:'utf8'}),loader:'ts'}));
 }}] : [];
 const result = await build({stdin:{contents:"export { resolveHover } from './src/analyzer/hover/resolveHover'; export { editorModuleSymbols } from './src/analyzer/symbols/editorModuleSymbols';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 const path = join(scratch,'bundle.cjs');
 writeFileSync(path,result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(path);
} finally { rmSync(scratch,{recursive:true,force:true}); }
const rows = [];
for (const count of [0,100,10000]) for (const mode of ['warmProcedure','firstProcedure','warmLocal']) {
 const samples = [];
 for (let round = -3; round < 9; round++) {
  const source = 'Private Function Target(ByVal Input As Long) As String\n'+Array.from({length:count},(_,i)=>`Dim Local${i} As Long\n`).join('')+'Dim HoverValue As Long\nDebug.Print hovervalue\nEnd Function\n'+"' round "+round;
  const start = mode === 'warmLocal' ? source.lastIndexOf('hovervalue') : source.indexOf('Target');
  const offset = start+2;
  const expected = mode === 'warmLocal' ? {signature:'HoverValue As Long',details:['Local in Target'],span:{start,end:start+10}} : {signature:'Function Target(Input As Long) As String',details:['Declared in Module: Module','Visibility: Private'],span:{start,end:start+6}};
  api.editorModuleSymbols('Module','standard',source);
  // Warm tokenizer + hover index outside cold-signature measurement using a local.
  const localStart = source.lastIndexOf('hovervalue');
  assert.deepEqual(api.resolveHover(source,localStart+2),{signature:'HoverValue As Long',details:['Local in Target'],span:{start:localStart,end:localStart+10}});
  if (mode !== 'firstProcedure') assert.deepEqual(api.resolveHover(source,offset),expected);
  const requests = mode === 'firstProcedure' ? 1 : 100;
  const begin = performance.now();
  const results = Array.from({length:requests},()=>api.resolveHover(source,offset));
  const elapsed = (performance.now()-begin)/requests;
  for (const result of results) assert.deepEqual(result,expected);
  if (round >= 0) samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);
 rows.push({count,mode,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeResultsCorrect:true});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete hover queries; source/symbol/token/index construction outside clock. FirstProcedure includes first signature construction only. No editor latency or cold lexer claim.',rows},null,2));
