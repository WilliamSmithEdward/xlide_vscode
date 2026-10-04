import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync, existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';

// node scripts/benchmark-interface-refactor-preflight.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const fullReparse = process.argv.includes('--full-reparse');
const root = process.cwd();
const scratch = mkdtempSync(join(tmpdir(), 'xlide-interface-preflight-'));
const bundlePath = join(scratch, 'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(build) {
  build.onLoad({filter:/[\\/]refactor[\\/]implementInterface\.ts$/},args => ({ contents:execFileSync('git',['show',`${baseline}:src/analyzer/refactor/implementInterface.ts`],{encoding:'utf8'}), loader:'ts', resolveDir:dirname(args.path) }));
 } }] : [];
 const result = await build({ stdin:{contents:"export { implementInterface } from './src/analyzer/refactor/implementInterface'; export { parseModule } from './src/analyzer/parser/parseModule';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins });
 writeFileSync(bundlePath,result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundlePath);
} finally {
 if (existsSync(bundlePath)) unlinkSync(bundlePath);
 rmdirSync(scratch);
}
const interfaceSource = 'Public Sub Work(ByVal Input As Long)\nEnd Sub';
const rows = [];
for (const count of [1,1000,10000]) for (const mode of ['noImplements','ambiguous','wrongInterface','missingSource','emptyInterface','success','already','warmNoImplements']) {
 const samples = [];
 for (let round = -3; round < 9; round++) {
  const body = Array.from({length:count},(_,i)=>`Sub P${fullReparse ? mode + (round + 3) + "_" : ""}${i}()\nDim Local As Long\nEnd Sub\n`).join('')+"' "+mode+' round '+round+'\n';
  const prefix = mode === 'noImplements' || mode === 'warmNoImplements' ? '' : mode === 'ambiguous' ? 'Implements IJob\nImplements IOther\n' : 'Implements IJob\n';
  const source = prefix + body + (mode === 'already' ? 'Private Sub IJob_Work(ByVal Input As Long)\nEnd Sub\n' : '');
  const moduleSources = mode === 'missingSource' ? {} : { IJob: mode === 'emptyInterface' ? 'Private Sub Hidden()\nEnd Sub' : interfaceSource };
  const input = { source, moduleSources, ...(mode === 'wrongInterface' ? {interfaceName:'Other'} : {}) };
  const reason = {noImplements:'This class implements no interface. Add an `Implements` statement first.',warmNoImplements:'This class implements no interface. Add an `Implements` statement first.',ambiguous:'This class implements IJob, IOther. Say which one to implement.',wrongInterface:"This class does not implement 'Other'.",missingSource:"The project has no module called 'IJob'.",emptyInterface:"'IJob' has no public members to implement.",already:"'IJob' is already implemented in full."}[mode];
  const expected = mode === 'success' ? {ok:true,title:"Implement 1 member of 'IJob'",edits:[{span:{start:source.length,end:source.length},newText:"\nPrivate Sub IJob_Work(ByVal Input As Long)\n    Err.Raise 5 'TODO: implement this interface member\nEnd Sub\n"}]} : {ok:false,reason};
  if (mode === 'warmNoImplements') api.parseModule(source);
  const begin = performance.now();
  const result = api.implementInterface(input);
  const elapsed = performance.now()-begin;
  assert.deepEqual(result,expected);
  if (round >= 0) samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);
 rows.push({count,mode,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeResultsCorrect:true});
}
console.log(JSON.stringify({baseline:baseline??null,fullReparse,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete refactor results on fresh class source; construction excluded, class parse included except explicitly warmed no-Implements control. Default trailing-comment changes permit incremental parser reuse; --full-reparse changes every procedure name each round. Interface source stable. No editor command latency claim.',rows},null,2));
