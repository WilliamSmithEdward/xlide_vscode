import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-interface-header-lines.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-interface-header-'));
const bundlePath = join(scratch, 'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(build) {
  build.onLoad({filter:/[\\/]refactor[\\/]implementInterface\.ts$/},args=>({contents:execFileSync('git',['show',`${baseline}:src/analyzer/refactor/implementInterface.ts`],{encoding:'utf8'}),resolveDir:dirname(args.path),loader:'ts'}));
 }}] : [];
 const bundle = await build({stdin:{contents:"export {implementInterface} from './src/analyzer/refactor/implementInterface'; export {parseModule} from './src/analyzer/parser/parseModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 writeFileSync(bundlePath,bundle.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundlePath);
} finally { if (existsSync(bundlePath)) unlinkSync(bundlePath); rmdirSync(scratch); }
const rows = [];
for (const bodyLines of [0,100,10000]) for (const eol of ['\n','\r\n','\r']) {
 const samples = [];
 let generatedCharacters;
 for (let round = -3; round < 9; round++) {
  const source = 'Implements IJob\n';
  const header = 'Public Sub Work(ByVal Input As Long)';
  const interfaceSource = [header,...Array.from({length:bodyLines},()=> '    Debug.Print "private body"'),'End Sub',"' round "+round].join(eol);
  // Header processing is measured with both ASTs warmed; source/parse outside.
  api.parseModule(source); api.parseModule(interfaceSource);
  const copied = baseline && eol === '\r' ? [header,...Array.from({length:bodyLines},()=> '    Debug.Print "private body"'),'End Sub'].join(eol) : header;
  const expected = {ok:true,title:"Implement 1 member of 'IJob'",edits:[{span:{start:source.length,end:source.length},newText:'\nPrivate '+copied.replace('Public ','').replace('Sub Work','Sub IJob_Work')+"\n    Err.Raise 5 'TODO: implement this interface member\nEnd Sub\n"}]};
  const begin = performance.now();
  const results = Array.from({length:20},()=>api.implementInterface({source,moduleSources:{IJob:interfaceSource}}));
  const elapsed = (performance.now()-begin)/20;
  for (const result of results) assert.deepEqual(result,expected);
  generatedCharacters = expected.edits[0].newText.length;
  if (round >= 0) samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);
 rows.push({bodyLines,eol,generatedCharacters,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),resultChecked:true,baselineKnownCorruption:!!baseline&&eol==='\r'});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete interface-refactor calls; AST/source construction outside clock. CR baseline output intentionally includes known body corruption, independently checked. No cold parse, byte allocation or editor-latency claim.',rows},null,2));
