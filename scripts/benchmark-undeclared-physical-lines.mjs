import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11),scratch=mkdtempSync(join(tmpdir(),'xlide-undeclared-lines-')),path=join(scratch,'bundle.cjs');
let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]diagnostics[\\/]rules[\\/]undeclared\.ts$/},a=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/rules/undeclared.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(a.path)}));}}]:[];
 const built=await build({plugins,stdin:{contents:"export {checkUndeclaredVariables} from './src/analyzer/diagnostics/rules/undeclared';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,built.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally {if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
const rows=[];
for(const prefixLines of [0,10000])for(const eol of ['\n','\r\n','\r']){
 const samples=[];
 for(let round=-3;round<9;round++){
  const prefix=Array(prefixLines).fill("' prior module text").join(eol)+(prefixLines?eol:'')+'Option Explicit'+eol+'Sub Go()'+eol;
  const source=prefix+'    missing = 1'+eol+'End Sub'+eol+"' round "+round,module=api.parseModule(source),symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:module});
  const query=()=>{const hits=[];api.checkUndeclaredVariables(source,module,symbols,undefined,new Set(),undefined,undefined,undefined,undefined,'standard',undefined,undefined,undefined,(...hit)=>hits.push(hit));return hits;};
  query();const begin=performance.now();const results=Array.from({length:10},query);const elapsed=(performance.now()-begin)/10;
  const wrong=!!baseline&&eol==='\r',at=wrong?0:prefix.length,fix={variableName:'missing',declaredType:'Long',edit:{span:{start:at,end:at},newText:(wrong?'':'    ')+'Dim missing As Long'+eol}};
  for(const hits of results){assert.equal(hits.length,1);assert.equal(hits[0][0],'undeclaredVariable');assert.deepEqual(hits[0][3].declareVariable,fix);assert.equal(api.applyVbaTextEdits(source,[fix.edit]),source.slice(0,at)+fix.edit.newText+source.slice(at));}
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({prefixLines,eol,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeFixDataChecked:true,baselineWrongPlacement:!!baseline&&eol==='\r'});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete warmed undeclared checker; AST/symbol construction outside clock. Complete independent fix data/count/rule checked afterward. CR baseline intentionally wrong placement, shared CR EOL retained. No full-analyzer, cold-parser, heap-byte or editor-latency claim.',rows},null,2));
