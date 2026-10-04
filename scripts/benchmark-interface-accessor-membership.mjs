import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-interface-accessor-membership.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(),'xlide-interface-membership-'));
const path = join(scratch,'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(build){build.onLoad({filter:/[\\/]refactor[\\/]implementInterface\.ts$/},args=>({contents:execFileSync('git',['show',`${baseline}:src/analyzer/refactor/implementInterface.ts`],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}] : [];
 const built = await build({stdin:{contents:"export {implementInterface} from './src/analyzer/refactor/implementInterface'; export {parseModule} from './src/analyzer/parser/parseModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false,plugins});
 writeFileSync(path,built.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally { if(existsSync(path))unlinkSync(path);rmdirSync(scratch); }
const rows=[];
for(const count of [1,1000,10000])for(const mode of ['partial','none','complete']){
 const samples=[];
 for(let round=-3;round<9;round++){
  const prefix='Implements IData\n'+Array.from({length:count},(_,i)=>`Private Sub Noise${i}()\nEnd Sub\n`).join('');
  const getter='Private Property Get IData_Value() As Long\nEnd Property\n';
  const setter='Private Property Let IData_Value(ByVal RHS As Long)\nEnd Property\n';
  const source=prefix+(mode==='none'?'':getter)+(mode==='complete'?setter:'')+"' round "+round+'\n';
  const interfaceSource='Public Value As Long';
  api.parseModule(source);api.parseModule(interfaceSource);
  const stub=header=>'Private '+header+"\n    Err.Raise 5 'TODO: implement this interface member\nEnd Property";
  const expected=mode==='complete'||baseline&&mode==='partial'?{ok:false,reason:"'IData' is already implemented in full."}:{ok:true,title:`Implement ${mode==='none'?"2 members":"1 member"} of 'IData'`,edits:[{span:{start:source.length,end:source.length},newText:'\n'+(mode==='none'?stub('Property Get IData_Value() As Long')+'\n\n':'')+stub('Property Let IData_Value(ByVal RHS As Long)')+'\n'}]};
  const begin=performance.now();const results=Array.from({length:20},()=>api.implementInterface({source,moduleSources:{IData:interfaceSource}}));const elapsed=(performance.now()-begin)/20;
  for(const result of results)assert.deepEqual(result,expected);
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({count,mode,medianMs:+samples[4].toFixed(5),maxMs:+samples[8].toFixed(5),completeResultsChecked:true,baselineKnownWrongRefusal:!!baseline&&mode==='partial'});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,scope:'Complete warmed refactor calls; AST/source construction outside clock. Partial baseline refusal is known incorrect and independently checked. No cold-parse, heap-byte or editor-latency claim.',rows},null,2));
