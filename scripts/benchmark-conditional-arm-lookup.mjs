import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-conditional-arm-lookup.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const conditionalPath = 'src/analyzer/conditional/conditionalCompilation.ts';
const scratch = mkdtempSync(join(tmpdir(), 'xlide-conditional-arms-'));
const paths = [join(scratch, 'api.cjs'), join(scratch, 'work.cjs')];
let api, work;
try {
 const plugins = baseline ? [{name: 'baseline', setup(b) { b.onLoad({filter: /[\\/]conditional[\\/]conditionalCompilation\.ts$/}, args => ({contents: execFileSync('git', ['show', `${baseline}:${conditionalPath}`], {encoding: 'utf8'}), loader: 'ts', resolveDir: dirname(args.path)})); }}] : [];
 const result = await build({plugins, stdin: {contents: "export {createConditionalActivityTracker} from './src/analyzer/conditional/conditionalCompilation';export {parseModule} from './src/analyzer/parser/parseModule';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: 'ts'}, bundle: true, platform: 'node', format: 'cjs', write: false});
 writeFileSync(paths[0], result.outputFiles[0].contents); api = createRequire(import.meta.url)(paths[0]);
 // Expose the private helper only in this throwaway measurement bundle.
 const source = baseline ? execFileSync('git', ['show', `${baseline}:${conditionalPath}`], {encoding: 'utf8'}) : readFileSync(conditionalPath, 'utf8');
 const counted = await build({stdin: {contents: source + '\nexport {armsDiverge as compareArms};', resolveDir: join(process.cwd(), dirname(conditionalPath)), loader: 'ts'}, bundle: true, platform: 'node', format: 'cjs', write: false});
 writeFileSync(paths[1], counted.outputFiles[0].contents); work = createRequire(import.meta.url)(paths[1]);
} finally { for (const path of paths) if (existsSync(path)) unlinkSync(path); rmdirSync(scratch); }
const rows = [], workCounts = [];
for (const depth of [1, 100, 1000]) {
 let reads = 0;
 const chain = (start) => {let parent; for (let i = 0; i < depth; i++) {const id = start+i;parent = {get chain(){reads++;return id;}, index: 0, parent};} return parent;};
 const a = chain(0), b = chain(depth);
 assert.equal(work.compareArms(a,b), false);
 if (baseline) assert.equal(reads, 2*depth*depth); else assert.ok(reads <= 4*depth);
 workCounts.push({depth, chainPropertyReads: reads, exclusive: false});
 for (const relation of ['independent', 'alternative', 'shared-parent']) {
  const block = (name,count) => '#If UNKNOWN Then\n'.repeat(depth) + Array(count).fill('Public '+name+' As Long\n').join('') + '#End If\n'.repeat(depth);
  const source = relation === 'independent' ? block('FirstValue',1)+block('SecondValue',1) : relation === 'alternative' ? '#If OUTER Then\n'+block('FirstValue',1)+'#Else\n'+block('SecondValue',1)+'#End If\n' : '#If OUTER Then\n'+block('FirstValue',1)+block('SecondValue',1)+'#End If\n';
  const tracker = api.createConditionalActivityTracker(api.parseModule(source));
  const span = (name) => {const start=source.indexOf(name);return {start,end:start+name.length};};
  const first=span('FirstValue'),second=span('SecondValue'),samples=[];
  for (let round=-3;round<9;round++) {const t=performance.now();let actual;for(let i=0;i<100;i++)actual=tracker.mutuallyExclusive(first,second);const elapsed=(performance.now()-t)/100;assert.equal(actual,relation==='alternative');assert.equal(tracker.mutuallyExclusive(second,first),relation==='alternative');if(round>=0)samples.push(elapsed);}
  samples.sort((a,b)=>a-b);rows.push({scope:'tracker-query',depth,relation,medianMs:+samples[4].toFixed(5),outputsChecked:true});
 }
 const block = (count) => '#If UNKNOWN Then\n'.repeat(depth)+Array(count).fill('Public RepeatedValue As Long\n').join('')+'#End If\n'.repeat(depth);
 const source = 'Option Explicit\n'+block(1)+block(32),declared=[];
 let at=0;while((at=source.indexOf('RepeatedValue',at))!==-1){declared.push(at);at++;}
 const expected=declared.slice(1).map(start=>({code:'duplicate-module-variable',message:"Duplicate declaration: 'RepeatedValue' is already declared at module level.",severity:'error',span:{start,end:start+13},specReference:'MS-VBAL 5.2.3',origin:'run'}));
 const samples=[];api.parseModule(source);
 for(let round=-3;round<9;round++){const failures=[];const begin=performance.now(),actual=api.analyzeModule(source,{onInternalError:e=>failures.push(String(e))}),elapsed=performance.now()-begin;assert.deepEqual(failures,[]);assert.deepEqual(actual,expected);if(round>=0)samples.push(elapsed);}
 samples.sort((a,b)=>a-b);rows.push({scope:'complete-module-diagnostics',depth,duplicates:32,medianMs:+samples[4].toFixed(5),completeOutputsChecked:true});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,workCounts,rows,scope:'Tracker query setup and module AST warming outside clock. Complete analyzeModule diagnostics and failure checks are independently verified. Private helper work reads measured in a temporary export-only bundle, with no production instrumentation. No cold-parser, retained heap, worker or renderer claim.'},null,2));
