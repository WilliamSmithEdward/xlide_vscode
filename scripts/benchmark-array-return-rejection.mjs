// Run: node scripts/benchmark-array-return-rejection.mjs [--baseline=<commit>] [--rounds=15]
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { readFileSync, mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.slice(9) ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw Error('rounds must be 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-array-return-rejection-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
 const plugins = [{ name: 'summary', setup(builder) {
  builder.onLoad({ filter: /rules[\\/]assignments\.ts$/ }, () => ({ contents: (baseline ? execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/assignments.ts'], { encoding: 'utf8' }) : readFileSync(join(root, 'src/analyzer/diagnostics/rules/assignments.ts'), 'utf8')) + '\nexport { arrayOnlyVariantFunctions };', loader: 'ts', resolveDir: join(root, 'src/analyzer/diagnostics/rules') }));
 }}];
 const result = await build({ plugins, stdin: { contents: "export { hostObjectModelForToken } from './src/analyzer/host/hostRegistry'; export { arrayOnlyVariantFunctions, checkAssignmentTypes } from './src/analyzer/diagnostics/rules/assignments'; export { forEachStatementWithHeaders } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(bundle, result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; } rmdirSync(scratch); }
const rows=[];
for(const count of [100,1000,10000])for(const mode of ['early-scalar','early-self-read','late-scalar','only-arrays']) {
 const middle='Debug.Print n\n'.repeat(count);
 const first=mode==='early-scalar'?'F = 1\n':mode==='early-self-read'?'Debug.Print F\n':'F = Array(1)\n';
 const source='Function F() As Variant\nDim n As Long\n'+first+middle+(mode==='late-scalar'?'F = 1\n':'')+'End Function\nSub P()\nDim a() As Long\na = F()\nEnd Sub';
 const mod=api.parseModule(source);
 for(const surface of ['summary','rule']) {
  let previousRoot;
  function prepare(){const symbols=api.buildModuleSymbols('M','standard',source,{parsedModule:mod});if(symbols.root===previousRoot)throw Error('Reused binding root');previousRoot=symbols.root;return ()=>{if(surface==='summary'){const result=api.arrayOnlyVariantFunctions(source,mod);if(result.has('f')!==(mode==='only-arrays'))throw Error('Wrong summary');}else{let hits=0;api.checkAssignmentTypes(source,mod,symbols,undefined,{},undefined,()=>hits++);if(hits!==(mode==='only-arrays'?1:0))throw Error('Wrong findings '+hits+' '+mode);}};}
  for(let i=0;i<3;i++)prepare()();const samples=[];
  for(let i=0;i<rounds;i++){const run=prepare(),start=performance.now();run();samples.push(performance.now()-start);}samples.sort((a,b)=>a-b);
  rows.push({name:count+'-'+mode+'-'+surface,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
