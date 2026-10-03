// Run: node scripts/benchmark-text-fold-depth.mjs [--baseline=<commit>]
// Exposes the private fold only inside a temporary benchmark bundle.
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-text-fold-depth-'));
const bundle = join(scratch, 'api.cjs');
let api;
try {
 const result = await build({ plugins: [{ name: 'fold-benchmark', setup(builder) {
  builder.onLoad({ filter: /(?:assignments|conditionValue)\.ts$/ }, ({ path }) => {
   const relative = path.slice(root.length + 1).replaceAll('\\', '/');
   let contents = baseline ? execFileSync('git', ['show', baseline + ':' + relative], { encoding: 'utf8' }) : readFileSync(path, 'utf8');
   if (path.endsWith('assignments.ts')) contents += '\nexport { spelledText as benchmarkSpelledText };';
   return { contents, loader: 'ts', resolveDir: dirname(path) };
  });
 }}], stdin: { contents: "export { benchmarkSpelledText } from './src/analyzer/diagnostics/rules/assignments'; export { rawExpressionTokens } from './src/analyzer/diagnostics/walker'; export { numberValue } from './src/analyzer/diagnostics/conditionValue'; export { analyzeModule } from './src/analyzer/diagnostics/analyzeModule';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(bundle, result.outputFiles[0].contents); api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch (e) { if (e.code !== 'ENOENT') throw e; } rmdirSync(scratch); }
const rows = [];
for (const depth of [10,100,255,5000]) for (const fn of ['CStr','StrConv','Abs']) {
 const text = fn === 'StrConv' ? 'StrConv('.repeat(depth) + '"bad"' + ', 1)'.repeat(depth) : (fn + '(').repeat(depth) + (fn === 'Abs' ? '1' : '"bad"') + ')'.repeat(depth);
 const toks = api.rawExpressionTokens(text), facts = { value: () => undefined };
 const scope = { callableShadows: new Set(), runtimeShadows: new Set() };
 const run = () => {
  try {
   const value = fn === 'Abs' ? api.numberValue(toks, facts) : api.benchmarkSpelledText(toks, () => undefined, new Map(), scope);
   if (depth < 256 && value === undefined) throw Error('Unexpected unknown fold');
   return value === undefined ? 'unknown' : 'folded';
  } catch (e) { if (e instanceof RangeError) return 'stack-overflow'; throw e; }
 };
 for (let i = 0; i < 3; i++) run();
 const samples = [], outcomes = new Set();
 for (let i = 0; i < 15; i++) { const start = performance.now(); outcomes.add(run()); samples.push(performance.now() - start); }
 samples.sort((a,b) => a-b);
 rows.push({ depth, fn, outcomes: [...outcomes], medianMs: +samples[7].toFixed(3), p95Ms: +samples[14].toFixed(3) });
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, warmups:3, rounds:15, baseline, rows }, null, 2));
