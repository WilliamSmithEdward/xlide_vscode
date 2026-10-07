import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-conditional-expression-depth.mjs [--baseline=COMMIT]
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-conditional-depth-'));
const path = join(scratch, 'bundle.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) {
  b.onLoad({ filter: /[\\/]conditional[\\/]conditionalCompilation\.ts$/ }, args => ({ contents: execFileSync('git', ['show', `${baseline}:src/analyzer/conditional/conditionalCompilation.ts`], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) }));
 } }] : [];
 const built = await build({ stdin: { contents: "export {evaluateConditionalExpression} from './src/analyzer/conditional/conditionalCompilation';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false, plugins });
 writeFileSync(path, built.outputFiles[0].contents); api = createRequire(import.meta.url)(path);
} finally { if (existsSync(path)) unlinkSync(path); rmdirSync(scratch); }
const rows = [];
for (const kind of ['not', 'sign', 'parentheses']) for (const count of [1, 100, 20000]) {
 const expression = kind === 'not' ? 'Not '.repeat(count) + '1' : kind === 'sign' ? '- '.repeat(count) + '1' : '('.repeat(count) + '1' + ')'.repeat(count);
 const expected = kind === 'parentheses' && count > 256 ? undefined : 1;
 const samples = []; let errors = 0;
 for (let round = -3; round < 9; round++) {
  const begin = performance.now(); let result, error;
  try { result = api.evaluateConditionalExpression(expression); } catch (e) { error = e; }
  const elapsed = performance.now() - begin;
  if (error) { assert.ok(baseline && count === 20000); assert.equal(error.name, 'RangeError'); errors++; }
  else assert.equal(result, count === 1 && kind === 'not' ? -2 : count === 1 && kind === 'sign' ? -1 : expected);
  if (round >= 0) samples.push(elapsed);
 }
 samples.sort((a,b) => a-b); rows.push({kind, count, medianMs: +samples[4].toFixed(5), errors, calls: 12, outputsChecked: true});
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, scope: 'Complete expression evaluation including tokenization. Deep baseline failures timed as failures, never compared as successful throughput. No editor, full analyzer or heap claim.', rows }, null, 2));
