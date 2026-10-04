import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const baseline = process.argv.find(s => s.startsWith('--baseline='))?.slice(11);
const dir = mkdtempSync(join(tmpdir(), 'xlide-extract-order-')), file = join(dir, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) {
  b.onLoad({ filter: /[\\/]refactor[\\/]extractMethod\.ts$/ }, p => ({
   contents: execFileSync('git', ['show', baseline + ':src/analyzer/refactor/extractMethod.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(p.path),
  }));
 }}] : [];
 const b = await build({ plugins, stdin: { contents: "export {extractMethod} from './src/analyzer/refactor/extractMethod';export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, b.outputFiles[0].contents); api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(dir); } }
const rows = [], workCounts = [];
for (const placement of ['small', 'medium-prefix', 'large-prefix', 'common-short', 'large-suffix', 'long-name-limit']) {
 for (const n of placement === 'long-name-limit' ? [400] : [1, 10, 64, 256, 1000]) {
  const names = Array.from({ length: n }, (_, i) => placement === 'common-short' && i === 0 ? 'a' : 'local' + (placement === 'long-name-limit' ? 'x'.repeat(200) : '') + i.toString().padStart(4, '0'));
  const order = names.map((_, i) => names[(i * 37) % n]);
  const notes = Array.from({ length: Math.ceil(n / 3) }, (_, i) => "' " + order.slice(i * 3, i * 3 + 3).join(' ')).join('\n') + '\n';
  const padding = Array(placement === 'medium-prefix' ? 1000 : 10000).fill("' " + (placement === 'common-short' ? 'a' : 'q').repeat(90)).join('\n') + '\n';
  const prefix = (['small', 'large-suffix'].includes(placement) ? '' : padding) + notes + 'Option Explicit\nSub Main()\n' + names.map(name => 'Dim ' + name + ' As Long').join('\n') + '\n';
  const body = names.map(name => name + ' = 1').join('\n');
  const source = prefix + body + '\nEnd Sub\n' + (placement === 'large-suffix' ? padding : '');
  assert(source.split('\n').every(line => line.length < 1023));
  const input = { source, span: { start: prefix.length, end: prefix.length + body.length }, name: 'Work' };
  const result = api.extractMethod(input); assert.equal(result.ok, true);
  const wantedOrder = [...names].sort((a, b) => source.indexOf(a) - source.indexOf(b));
  const helper = 'Private Sub Work()\n' + wantedOrder.map(name => 'Dim ' + name + ' As Long').join('\n') + '\n' + body + '\nEnd Sub';
  assert.equal(result.title, "Extract 'Work'");
  assert.deepEqual(result.renameSpan, { start: prefix.length, end: prefix.length + 4 });
  assert.equal(result.edits.find(e => e.span.start === input.span.start && e.span.end === input.span.end)?.newText, 'Work');
  assert(result.edits.some(e => e.newText === '\n\n' + helper + '\n'));
  const fullExpected = { result, source: api.applyVbaTextEdits(source, result.edits) }; assert(fullExpected.source.includes(helper));
  const wanted = new Set(names), originalIndexOf = String.prototype.indexOf; let searches = 0;
  try {
   String.prototype.indexOf = function (needle, position) { if (String(this) === source && wanted.has(needle) && position === undefined) searches++; return originalIndexOf.call(this, needle, position); };
   assert.deepEqual(api.extractMethod(input), result);
  } finally { String.prototype.indexOf = originalIndexOf; }
  const batched = (['large-prefix', 'common-short'].includes(placement) && n >= 64)
   || (placement === 'medium-prefix' && n >= 256) || (['small', 'large-suffix'].includes(placement) && n === 1000);
  assert.equal(searches, baseline ? n === 1 ? 0 : n : batched ? 0 : n === 1 ? 0 : n);
  workCounts.push({ placement, n, independentSourceSearches: searches });
  const digest = createHash('sha256').update(JSON.stringify(fullExpected)).digest('hex');
  const extraction = () => api.extractMethod(input);
  const applied = () => { const result = extraction(); assert.equal(result.ok, true); return { result, source: api.applyVbaTextEdits(source, result.edits) }; };
  for (const [scope, run, expected] of [['public-extraction', extraction, result], ['extraction-and-apply-edits', applied, fullExpected]]) {
   const samples = [];
   for (let round = -3; round < 9; round++) {
    const start = performance.now(), actual = run(), elapsed = performance.now() - start; assert.deepEqual(actual, expected);
    if (round >= 0) samples.push(elapsed);
   }
   samples.sort((a, b) => a - b);
   rows.push({ placement, n, sourceCodeUnits: source.length, scope, medianMs: +samples[4].toFixed(5), outputDigest: digest });
  }
 }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, workCounts, rows, scope: 'Moved locals produce a zero-parameter helper; all input physical lines stay below 1023 characters. Independent exact primary edit/header/order/rename assertions and full result plus applied-text snapshots. Work counters untimed and restored. Public API parse/lookup cache behavior included; application scope additionally applies the edits. Repeated-source trials, not native/editor, cold-start, heap, or universal speedup evidence.' }, null, 2));
