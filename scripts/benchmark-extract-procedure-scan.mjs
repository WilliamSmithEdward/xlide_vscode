import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { readFileSync, mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const baseline = process.argv.find(s => s.startsWith('--baseline='))?.slice(11);
async function load(instrument) {
 const dir = mkdtempSync(join(tmpdir(), 'xlide-extract-procedure-')), file = join(dir, 'api.cjs');
 try {
  const plugins = [{ name: 'audit', setup(b) {
   if (baseline) b.onLoad({ filter: /[\\/]refactor[\\/]extractMethod\.ts$/ }, p => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/refactor/extractMethod.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(p.path) }));
   if (instrument) b.onLoad({ filter: /[\\/]vbaSourceScan\.ts$/ }, p => {
    let source = readFileSync(p.path, 'utf8');
    const anchor = 'for (let i = firstLine; i <= lastLine; i++) {';
    assert.equal(source.split(anchor).length, 2);
    source = source.replace(anchor, anchor + '\nauditIdentifierLines++;') + '\nexport let auditIdentifierLines = 0; export function resetAuditIdentifierLines() { auditIdentifierLines = 0; }';
    return { contents: source, loader: 'ts', resolveDir: dirname(p.path) };
   });
  }}];
  const built = await build({ plugins, stdin: { contents: "export {extractMethod} from './src/analyzer/refactor/extractMethod';export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';" + (instrument ? "export {auditIdentifierLines,resetAuditIdentifierLines} from './src/vbaSourceScan';" : ''), resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
  writeFileSync(file, built.outputFiles[0].contents); return createRequire(import.meta.url)(file);
 } finally { try { unlinkSync(file); } finally { rmdirSync(dir); } }
}
const api = await load(false), counted = await load(true), rows = [];
for (const count of [0, 100, 1000, 10000]) for (const placement of ['before', 'after']) for (const kind of ['Sub', 'Function']) {
 const padding = Array.from({ length: Math.ceil(count / 500) }, (_, i) => 'Sub Other' + i + '()\nDim x As Long\n' + Array(Math.min(500, count - i * 500)).fill('x = x + 1').join('\n') + '\nEnd Sub\n').join('');
 const prefix = 'Option Explicit\n' + (placement === 'before' ? padding : '') + (kind === 'Sub' ? 'Sub Main()' : 'Function Main() As Long') + '\nDim x As Long\nx = 3\n';
 const body = 'Debug.Print x', source = prefix + body + '\n' + (kind === 'Function' ? 'Main = x\n' : '') + 'End ' + kind + '\n' + (placement === 'after' ? padding : '');
 assert(source.split('\n').every(line => line.length < 1023));
 const input = { source, span: { start: prefix.length, end: prefix.length + body.length }, name: 'Work' }, result = api.extractMethod(input);
 assert.equal(result.ok, true);
 const parameter = 'ByVal x As Long';
 const edited = api.applyVbaTextEdits(source, result.edits);
 assert(edited.includes('Private Sub Work(' + parameter + ')\n' + body + '\nEnd Sub'));
 assert(edited.includes('Work x'));
 counted.resetAuditIdentifierLines(); assert.deepEqual(counted.extractMethod(input), result);
 const identifierLines = counted.auditIdentifierLines;
 assert.equal(identifierLines, baseline ? source.split('\n').length + (kind === 'Function' ? 1 : 0) : (kind === 'Sub' ? 5 : 7));
 const digest = createHash('sha256').update(JSON.stringify({ result, edited })).digest('hex');
 for (const scope of ['public-extraction', 'extraction-and-apply-edits']) {
  const expected = scope === 'public-extraction' ? result : { result, edited }, samples = [];
  for (let i = -3; i < 9; i++) {
   const start = performance.now(), actual = api.extractMethod(input), value = scope === 'public-extraction' ? actual : { result: actual, edited: api.applyVbaTextEdits(source, actual.edits) }, elapsed = performance.now() - start;
   assert.deepEqual(value, expected); if (i >= 0) samples.push(elapsed);
  }
  samples.sort((a,b) => a-b);
  rows.push({ count, placement, kind, sourceCodeUnits: source.length, identifierLines, scope, medianMs: +samples[4].toFixed(5), outputDigest: digest });
 }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, rows, scope: 'Repeated-source public API and optional edit application. Instrumented identifier-line counters use a separate untimed bundle. Independently asserted helper and invocation, complete output digest. No native/editor, cold-start, memory, or universal speedup claim.' }, null, 2));
