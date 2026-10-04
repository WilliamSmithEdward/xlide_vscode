import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
const baseline = process.argv.find(a => a.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-late-bound-class-')), file = join(scratch, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) {
  b.onLoad({ filter: /[\\/]rules[\\/]lateBoundMembers\.ts$/ }, a => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/lateBoundMembers.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(a.path) }));
 } }] : [];
 const bundle = await build({ plugins, stdin: { contents: "export {checkRuntimeMemberNotFound} from './src/analyzer/diagnostics/rules/lateBoundMembers';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, bundle.outputFiles[0].contents);
 api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(scratch); } }
const rows = [], workCounts = [];
for (const kind of ['repeated-direct', 'repeated-collection', 'distinct-direct']) for (const count of [1, 100, 1000]) {
 let names = 0, memberNames = 0;
 const classes = Array.from({ length: count + 1 }, (_, i) => ({
  get name() { names++; return 'Class' + (i + 1); }, moduleName: 'Class' + (i + 1), kind: 'class', exhaustive: true,
  members: Array.from({ length: kind === 'distinct-direct' ? 1 : i === 0 ? count : 0 }, (_, j) => ({ get name() { memberNames++; return 'M' + j; }, moduleName: 'Class' + (i + 1), kind: 'property', returns: 'Long' })),
 }));
 const types = Array.from({ length: count }, (_, i) => 'Class' + (kind === 'distinct-direct' ? i + 1 : 1));
 const source = 'Option Explicit\n' + types.map((type, i) => {
  const body = kind === 'repeated-collection' ? `Dim c As New Collection\nc.Add New ${type}\nDebug.Print c(1).Nope` : `Dim actor As Object\nSet actor = New ${type}\nDebug.Print actor.Nope`;
  return `Sub P${i}()\n${body}\nEnd Sub\n`;
 }).join('');
 const mod = api.parseModule(source), symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 let from = 0;
 const expectedRule = types.map(type => {
  const receiver = kind === 'repeated-collection' ? 'c(1)' : 'actor', needle = kind === 'repeated-collection' ? 'c(1).Nope' : 'Nope';
  const start = source.indexOf(needle, from); from = start + needle.length;
  return ['runtimeMemberNotFound', `'${receiver}' holds a ${type} here, which has no member 'Nope'. This will raise Run-time error '438': Object doesn't support this property or method.`, { start, end: start + 4 }];
 });
 const runRule = () => { const out = []; api.checkRuntimeMemberNotFound(source, mod, symbols, { projectClassMembers: classes }, undefined, (...args) => out.push(args)); return out; };
 names = 0; memberNames = 0;
 assert.deepEqual(runRule(), expectedRule);
 workCounts.push({ kind, count, names, memberNames });
 if (!baseline) { assert.ok(names <= count * 4 + 10); assert.ok(memberNames <= count * 4 + 10); }
 // Instrumentation is only for the separate untimed work check.
 for (let i = 0; i < classes.length; i++) {
  Object.defineProperty(classes[i], 'name', { value: 'Class' + (i + 1), writable: true, configurable: true, enumerable: true });
  for (let j = 0; j < classes[i].members.length; j++) Object.defineProperty(classes[i].members[j], 'name', { value: 'M' + j, writable: true, configurable: true, enumerable: true });
 }
 const runModule = () => {
  const errors = [], diagnostics = api.analyzeModule(source, { projectClassMembers: classes, onInternalError: e => errors.push(String(e)) });
  assert.deepEqual(errors, []); return diagnostics;
 };
 const expectedModule = runModule();
 assert.equal(expectedModule.length, count);
 assert.ok(expectedModule.every(d => d.code === 'runtime-member-not-found'));
 for (const scope of ['rule', 'complete-module-diagnostics']) {
  const samples = [];
  for (let round = -3; round < 9; round++) {
   const start = performance.now(), actual = scope === 'rule' ? runRule() : runModule(), elapsed = performance.now() - start;
   assert.deepEqual(actual, scope === 'rule' ? expectedRule : expectedModule);
   if (round >= 0) samples.push(elapsed);
  }
  samples.sort((a, b) => a - b);
  rows.push({ kind, count, scope, medianMs: +samples[4].toFixed(5) });
 }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, workCounts, rows, scope: 'Parsed AST/symbols outside direct-rule timing; complete module includes analyzer preparation and error checks. Exact independently expected rule messages/spans and N expected module diagnostics; equality outside clocks. Plain timed metadata, untimed work counters. No editor/cold/heap claim.' }, null, 2));
