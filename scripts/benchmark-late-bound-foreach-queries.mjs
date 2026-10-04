import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
const baseline = process.argv.find(a => a.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-late-bound-foreach-')), file = join(scratch, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) {
  b.onLoad({ filter: /[\\/]diagnostics[\\/](?:rules[\\/]lateBoundMembers|typeInference)\.ts$/ }, a => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/' + (a.path.endsWith('typeInference.ts') ? 'typeInference.ts' : 'rules/lateBoundMembers.ts')], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(a.path) }));
 } }] : [];
 const bundle = await build({ plugins, stdin: { contents: "export {checkRuntimeMemberNotFound} from './src/analyzer/diagnostics/rules/lateBoundMembers';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, bundle.outputFiles[0].contents);
 api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(scratch); } }
const rows = [], workCounts = [];
for (const kind of ['held-same', 'host-loops', 'shared-repeated', 'shared-distinct', 'direct-repeated', 'reverse-distinct', 'generic-collection']) for (const count of [1, 100, 1000]) {
 let names = 0, interfaces = 0, implementedNames = 0;
 const implementsLists = Array.from({ length: count + 2 }, (_, i) => {
  if ((kind === 'direct-repeated' || kind === 'generic-collection') && i === 1) return Array.from({ length: count }, (_, n) => n === count - 1 ? kind === 'generic-collection' ? 'Collection' : 'Class1' : 'Extra' + n);
  if (kind === 'reverse-distinct' && i === 0) return Array.from({ length: count }, (_, n) => 'Class' + (n + 2));
  return i === count + 1 ? kind === 'shared-repeated' ? ['Class1', 'Class2'] : kind === 'shared-distinct' ? Array.from({ length: count + 1 }, (_, n) => 'Class' + (n + 1)) : [] : [];
 });
 const countedLists = implementsLists.map(list => new Proxy(list, { get(target, key, receiver) { if (typeof key === 'string' && /^(0|[1-9]\d*)$/.test(key)) implementedNames++; return Reflect.get(target, key, receiver); } }));
 const classes = Array.from({ length: count + 2 }, (_, i) => ({ get name() { names++; return 'Class' + (i + 1); }, moduleName: 'Class' + (i + 1), kind: 'class', exhaustive: true, members: [], get implements() { interfaces++; return countedLists[i]; } }));
 const source = 'Option Explicit\n' + (kind === 'host-loops'
  ? Array.from({ length: count }, (_, i) => `Sub P${i}()\nDim actor As Class1\nFor Each actor In Worksheets\nDebug.Print 1\nNext\nEnd Sub\n`).join('')
  : 'Sub Go()\nDim c As New Collection\nDim actor As ' + (kind === 'generic-collection' ? 'Collection' : 'Class1') + '\n' + Array.from({ length: count }, (_, i) => 'c.Add New ' + (kind === 'shared-distinct' || kind === 'reverse-distinct' ? 'Class' + (i + 2) : kind === 'shared-repeated' || kind === 'direct-repeated' || kind === 'generic-collection' ? 'Class2' : 'Class1') + '\n').join('') + 'For Each actor In c\nDebug.Print 1\nNext\nEnd Sub\n');
 const mod = api.parseModule(source), symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const expectedRule = [];
 if (kind === 'host-loops') { let from = 0; for (let i = 0; i < count; i++) { const start = source.indexOf('Worksheets', from); from = start + 10; expectedRule.push(['assignmentObjectTypeMismatch', "For Each Sets the items of 'Worksheets', each a Worksheet, into 'actor', a Class1. This will raise Run-time error '13': Type mismatch.", { start, end: from }]); } }
 const runRule = () => { const out = []; api.checkRuntimeMemberNotFound(source, mod, symbols, { projectClassMembers: classes }, undefined, (...args) => out.push(args)); return out; };
 names = 0; interfaces = 0; implementedNames = 0;
 assert.deepEqual(runRule(), expectedRule);
 workCounts.push({ kind, count, names, interfaces, implementedNames });
 if (!baseline) { assert.ok(names <= count * 4 + 16); assert.ok(interfaces <= count * 4 + 16); assert.ok(implementedNames <= count * 4 + 16); }
 // Instrumentation is exclusively untimed; actual clocks use plain metadata properties.
 for (let i = 0; i < classes.length; i++) {
  Object.defineProperty(classes[i], 'name', { value: 'Class' + (i + 1), writable: true, configurable: true, enumerable: true });
  Object.defineProperty(classes[i], 'implements', { value: implementsLists[i], writable: true, configurable: true, enumerable: true });
 }
 const runModule = () => {
  const errors = [], diagnostics = api.analyzeModule(source, { projectClassMembers: classes, onInternalError: e => errors.push(String(e)) });
  assert.deepEqual(errors, []); return diagnostics;
 };
 const expectedModule = runModule();
 if (kind === 'host-loops') { assert.equal(expectedModule.length, count); assert.ok(expectedModule.every(d => d.code === 'assignment-object-type-mismatch')); } else assert.deepEqual(expectedModule, []);
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
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, workCounts, rows, scope: 'Parsed AST/symbols outside direct-rule timing; complete module includes analyzer preparation and error checks. Independently exact host diagnostics/messages/spans or empty held/shared results, also checked in complete calls; equality outside clocks. Plain timed metadata, untimed work counters. No editor/cold/heap claim.' }, null, 2));
