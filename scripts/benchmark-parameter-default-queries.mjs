import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
const baseline = process.argv.find(a => a.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-parameter-default-')), file = join(scratch, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) {
  b.onLoad({ filter: /[\\/]rules[\\/]declarations\.ts$/ }, a => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/declarations.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(a.path) }));
 } }] : [];
 const bundle = await build({ plugins, stdin: { contents: "export {checkParameterDefaultValues,checkNonConstantParameterDefaults} from './src/analyzer/diagnostics/rules/declarations';export {parseModule} from './src/analyzer/parser/parseModule';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, bundle.outputFiles[0].contents);
 api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(scratch); } }
const rows = [], workCounts = [];
for (const kind of ['repeated', 'distinct']) for (const count of [1, 100, 1000]) {
 let names = 0;
 const classes = Array.from({ length: count + 1 }, (_, i) => ({ get name() { names++; return 'Class' + (i + 1); }, moduleName: 'Class' + (i + 1), kind: 'class', members: [] }));
 const types = Array.from({ length: count }, (_, i) => 'Class' + (kind === 'distinct' ? i + 1 : 1));
 const source = 'Option Explicit\n' + types.map((type, i) => `Sub P${i}(Optional actor As ${type} = 0, Optional n As Long = Factory())\nEnd Sub\n`).join('');
 const mod = api.parseModule(source);
 const expected = {};
 for (const rule of ['checkParameterDefaultValues', 'checkNonConstantParameterDefaults']) {
  let from = 0;
  expected[rule] = types.map(type => {
   const needle = rule === 'checkParameterDefaultValues' ? '= 0' : 'Factory()', index = source.indexOf(needle, from);
   const start = index + (needle === '= 0' ? 2 : 0), end = index + needle.length;
   from = end;
   return rule === 'checkParameterDefaultValues'
    ? ['parameterDefaultTypeMismatch', `Optional parameter 'actor' expects ${type}, but its default value is numeric literal 0. Optional object parameter defaults must be Nothing.`, { start, end }]
    : ['parameterDefaultNotConstant', "Optional parameter 'n' default must be a constant expression; the call 'Factory(...)' is not constant.", { start, end }];
  });
 }
 const runRule = rule => { const out = []; api[rule](source, mod, undefined, { projectClassMembers: classes }, (...args) => out.push(args)); return out; };
 for (const rule of Object.keys(expected)) {
  names = 0;
  assert.deepEqual(runRule(rule), expected[rule]);
  workCounts.push({ kind, count, rule, names });
  if (!baseline) assert.ok(names <= count * 3 + 5);
 }
 // Counters are used only in the untimed checks. Timed metadata matches plain production properties.
 for (let i = 0; i < classes.length; i++) Object.defineProperty(classes[i], 'name', { value: 'Class' + (i + 1), writable: true, configurable: true, enumerable: true });
 const runModule = () => {
  const errors = [], diagnostics = api.analyzeModule(source, { projectClassMembers: classes, onInternalError: e => errors.push(String(e)) });
  assert.deepEqual(errors, []);
  return diagnostics;
 };
 const expectedModule = runModule();
 assert.equal(expectedModule.length, count * 2);
 for (const code of ['parameter-default-type-mismatch', 'parameter-default-not-constant']) assert.equal(expectedModule.filter(d => d.code === code).length, count);
 for (const scope of [...Object.keys(expected), 'complete-module-diagnostics']) {
  const samples = [];
  for (let round = -3; round < 9; round++) {
   const start = performance.now(), actual = scope === 'complete-module-diagnostics' ? runModule() : runRule(scope), elapsed = performance.now() - start;
   assert.deepEqual(actual, scope === 'complete-module-diagnostics' ? expectedModule : expected[scope]);
   if (round >= 0) samples.push(elapsed);
  }
  samples.sort((a, b) => a - b);
  rows.push({ kind, count, scope, medianMs: +samples[4].toFixed(5) });
 }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, workCounts, rows, scope: 'Parsed AST outside direct-rule timing; complete module includes analyzer preparation and error checks. Independently exact rule messages/spans and N of each expected diagnostic code; equality outside clocks. Plain timed metadata, untimed work counters. No editor/cold/heap claim.' }, null, 2));
