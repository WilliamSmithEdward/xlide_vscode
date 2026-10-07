import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { performance } from 'node:perf_hooks';
import { writeFileSync, unlinkSync, mkdtempSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const baseline = process.argv.find(s => s.startsWith('--baseline='))?.slice(11), scratch = mkdtempSync(join(tmpdir(), 'xlide-runtime-metadata-')), file = join(scratch, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) { b.onLoad({ filter: /[\\/]rules[\\/]lateBoundMembers\.ts$/ }, p => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/lateBoundMembers.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(p.path) })); } }] : [];
 const b = await build({ plugins, stdin: { contents: "export {checkRuntimeMemberNotFound} from './src/analyzer/diagnostics/rules/lateBoundMembers';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, b.outputFiles[0].contents); api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(scratch); } }
const rows = [], workCounts = [], messages = {
 worksheetFunction: "WorksheetFunction has no function 'Missing'. The VBE compiles the name; this will raise Run-time error '438': Object doesn't support this property or method.",
 activeSheet: "ActiveSheet has no member 'Missing': neither a Worksheet nor a Chart has one, and no document module of the project declares it. This will raise Run-time error '438': Object doesn't support this property or method.",
 formControls: 'The form Form1 has no control named "Missing". This will raise Run-time error \'-2147024809\': Could not find the specified object.',
};
for (const scope of Object.keys(messages)) for (const n of [1, 100, 1000]) {
 let names = 0, returns = 0;
 const items = Array.from({ length: n }, (_, i) => ({ get name() { names++; return 'M' + i; }, kind: 'method' })), controls = Array.from({ length: n }, (_, i) => ({ name: 'Control' + i, moduleName: 'Form1', kind: 'property', get returns() { returns++; return 'MSForms.TextBox'; } }));
 const model = { hostName: 'Excel', source: 'benchmark model', types: {
  'Excel.Application': { displayName: 'Application', exhaustive: true, members: [{ name: 'WorksheetFunction', kind: 'property', returns: 'Excel.WorksheetFunction' }] },
  'Excel.WorksheetFunction': { displayName: 'WorksheetFunction', members: scope === 'worksheetFunction' ? items : [] },
  'Excel.Worksheet': { displayName: 'Worksheet', members: scope === 'activeSheet' ? items : [] }, 'Excel.Chart': { displayName: 'Chart', members: [] },
 }, aliases: { worksheet: 'Excel.Worksheet', application: 'Excel.Application' }, globals: { WorksheetFunction: 'Excel.WorksheetFunction', ActiveSheet: 'union:Excel.Worksheet|Excel.Chart' } };
 const classes = scope === 'formControls' ? [{ name: 'Form1', moduleName: 'Form1', kind: 'userform', exhaustive: true, members: controls }] : [];
 const line = scope === 'worksheetFunction' ? 'Debug.Print WorksheetFunction.Missing' : scope === 'activeSheet' ? 'Debug.Print ActiveSheet.Missing' : 'Debug.Print actor.Controls("Missing")';
 const source = ['Option Explicit', 'Sub Go()', ...(scope === 'formControls' ? ['Dim actor As New Form1'] : []), ...Array(n).fill(line), 'End Sub', ''].join('\n'), mod = api.parseModule(source), symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 let from = 0;
 const expected = Array.from({ length: n }, () => { const marker = scope === 'formControls' ? '"Missing"' : 'Missing', start = source.indexOf(marker, from); from = start + marker.length; return ['runtimeMemberNotFound', messages[scope], { start, end: from }]; });
 const rule = () => { const out = []; api.checkRuntimeMemberNotFound(source, mod, symbols, { model, projectClassMembers: classes, parsedModule: mod, memberSurfaceCache: new Map() }, undefined, (...v) => out.push(v)); return out; };
 const full = () => { const errors = [], diagnostics = api.analyzeModule(source, { hostModel: model, projectClassMembers: classes, onInternalError: e => errors.push(String(e)) }); return { diagnostics: diagnostics.map(d => [d.code, d.message, d.span]), errors }; };
 const cases = [['public-rule', rule, expected], ['complete-module-diagnostics', full, { diagnostics: expected.map(r => ['runtime-member-not-found', r[1], r[2]]), errors: [] }]];
 for (const [apiScope, run, want] of cases) { names = 0; returns = 0; assert.deepEqual(run(), want); workCounts.push({ scope, n, apiScope, names, returns }); }
 for (let i = 0; i < n; i++) { Object.defineProperty(items[i], 'name', { value: 'M' + i, writable: true, enumerable: true, configurable: true }); Object.defineProperty(controls[i], 'returns', { value: 'MSForms.TextBox', writable: true, enumerable: true, configurable: true }); }
 for (const [apiScope, run, want] of cases) { const samples = []; for (let round = -3; round < 9; round++) { const start = performance.now(), actual = run(), elapsed = performance.now() - start; assert.deepEqual(actual, want); if (round >= 0) samples.push(elapsed); } samples.sort((a, b) => a - b); rows.push({ scope, n, apiScope, medianMs: +samples[4].toFixed(5) }); }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, workCounts, rows, scope: 'N metadata entries and N missing-name references. Host index warm by timing; public parse/symbol setup excluded, full setup included. Fresh query construction timed. Getter work only untimed, plain properties timed. Independent exact N diagnostics and empty internal errors asserted outside the clock. No editor/heap/cold-start claim.' }, null, 2));
