import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { performance } from 'node:perf_hooks';
import { writeFileSync, unlinkSync, mkdtempSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const baseline = process.argv.find(s => s.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-project-query-')), file = join(scratch, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) { b.onLoad({ filter: /[\\/]completion[\\/]memberAccess\.ts$/ }, p => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/completion/memberAccess.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(p.path) })); } }] : [];
 const b = await build({ plugins, stdin: { contents: "export {projectClassMemberAt} from './src/analyzer/completion/memberAccess';export {parseModule} from './src/analyzer/parser/parseModule';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, b.outputFiles[0].contents); api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(scratch); } }
const rows = [], workCounts = [];
for (const n of [1, 100, 1000]) {
 let reads = 0;
 const members = Array.from({ length: n }, (_, i) => ({ get name() { reads++; return 'M' + i; }, moduleName: 'Class1', kind: 'property', returns: 'Object', writable: true }));
 const classes = [{ name: 'Class1', moduleName: 'Class1', kind: 'class', exhaustive: true, members }];
 const source = 'Sub Go()\nDim actor As Class1\nactor.\nEnd Sub\n', offset = source.indexOf('actor.') + 6, mod = api.parseModule(source);
 const raw = position => { const ctx = { projectClassMembers: classes, parsedModule: mod, memberSurfaceCache: new Map() }; return Array.from({ length: n }, (_, i) => api.projectClassMemberAt(source, offset, 'M' + (position === 'first' ? 0 : position === 'last' ? n - 1 : i), ctx)); };
 const fullSource = ['Option Explicit', 'Sub Go()', 'Dim actor As New Class1', ...Array.from({ length: n }, (_, i) => 'actor.M' + i + ' = 1'), 'End Sub', ''].join('\n');
 let from = 0;
 const diagnostics = Array.from({ length: n }, (_, i) => { const name = 'actor.M' + i, start = fullSource.indexOf(name, from) + 6; from = start + name.length - 6; return ['set-required', "Object assignment to '" + name + "' requires Set because it expects Object.", { start, end: start + name.length - 6 }]; });
 const full = () => { const errors = [], output = api.analyzeModule(fullSource, { projectClassMembers: classes, onInternalError: e => errors.push(String(e)) }); return { diagnostics: output.map(d => [d.code, d.message, d.span]), errors }; };
 const cases = [['raw-first-pass', () => raw('first'), Array(n).fill(members[0])], ['raw-last-pass', () => raw('last'), Array(n).fill(members[n - 1])], ['raw-distinct-pass', () => raw('distinct'), members], ['complete-module-diagnostics', full, { diagnostics, errors: [] }]];
 for (const [scope, run, want] of cases) { reads = 0; assert.deepEqual(run(), want); workCounts.push({ n, scope, reads }); }
 for (let i = 0; i < n; i++) Object.defineProperty(members[i], 'name', { value: 'M' + i, writable: true, enumerable: true, configurable: true });
 for (const [scope, run, want] of cases) { const samples = []; for (let round = -3; round < 9; round++) { const start = performance.now(), actual = run(), elapsed = performance.now() - start; assert.deepEqual(actual, want); if (round >= 0) samples.push(elapsed); } samples.sort((a, b) => a - b); rows.push({ n, scope, medianMs: +samples[4].toFixed(5) }); }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, workCounts, rows, scope: 'Class1 has N signatureless Object properties. Raw queries have a fresh pass surface cache and compare selected metadata object identity. Parsing excluded from raw-query timings, included in full diagnostics. Getters count only untimed work; timing uses plain properties. Complete expected rows/diagnostics and empty errors asserted outside the clock. No editor/heap/cold-start claim.' }, null, 2));
