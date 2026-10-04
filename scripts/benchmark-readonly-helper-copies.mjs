import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { dirname, join, relative } from 'node:path';
import { performance } from 'node:perf_hooks';
import { writeFileSync, unlinkSync, mkdtempSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const files = ['diagnostics/typeInference.ts', 'diagnostics/straightLineValues.ts', 'diagnostics/rules/typeMembers.ts', 'diagnostics/rules/overflow.ts', 'diagnostics/rules/lateBoundMembers.ts', 'diagnostics/rules/handlerFlow.ts', 'diagnostics/rules/formContents.ts', 'diagnostics/rules/errorValues.ts', 'diagnostics/rules/collectionState.ts', 'diagnostics/rules/accessData.ts'].map(p => 'src/analyzer/' + p);
const baseline = process.argv.find(s => s.startsWith('--baseline='))?.slice(11), scratch = mkdtempSync(join(tmpdir(), 'xlide-helper-copies-')), file = join(scratch, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) { b.onLoad({ filter: /\.ts$/ }, p => { const name = relative(process.cwd(), p.path).replace(/\\/g, '/'); if (files.includes(name)) return { contents: execFileSync('git', ['show', baseline + ':' + name], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(p.path) }; }); } }] : [];
 const b = await build({ plugins, stdin: { contents: "export {checkRuntimeMemberNotFound} from './src/analyzer/diagnostics/rules/lateBoundMembers';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';export {statementTokensCached as statementTokens} from './src/analyzer/lexer/tokenHelpers';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, b.outputFiles[0].contents); api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(scratch); } }
const rows = [], workCounts = [], message = 'CreateObject: "Dictionary" lacks its library: the ProgID is "scripting.dictionary". This will raise Run-time error \'429\': ActiveX component can\'t create object.';
for (const fixture of ['wide-call', 'many-calls']) for (const n of [1, 100, 1000]) {
 const lines = Array(fixture === 'wide-call' ? 1 : n).fill('Set obj = CreateObject("Dictionary"' + (fixture === 'wide-call' ? ', ' + Array(n).fill('0').join(', ') : '') + ')');
 const source = ['Option Explicit', 'Sub Go()', 'Dim obj As Object', ...lines, 'End Sub', ''].join('\n'), mod = api.parseModule(source), symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 let from = 0;
 const expected = lines.map(() => { const start = source.indexOf('"Dictionary"', from); from = start + 12; return ['runtimeArgumentValue', message, { start, end: from }]; });
 const rule = () => { const out = []; api.checkRuntimeMemberNotFound(source, mod, symbols, { parsedModule: mod, memberSurfaceCache: new Map() }, undefined, (...v) => out.push(v)); return out; };
 const full = () => { const errors = [], diagnostics = api.analyzeModule(source, { onInternalError: e => errors.push(String(e)) }); return { diagnostics, errors }; };
 const procedure = mod.members.find(m => m.kind === 'Procedure'), arrays = procedure.body.filter(node => source.slice(node.span.start, node.span.end).includes('CreateObject')).map(node => api.statementTokens(source, node.span));
 let yields = 0, iterators = 0;
 const originals = arrays.map(tokens => Object.getOwnPropertyDescriptor(tokens, Symbol.iterator));
 try { for (const tokens of arrays) Object.defineProperty(tokens, Symbol.iterator, { configurable: true, value: function* () { iterators++; for (let i = 0; i < this.length; i++) { yields++; yield this[i]; } } }); assert.deepEqual(rule(), expected); workCounts.push({ fixture, n, tokenElements: arrays.reduce((sum, tokens) => sum + tokens.length, 0), iterators, yields }); }
 finally { for (let i = 0; i < arrays.length; i++) { if (originals[i]) Object.defineProperty(arrays[i], Symbol.iterator, originals[i]); else delete arrays[i][Symbol.iterator]; } }
 const fullExpected = full(); assert.deepEqual(fullExpected.errors, []); assert.deepEqual(fullExpected.diagnostics.filter(d => d.code === 'runtime-argument-value').map(d => [d.code, d.message, d.span]), expected.map(r => ['runtime-argument-value', r[1], r[2]]));
 const digest = createHash('sha256').update(JSON.stringify(fullExpected)).digest('hex');
 for (const [scope, run, want] of [['public-rule', rule, expected], ['complete-module-diagnostics', full, fullExpected]]) { const samples = []; for (let round = -3; round < 9; round++) { const start = performance.now(), actual = run(), elapsed = performance.now() - start; assert.deepEqual(actual, want); if (round >= 0) samples.push(elapsed); } samples.sort((a, b) => a - b); rows.push({ fixture, n, scope, medianMs: +samples[4].toFixed(5), fullOutputDigest: digest }); }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, workCounts, rows, scope: 'Wide-call: one CreateObject with N extra arguments; many-calls: N ordinary one-argument calls. Iterator work is untimed/restored; timed arrays use normal iteration. Exact independent public findings and primary full findings asserted. Complete full outputs/digests and empty errors checked outside clock. Public parse/symbol setup excluded, full setup included. No heap/editor/cold-start or universal speedup claim.' }, null, 2));
