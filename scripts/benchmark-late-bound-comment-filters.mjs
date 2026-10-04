import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { performance } from 'node:perf_hooks';
import { writeFileSync, unlinkSync, mkdtempSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
const baseline = process.argv.find(s => s.startsWith('--baseline='))?.slice(11), scratch = mkdtempSync(join(tmpdir(), 'xlide-comment-filters-')), file = join(scratch, 'api.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(b) { b.onLoad({ filter: /[\\/]rules[\\/]lateBoundMembers\.ts$/ }, p => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/lateBoundMembers.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(p.path) })); } }] : [];
 const b = await build({ plugins, stdin: { contents: "export {checkRuntimeMemberNotFound} from './src/analyzer/diagnostics/rules/lateBoundMembers';export {parseModule} from './src/analyzer/parser/parseModule';export {buildModuleSymbols} from './src/analyzer/symbols/buildModuleSymbols';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';export {statementTokensCached as statementTokens} from './src/analyzer/lexer/tokenHelpers';", resolveDir: process.cwd(), loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(file, b.outputFiles[0].contents); api = createRequire(import.meta.url)(file);
} finally { try { unlinkSync(file); } finally { rmdirSync(scratch); } }
const rows = [], workCounts = [];
for (const fixture of ['bare-values', 'nested-one-argument']) for (const n of [1, 100, 1000]) {
 const body = 'obj.Remove ' + (fixture === 'bare-values' ? Array(n).fill('1').join(', ') : 'Array(' + Array(n).fill('1').join(', ') + ')') + " ' trailing comment";
 const source = ['Option Explicit', 'Sub Go()', 'Dim obj As Object', 'Set obj = New Collection', body, 'End Sub', ''].join('\n'), mod = api.parseModule(source), symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const start = source.indexOf('obj.Remove') + 4, expected = fixture === 'bare-values' && n > 1 ? [['runtimeMemberNotFound', "'obj' holds a Collection here: its Remove takes at most 1 argument(s), and " + n + " are passed. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment.", { start, end: start + 6 }]] : [];
 const rule = () => { const out = []; api.checkRuntimeMemberNotFound(source, mod, symbols, { parsedModule: mod, memberSurfaceCache: new Map() }, undefined, (...v) => out.push(v)); return out; };
 const full = () => { const errors = [], diagnostics = api.analyzeModule(source, { onInternalError: e => errors.push(String(e)) }); return { diagnostics, errors }; };
 const statement = mod.members.find(m => m.kind === 'Procedure').body.find(node => source.slice(node.span.start, node.span.end).includes('obj.Remove')), tokens = api.statementTokens(source, statement.span), descriptors = tokens.map(t => Object.getOwnPropertyDescriptor(t, 'kind'));
 let filterKindReads = 0;
 try { for (let i = 0; i < tokens.length; i++) { const token = tokens[i], kind = token.kind; Object.defineProperty(token, 'kind', { configurable: true, get() { const stack = new Error().stack; if (stack.includes('at Array.filter') && stack.includes('at argumentRefusal')) filterKindReads++; return kind; } }); } assert.deepEqual(rule(), expected); }
 finally { for (let i = 0; i < tokens.length; i++) Object.defineProperty(tokens[i], 'kind', descriptors[i]); }
 workCounts.push({ fixture, n, sourceTokens: tokens.length, filterKindReads }); if (!baseline) assert.equal(filterKindReads, 0);
 const fullExpected = full(); assert.deepEqual(fullExpected.errors, []); assert.deepEqual(fullExpected.diagnostics.filter(d => d.code === 'runtime-member-not-found').map(d => [d.code, d.message, d.span]), expected.map(r => ['runtime-member-not-found', r[1], r[2]]));
 const digest = createHash('sha256').update(JSON.stringify(fullExpected)).digest('hex');
 for (const [scope, run, want] of [['public-rule', rule, expected], ['complete-module-diagnostics', full, fullExpected]]) { const samples = []; for (let round = -3; round < 9; round++) { const start = performance.now(), actual = run(), elapsed = performance.now() - start; assert.deepEqual(actual, want); if (round >= 0) samples.push(elapsed); } samples.sort((a, b) => a - b); rows.push({ fixture, n, scope, medianMs: +samples[4].toFixed(5), fullOutputDigest: digest }); }
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, warmups: 3, rounds: 9, workCounts, rows, scope: 'One late-bound Collection.Remove: N bare values or one nested Array with N values, plus trailing comment. Kind getters restored before timing. Independent public/primary full output and complete full snapshots/digests checked outside clock; errors empty. Public parse/symbol setup excluded, full setup included. No heap/editor/cold-start or universal speedup claim.' }, null, 2));
