// Run: node scripts/benchmark-semantic-global-gates.mjs [--baseline=COMMIT]
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir, cpus } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-semantic-global-gates-'));
const file = join(scratch, 'api.cjs');
let api;
try {
  const plugins = baseline ? [{ name: 'baseline', setup(builder) {
    builder.onLoad({ filter: /typeSemanticTokens\.ts$/ }, args => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/semantic/typeSemanticTokens.ts'], { cwd: root, encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) }));
  }}] : [];
  const bundled = await build({ plugins, stdin: { contents: "export { collectHostGlobalTokens } from './src/analyzer/semantic/typeSemanticTokens'; export { parseModule } from './src/analyzer/parser/parseModule'; export { tokenizeCached } from './src/analyzer/lexer/tokenize';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
  writeFileSync(file, bundled.outputFiles[0].contents);
  api = createRequire(import.meta.url)(file);
} finally {
  try { unlinkSync(file); } catch (error) { if (error.code !== 'ENOENT') throw error; }
  rmdirSync(scratch);
}
const rows = [];
for (const count of [1, 100, 10000]) for (const shape of ['member', 'type', 'bare']) {
  const samples = [];
  for (let round = -3; round < 9; round++) {
    const line = shape === 'member' ? 'receiver.UnknownMember: receiver.Application: receiver!xlUp' : shape === 'type' ? 'Set receiver = New Application' : 'Debug.Print Application';
    const source = 'Sub Run()\nDim receiver As Object\n' + (line + '\n').repeat(count) + 'End Sub\n' + "' snapshot " + round;
    api.parseModule(source); api.tokenizeCached(source);
    const expected = [];
    if (shape === 'bare') {
      let start = source.indexOf('Application');
      while (start >= 0) { expected.push({ name: 'Application', tokenType: 'variable', span: { start, end: start + 11 }, modifiers: ['defaultLibrary'] }); start = source.indexOf('Application', start + 11); }
    }
    const begin = performance.now();
    const result = api.collectHostGlobalTokens(source);
    const elapsed = performance.now() - begin;
    assert.deepEqual(result, expected);
    if (round >= 0) samples.push(elapsed);
  }
  samples.sort((a, b) => a - b);
  rows.push({ count, shape, medianMs: samples[4], p95Ms: samples[8], completeExpectedResults: true });
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds: 9, warmups: 3, scope: 'One actual global semantic collector per fresh snapshot, lexer/parser primed; declaration gathering, filtering and lookup included; source/model construction and assertions excluded; no whole-editor timing', rows }, null, 2));
