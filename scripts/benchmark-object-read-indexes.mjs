// Run: node scripts/benchmark-object-read-indexes.mjs [--baseline=<commit>] [--rounds=15]
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.slice(9) ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw Error('rounds must be 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-object-read-indexes-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(builder) {
  builder.onLoad({ filter: /rules[\\/]objectValues\.ts$/ }, () => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/objectValues.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: join(root, 'src/analyzer/diagnostics/rules') }));
 }}] : [];
 const result = await build({ plugins, stdin: { contents: "export { hostObjectModelForToken } from './src/analyzer/host/hostRegistry'; export { checkObjectDefaultValues } from './src/analyzer/diagnostics/rules/objectValues'; export { forEachStatementWithHeaders } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(bundle, result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; } rmdirSync(scratch); }
const rows = [];
for (const count of [1, 10, 100, 1000, 3000]) {
 for (const mode of ['host-types', 'late-bound', 'late-bound-no-host', 'scalar-control']) {
  const type = mode === 'host-types' ? 'Application' : mode.startsWith('late-bound') ? 'Object' : 'Long';
  const source = 'Sub P()\nDim n As Long\n' + Array.from({ length: count }, (_, i) => 'Dim x' + i + ' As ' + type).join('\n') + '\n' + 'n = 1\n'.repeat(1000) + 'End Sub';
  const mod = api.parseModule(source);
  let previousRoot;
  function prepare() {
   const symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
   if (symbols.root === previousRoot) throw Error('Reused binding root');
   previousRoot = symbols.root;
   return () => {
    let hits = 0;
    const factory = api.checkObjectDefaultValues(source, symbols, mode === 'late-bound-no-host' ? { model: api.hostObjectModelForToken('other') } : {}, () => hits++);
    for (const proc of mod.members) { if (proc.kind === 'Procedure') { const visit = factory(proc); if (visit) api.forEachStatementWithHeaders(source, proc.body, visit); } }
    if (hits !== 0) throw Error('Unexpected findings: ' + hits);
   };
  }
  for (let i = 0; i < 3; i++) prepare()();
  const samples = [];
  for (let i = 0; i < rounds; i++) { const run = prepare(), start = performance.now(); run(); samples.push(performance.now() - start); }
  samples.sort((a, b) => a - b);
  rows.push({ name: count + '-' + mode, medianMs: +samples[Math.floor(rounds / 2)].toFixed(3), p95Ms: +samples[Math.ceil(rounds * .95) - 1].toFixed(3) });
 }
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, rounds, rows }, null, 2));
