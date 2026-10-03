// Run: node scripts/benchmark-object-read-token-scans.mjs [--baseline=<commit>] [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-object-read-token-scans-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(builder) {
  builder.onLoad({ filter: /rules[\\/]objectValues\.ts$/ }, () => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/objectValues.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: join(root, 'src/analyzer/diagnostics/rules') }));
 }}] : [];
 const result = await build({ plugins, stdin: { contents: "export { hostObjectModelForToken } from './src/analyzer/host/hostRegistry'; export { checkObjectDefaultValues } from './src/analyzer/diagnostics/rules/objectValues'; export { statementTokens, forEachStatementWithHeaders } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(bundle, result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; } rmdirSync(scratch); }
const rows = [];
for (const count of [1, 10, 100, 500, 1500]) {
 for (const mode of ['operators', 'print-items', 'value-arguments', 'scalar-control']) {
  const item = mode === 'value-arguments' ? 'CStr(x)' : 'n';
  const separator = mode === 'print-items' ? '; ' : ' + ';
  const chunks = [];
  for (let i = 0; i < count; i += 75) chunks.push(Array.from({ length: Math.min(75, count - i) }, () => item).join(separator));
  const statement = mode === 'scalar-control' ? 'n = 1' : 'Debug.Print ' + chunks.join(separator + '_\n ');
  const source = 'Sub P()\nDim n As Long\nDim x As Object\n' + statement + '\nEnd Sub';
  const mod = api.parseModule(source);
  if (mode !== 'scalar-control') {
   const proc = mod.members.find(member => member.kind === 'Procedure');
   const print = proc?.body.find(node => node.kind === 'Statement' && node.raw.trim().startsWith('Debug.Print'));
   if (!print || api.statementTokens(source, print.span).filter(tok => tok.rawText === (mode === 'value-arguments' ? 'x' : 'n')).length !== count) throw Error('Fixture did not retain all reads');
  }
  let previousRoot;
  function prepare() {
   const symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
   if (symbols.root === previousRoot) throw Error('Reused binding root');
   previousRoot = symbols.root;
   return () => {
    let hits = 0;
    const factory = api.checkObjectDefaultValues(source, symbols, {}, () => hits++);
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
