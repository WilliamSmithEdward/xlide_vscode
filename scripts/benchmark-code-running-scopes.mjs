// Run: node scripts/benchmark-code-running-scopes.mjs [--rounds=15] [--baseline=<git-ref>]
import { build } from 'esbuild';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const scratch = mkdtempSync(join(tmpdir(), 'xlide-code-running-scopes-'));
const bundle = join(scratch, 'api.cjs');
let api;
try {
    const plugins = baseline ? [{ name: 'baseline', setup(builder) {
        builder.onLoad({ filter: /diagnostics[\\/]typeInference\.ts$/ }, () => ({
            contents: execFileSync('git', ['show', `${baseline}:src/analyzer/diagnostics/typeInference.ts`], { cwd: root, encoding: 'utf8' }),
            loader: 'ts', resolveDir: join(root, 'src/analyzer/diagnostics'),
        }));
    } }] : [];
    const result = await build({ plugins, stdin: { contents: `
        export { statementMayChangeModuleVariable } from './src/analyzer/diagnostics/typeInference';
        export { parseModule } from './src/analyzer/parser/parseModule';
        export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';
    `, resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(bundle, result.outputFiles[0].contents);
    api = createRequire(import.meta.url)(bundle);
} finally {
    try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    rmdirSync(scratch);
}
const results = [];
for (const count of [100, 500, 2000]) {
    const source = 'Option Explicit\nPrivate tracked As Long\n'
        + Array.from({ length: count }, (_, i) => `Private Const K${i} As Long = ${i}\n`).join('')
        + 'Sub P()\n' + Array.from({ length: 80 }, (_, i) => `Dim x${i} As Long\n`).join('')
        + Array.from({ length: count }, () => `x0 = K${count - 1} + 1\n`).join('') + 'End Sub\n';
    const parsedModule = api.parseModule(source);
    const proc = parsedModule.members.find(member => member.kind === 'Procedure');
    const statements = proc.body.filter(node => node.kind === 'Assignment');
    const samples = [];
    for (let round = 0; round < rounds + 3; round++) {
        const symbols = api.buildModuleSymbols('M', 'standard', source, { parsedModule });
        const start = performance.now();
        for (const node of statements) {
            if (api.statementMayChangeModuleVariable(source, symbols, proc, node.span, 'tracked')) {
                throw new Error('A local/constant-only statement unexpectedly runs external code');
            }
        }
        if (round >= 3) samples.push(performance.now() - start);
    }
    samples.sort((a, b) => a - b);
    results.push({ constants: count, statements: count, locals: 80,
        medianMs: +samples[Math.floor(samples.length / 2)].toFixed(3),
        p95Ms: +samples[Math.ceil(samples.length * 0.95) - 1].toFixed(3) });
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, baseline: baseline ?? null, rounds, results }, null, 2));
