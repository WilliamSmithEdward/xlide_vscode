// Read-only worker-pipeline benchmark; uses the same sources and metadata as Run Analysis.
// node scripts/benchmark-run-analysis.mjs <workbook.xlsm> [--rounds=5]
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const input = process.argv.slice(2).find(arg => !arg.startsWith('--'));
if (!input) throw new Error('Pass a workbook path. The workbook is only read.');
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 5);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 30) throw new Error('rounds must be an integer from 3 to 30');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-run-analysis-'));
try {
    const bundle = join(scratch, 'api.cjs');
    await build({ stdin: { contents: `
        export { readModules } from './src/vba/projectService';
        export { AnalysisWorkerState } from './src/analysisWorkerLogic';
        export { moduleKindFromType } from './src/vbaProjectAnalysis';
        export { hostTokenForFileName } from './src/analyzer/host/hostRegistry';
        export { referencedHostTokens } from './src/analyzer/host/hostLibraries';
    `, resolveDir: root, loader: 'ts' }, outfile: bundle, bundle: true, platform: 'node', format: 'cjs' });
    const api = createRequire(import.meta.url)(bundle);
    const workbook = resolve(input);
    const readStart = performance.now();
    const modules = api.readModules(workbook);
    const readMs = performance.now() - readStart;
    const host = api.hostTokenForFileName(workbook);
    const references = modules[0]?.projectReferences ?? [];
    const referencedHosts = api.referencedHostTokens(host, references);
    const referencedLibraries = references.length ? references.map(reference => reference.name) : undefined;
    const samples = [];
    let expectedDigest;
    const run = () => {
        // A fresh handler measures a full pass, without saved incremental state.
        // Lexer/parser caches are warm after the first run, as in the editor.
        const worker = new api.AnalysisWorkerState();
        const start = performance.now();
        worker.handle({ kind: 'seed', projectKey: 'benchmark', generation: 1,
            modules: modules.map(mod => ({ ...mod, moduleName: mod.name })) });
        const seedMs = performance.now() - start;
        const rows = [];
        const findings = [];
        for (const mod of modules) {
            const moduleStart = performance.now();
            const result = worker.handle({ kind: 'analyze', requestId: 1,
                docKey: mod.name, projectKey: 'benchmark', generation: 1,
                source: mod.source, moduleName: mod.name, moduleType: mod.type,
                moduleKind: api.moduleKindFromType(mod.type), documentType: mod.documentType,
                host, referencedHosts, referencedLibraries,
                designerClass: mod.designerClass, workbookSheets: modules[0]?.projectSheets });
            if (result?.kind !== 'result' || result.analysisFailures?.length) throw new Error(JSON.stringify(result));
            rows.push({ module: mod.name, durationMs: performance.now() - moduleStart,
                diagnostics: result.diagnostics.length, suppressed: result.suppressedDiagnostics.length });
            findings.push([mod.name, result.diagnostics, result.suppressedDiagnostics]);
        }
        const totalMs = performance.now() - start;
        const digest = createHash('sha256').update(JSON.stringify(findings)).digest('hex');
        expectedDigest ??= digest;
        if (digest !== expectedDigest) throw new Error('Diagnostics changed between rounds');
        return { totalMs, seedMs, rows };
    };
    const first = run();
    for (let i = 0; i < rounds; i++) samples.push(run());
    const median = values => +values.sort((a, b) => a - b)[Math.floor(values.length / 2)].toFixed(2);
    console.log(JSON.stringify({ node: process.version, workbook, rounds,
        readMs: +readMs.toFixed(2), firstFullMs: +first.totalMs.toFixed(2),
        medianFullMs: median(samples.map(sample => sample.totalMs)),
        medianSeedMs: median(samples.map(sample => sample.seedMs)), diagnosticDigest: expectedDigest,
        modules: first.rows.map((row, index) => ({ module: row.module, sourceBytes: Buffer.byteLength(modules[index].source),
            diagnostics: row.diagnostics, suppressed: row.suppressed,
            medianMs: median(samples.map(sample => sample.rows[index].durationMs)) })) }, null, 2));
} finally {
    rmSync(scratch, { recursive: true, force: true });
}
