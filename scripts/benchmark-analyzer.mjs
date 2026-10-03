// Reproducible pure-analyzer timings; no VS Code or Office application required.
// Run: node scripts/benchmark-analyzer.mjs [--rounds=7]
import { buildSync } from 'esbuild';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 7);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-analyzer-benchmark-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { analyzeVbaModuleSource } from './src/vbaModuleAnalysis';
        export { analyzeModuleRulesIncremental } from './src/analyzer/diagnostics/incrementalRules';
        export { parseModule } from './src/analyzer/parser/parseModule';
        export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';
        export { ProjectIndex } from './src/analyzer/symbols/projectIndex';
        export { classMemberValues } from './src/analyzer/symbols/classMemberFacts';
        export { isNumericType, knownLocalLiteralValuesAt } from './src/analyzer/diagnostics/typeInference';
        export { checkMalformedLines } from './src/analyzer/diagnostics/rules/malformedLines';
        `, resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(bundle, result.outputFiles[0].contents);
    api = createRequire(import.meta.url)(bundle);
} finally {
    try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    rmdirSync(scratch);
}
function makeSource(procedures, statements) {
    return 'Option Explicit\n' + Array.from({ length: procedures }, (_, i) =>
        `Public Sub Proc${i}(ByVal a As Long)\nDim total As Long\n`
        + Array.from({ length: statements }, (_, j) => `total = a + ${j}\n`).join('')
        + 'If total > 2 Then\ntotal = total + 1\nEnd If\nEnd Sub\n').join('\n');
}
const rows = [];
function measure(name, run) {
    for (let i = 0; i < 3; i++) run();
    const samples = [];
    for (let i = 0; i < rounds; i++) {
        const start = performance.now();
        run();
        samples.push(performance.now() - start);
    }
    samples.sort((a, b) => a - b);
    rows.push({ name, medianMs: +samples[Math.floor(samples.length / 2)].toFixed(3),
        p95Ms: +samples[Math.ceil(samples.length * 0.95) - 1].toFixed(3) });
}
for (const [name, source] of [['many-100', makeSource(100, 12)], ['many-300', makeSource(300, 12)], ['body-6000', makeSource(1, 6000)]]) {
    const run = () => {
        const result = api.analyzeVbaModuleSource({ source, moduleName: 'Module', moduleKind: 'standard' });
        if (result.analysisFailures?.length) throw new Error(JSON.stringify(result.analysisFailures));
        if (result.diagnostics.length) throw new Error('Unexpected diagnostics in ' + name);
    };
    measure(name + '/full-warm', run);
    const first = api.analyzeModuleRulesIncremental(source, {}, undefined, []);
    const changed = source.replace('total = a + 0', 'total = a + 99');
    measure(name + '/body-edit', () => api.analyzeModuleRulesIncremental(changed, {}, first.state, []));
}
let numericCount = 0;
measure('numeric-type/1000000', () => {
    for (let i = 0; i < 1000000; i++) numericCount += +api.isNumericType(i % 2 ? 'long' : 'object');
});
const localSource = makeSource(1, 1000);
const localModule = api.parseModule(localSource);
measure('literal-setup/20-rules', () => {
    const localSymbols = api.buildModuleSymbols('Module', 'standard', localSource);
    for (let i = 0; i < 20; i++) api.knownLocalLiteralValuesAt(localSource, localModule.members.find(m => m.kind === 'Procedure'), localSymbols, undefined);
});
for (const count of [100, 500, 1000]) {
    const source = makeSource(count, 12);
    const mod = api.parseModule(source);
    measure('malformed-headers/' + count, () => api.checkMalformedLines(source, mod, undefined, () => {}));
    const classSource = 'Option Explicit\n' + Array.from({ length: count }, (_, i) =>
        `Public Function F${i}() As Variant\nF${i} = ${i}\nEnd Function\n`).join('');
    const symbols = api.buildModuleSymbols('Class1', 'class', classSource);
    measure('class-member-facts/' + count, () => {
        if (api.classMemberValues(classSource, symbols.root.children).size !== count) throw new Error('Missing class member facts');
    });
}
const index = new api.ProjectIndex();
for (let i = 0; i < 50; i++) index.setModule({ moduleName: 'M' + i, moduleKind: 'standard', source: makeSource(20, 12) });
measure('project-index/50-modules-warm', () => {
    for (let i = 0; i < 50; i++) {
        index.visibleProcedureNames('M' + i);
        index.visibleIdentifierSymbols('M' + i);
        index.projectMemberSurfaces('M' + i);
    }
});
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, rounds, numericCount, results: rows }, null, 2));
