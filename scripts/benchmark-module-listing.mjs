// Reproducible timings for listing a workbook's modules; no VS Code or Office application required.
// Run: node scripts/benchmark-module-listing.mjs [--rounds=15] [--rows=60000] [--baseline=<commit>]
//
// Two workbooks, each with its first sheet grown to --rows rows of data: one
// whose sheet has no shapes (SheetsFixture) and one whose sheet has shapes and
// form controls (ShapesFixture). Each is listed warm - the project already
// parsed and cached - through the desktop codec (zlib) and through the web
// build's (the pure-TypeScript inflate). The digest is a hash of everything
// the calls returned, so two runs can be checked to have given the same answers.
// --baseline builds the two files the change touched as that commit had them,
// and everything else from the working tree.
import { build } from 'esbuild';
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { performance } from 'node:perf_hooks';
import { execFileSync } from 'node:child_process';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const require = createRequire(import.meta.url);
const option = (name, fallback) => Number(process.argv.find(arg => arg.startsWith(`--${name}=`))?.split('=')[1] ?? fallback);
const rounds = option('rounds', 15);
const rowCount = option('rows', 60000);
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice('--baseline='.length);
const CHANGED = ['xlsx', 'xlsxShapes'];
const baselineSources = {
    name: 'baseline-sources',
    setup(b) {
        b.onLoad({ filter: new RegExp(`[\\\\/]src[\\\\/]vba[\\\\/](${CHANGED.join('|')})\\.ts$`) }, args => ({
            contents: execFileSync('git', ['show', `${baseline}:src/vba/${args.path.split(/[\\/]/).pop()}`], { cwd: root, encoding: 'utf8', maxBuffer: 64 << 20 }),
            loader: 'ts',
            resolveDir: dirname(args.path),
        }));
    },
};
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
if (!Number.isInteger(rowCount) || rowCount < 0 || rowCount > 1000000) throw new Error('rows must be an integer from 0 to 1000000');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-listing-benchmark-'));
const entry = `
    export { listModules, readModules, listShapes } from './src/vba/projectService';
    export { ZipArchive } from './src/vba/zip';
    export { primeHostFile } from './src/vba/hostPlatformWeb';
    `;
async function load(platform) {
    const { webLeafSwap, webInject } = require(join(root, 'webBuild.js'));
    const result = await build({ stdin: { contents: entry, resolveDir: root, loader: 'ts' }, bundle: true, format: 'cjs', write: false,
        ...(platform === 'web'
            ? { platform: 'browser', inject: webInject(), plugins: [...(baseline ? [baselineSources] : []), webLeafSwap()] }
            : { platform: 'node', plugins: baseline ? [baselineSources] : [] }) });
    const bundle = join(scratch, `${platform}.cjs`);
    writeFileSync(bundle, result.outputFiles[0].contents);
    return require(bundle);
}
function measure(run) {
    for (let i = 0; i < 3; i++) run();
    const samples = [];
    for (let i = 0; i < rounds; i++) {
        const start = performance.now();
        run();
        samples.push(performance.now() - start);
    }
    samples.sort((a, b) => a - b);
    return { medianMs: +samples[Math.floor(samples.length / 2)].toFixed(3), p95Ms: +samples[Math.ceil(samples.length * 0.95) - 1].toFixed(3) };
}
const rows = [];
let sheetXmlBytes = 0;
try {
    const node = await load('node');
    const web = await load('web');
    for (const fixture of ['SheetsFixture.xlsm', 'ShapesFixture.xlsm']) {
        // Rows that do not repeat, so the part does not compress to a few bytes.
        const zip = node.ZipArchive.read(readFileSync(join(root, 'tests', 'fixtures', 'binaries', fixture)));
        const part = 'xl/worksheets/sheet1.xml';
        const data = Array.from({ length: rowCount }, (_, i) =>
            `<row r="${i + 2}"><c r="A${i + 2}"><v>${(i * 7919) % 100003}</v></c><c r="B${i + 2}" t="str"><v>item ${i.toString(36)}</v></c></row>`).join('');
        const xml = zip.read(part).toString('utf8').replace(/<sheetData\b[^>]*\/>|<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${data}</sheetData>`);
        sheetXmlBytes = Buffer.byteLength(xml);
        zip.write(part, Buffer.from(xml, 'utf8'));
        const file = join(scratch, fixture);
        const bytes = zip.toBytes();
        writeFileSync(file, bytes);
        web.primeHostFile(file, bytes, 1);
        for (const [platform, api] of [['node', node], ['web', web]]) {
            const digest = createHash('sha256')
                .update(JSON.stringify([api.listModules(file), api.readModules(file), api.listShapes(file)]))
                .digest('hex').slice(0, 16);
            for (const call of ['listModules', 'readModules', 'listShapes']) {
                rows.push({ name: `${fixture}/${platform}/${call}`, ...measure(() => api[call](file)), digest });
            }
        }
    }
} finally {
    rmSync(scratch, { recursive: true, force: true });
}
console.log(JSON.stringify({ node: process.version, cpu: cpus()[0]?.model, baseline: baseline ?? null, rounds, rows: rowCount, sheetXmlBytes, results: rows }, null, 2));
