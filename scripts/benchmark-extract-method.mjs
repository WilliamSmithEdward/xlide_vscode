// Run in the target checkout: node scripts/benchmark-extract-method.mjs [--rounds=15]
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-extract-method-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const plugins = baseline ? [{ name: 'baseline', setup(builder) {
        builder.onLoad({ filter: /extractMethod\.ts$/ }, args => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/refactor/extractMethod.ts'], { cwd: root, encoding: 'utf8' }), loader: 'ts', resolveDir: dirname(args.path) }));
    } }] : [];
    const result = await build({ plugins, stdin: { contents: `
        export { extractMethod } from './src/analyzer/refactor/extractMethod';
        export { findIdentifierOccurrences } from './src/vbaSourceScan';
        `, resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
    writeFileSync(bundle, result.outputFiles[0].contents);
    api = createRequire(import.meta.url)(bundle);
} finally {
    try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    rmdirSync(scratch);
}
const rows = [];
function measure(name, run) {
    for (let i = 0; i < 3; i++) run();
    const samples = [];
    for (let i = 0; i < rounds; i++) {
        const start = performance.now(); run(); samples.push(performance.now() - start);
    }
    samples.sort((a,b)=>a-b);
    rows.push({ name, medianMs: +samples[Math.floor(samples.length/2)].toFixed(3),
        p95Ms: +samples[Math.ceil(samples.length*.95)-1].toFixed(3) });
}
if (process.argv.includes('--local-order')) {
    const signature = result => createHash('sha256').update(JSON.stringify(result)).digest('hex');
    let salt = 0;
    for (const [count, padding, orderMode, form] of [[1,0,'ordered','read'],[5,0,'shuffled','read'],[100,0,'shuffled','read'],[1000,0,'ordered','read'],[1000,8000,'shuffled','read'],[1000,100000,'shuffled','read'],[3000,100000,'shuffled','write'],[1000,100000,'shuffled','few']]) {
        const names = Array.from({ length: count }, (_, i) => 'local' + i.toString().padStart(4, '0'));
        const order = orderMode === 'ordered' ? names : names.map((_, i) => names[(i * 37) % count]);
        const comments = Array.from({ length: Math.ceil(padding / 200) }, (_, i) => "' " + 'x'.repeat(Math.min(200, padding - i * 200)) + '\n').join('');
        const orderComments = Array.from({ length: Math.ceil(order.length / 12) }, (_, i) => "' " + order.slice(i * 12, (i + 1) * 12).join(' ') + '\n').join('');
        const prefix = comments + orderComments + 'Option Explicit\nSub Main()\n' + names.map(name => 'Dim ' + name + ' As Long\n').join('');
        const touched = form === 'few' ? names.slice(0, 5) : names;
        const selected = touched.map(name => form === 'write' ? name + '=1' : 'Debug.Print ' + name).join('\n');
        const base = prefix + selected + '\nDebug.Print "after"\nEnd Sub\n';
        for (const mode of ['warm', 'fresh']) {
            const samples = []; let expected;
            for (let round = -3; round < rounds; round++) {
                const source = mode === 'fresh' ? base + "' sample " + (++salt) + '\n' : base;
                const input = { source, span: { start: prefix.length, end: prefix.length + selected.length }, name: 'Work' };
                const start = performance.now(), result = api.extractMethod(input), elapsed = performance.now() - start;
                if (!result.ok) throw new Error(result.reason);
                const current = signature(result);
                if (expected && current !== expected) throw new Error('Changed complete refactor result');
                expected = current;
                if (round >= 0) samples.push(elapsed);
            }
            samples.sort((a,b)=>a-b);
            rows.push({ name: count + '/' + padding + '/' + orderMode + '/' + form + '/' + mode, locals: count, touched: touched.length, medianMs: +samples[Math.floor(rounds/2)].toFixed(3), p95Ms: +samples[Math.ceil(rounds*.95)-1].toFixed(3), resultSignature: expected });
        }
    }
} else {
for(const [locals, statements] of [[100,1000],[300,3000],[1000,2000]]) {
    const prefix='Option Explicit\nPublic Sub Main()\n'+Array.from({length:locals},(_,i)=>'Dim v'+i+' As Long\n').join('');
    const selected=Array.from({length:5},(_,i)=>'v'+i+' = v'+i+' + 1\n').join('');
    const source=prefix+selected+Array.from({length:statements},(_,i)=>'Debug.Print v'+(i%locals)+'\n').join('')+'End Sub\n';
    const input={source,span:{start:prefix.length,end:prefix.length+selected.length-1},name:'Extracted'};
    measure(locals+'-locals/'+statements+'-statements',()=>{const result=api.extractMethod(input);if(!result.ok)throw new Error(result.reason);});
}
const referenceSource=Array.from({length:10000},(_,i)=>'v'+(i%100)+' = v'+((i+1)%100)+' + 1\n').join('');
measure('single-name-scan/10000-statements',()=>{if(api.findIdentifierOccurrences(referenceSource,'v0').length!==200)throw new Error('Missing occurrences');});
}
console.log(JSON.stringify({ baseline: baseline ?? null, node: process.version, cpu: cpus()[0]?.model, rounds, rows }, null, 2));
