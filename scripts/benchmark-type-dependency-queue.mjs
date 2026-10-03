// Run in the target checkout: node scripts/benchmark-type-dependency-queue.mjs [--rounds=15]
import { buildSync } from 'esbuild';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.split('=')[1] ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be an integer from 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-type-dependency-queue-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { checkDeclarationOrder } from './src/analyzer/diagnostics/rules/declarationOrder';
        export { parseModule } from './src/analyzer/parser/parseModule';
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
const source = 'Public Type Root\n    a As Entry\nEnd Type';
const module = api.parseModule(source);
function surface(name, references) {
    return { name, moduleName: 'Other', kind: 'userType', members: references.map((returns,i)=>({ name:'field'+i, kind:'property', moduleName:'Other', returns })) };
}
for (const [width,cycle] of [[25,false],[100,false],[300,false],[300,true]]) {
    const leaves=Array.from({length:width},(_,i)=>'Leaf'+i);
    const branches=Array.from({length:width},(_,i)=>'Branch'+i);
    const project=[surface('Entry',branches), ...branches.map(name=>surface(name,leaves)), ...leaves.map((name,i)=>surface(name,cycle&&i===width-1?['Root']:['Long']))];
    measure(width+'-wide-shared-graph/'+(cycle?'cyclic':'acyclic'),()=>{
        let found=0;
        api.checkDeclarationOrder(source,module,'Module',undefined,project,undefined,()=>{found++;});
        if(found!==Number(cycle))throw new Error('Wrong diagnostics: '+found);
    });
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
