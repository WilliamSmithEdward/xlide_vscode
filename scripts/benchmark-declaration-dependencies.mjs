// Run in the target checkout: node scripts/benchmark-declaration-dependencies.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-declaration-dependencies-'));
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
for(const [roots,length] of [[1,100],[100,100],[200,500]]) {
    const source=Array.from({length:roots},(_,i)=>'Public Const A'+i+' = Other.B0\n').join('');
    const module=api.parseModule(source), external=new Map();
    for(let i=0;i<length;i++)external.set('other.b'+i,i===length-1?'1':'B'+(i+1));
    measure(roots+'-roots/'+length+'-external-chain',()=>{let found=0;api.checkDeclarationOrder(source,module,'Module',external,undefined,undefined,()=>{found++;});if(found!==0)throw new Error('Unexpected diagnostic');});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
