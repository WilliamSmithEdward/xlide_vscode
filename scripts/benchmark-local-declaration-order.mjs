// Run in the target checkout: node scripts/benchmark-local-declaration-order.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-local-declaration-order-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { checkLocalDeclarationOrder } from './src/analyzer/diagnostics/rules/localDeclarationOrder';
        export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';
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
for(const [count,explicit,constants] of [[100,true,false],[1000,true,false],[5000,true,false],[1000,false,true]]) {
    const body=constants?Array.from({length:count},(_,i)=>'Const A'+i+' = B'+i+'\n').join(''):Array.from({length:count},(_,i)=>'Debug.Print B'+i+'\n').join('');
    const source=(explicit?'Option Explicit\n':'')+'Sub Main()\n'+body+Array.from({length:count},(_,i)=>'Dim B'+i+' As Long\n').join('')+'End Sub';
    const module=api.parseModule(source), symbols=api.buildModuleSymbols('Module','standard',source,{parsedModule:module});
    measure(count+'-forward-uses/'+(explicit?'explicit':'implicit')+'/'+(constants?'const':'statement'),()=>{let found=0;api.checkLocalDeclarationOrder(source,module,symbols,undefined,undefined,undefined,()=>{found++;});if(found!==count)throw new Error('Wrong diagnostics: '+found);});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
