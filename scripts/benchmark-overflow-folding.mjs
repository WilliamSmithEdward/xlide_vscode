// Run in the target checkout: node scripts/benchmark-overflow-folding.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-overflow-folding-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { checkOverflow } from './src/analyzer/diagnostics/rules/overflow';
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
function procedure(name, body) { return 'Sub '+name+'(ByVal x As Long)\nDim total As Long\n'+body+'\nEnd Sub'; }
const fixtures=[
 ...[100,1000,3000].map(count=>[count+'-operand-logical-chain','Private Const C = '+Array.from({length:count},()=> '1').join(' And ')]),
 ['100-consts/10-nested-sgn',Array.from({length:100},(_,i)=>'Private Const C'+i+' = '+'Sgn('.repeat(10)+'1'+')'.repeat(10)).join('\n')],
 ['300-procedures/arithmetic',Array.from({length:300},(_,i)=>procedure('P'+i,Array.from({length:12},(_,j)=>'total = x + '+j).join('\n'))).join('\n')],
 ['10000-statements/arithmetic',procedure('Main',Array.from({length:10000},(_,i)=>'total = x + '+i).join('\n'))],
 ['1000-statements/100-operands',procedure('Main',Array.from({length:1000},()=> 'total = '+Array.from({length:100},()=> '1').join(' + ')).join('\n'))],
 ['1000-statements/logical',procedure('Main',Array.from({length:1000},()=> 'total = 1 And 2 Or 3 Xor 4 Eqv 5 Imp 6').join('\n'))],
];
for(const [name,source] of fixtures) {
 const module=api.parseModule(source), symbols=api.buildModuleSymbols('Module','standard',source,{parsedModule:module});
 measure(name,()=>{let found=0;api.checkOverflow(source,module,symbols,undefined,undefined,undefined,()=>{found++;});if(found!==0)throw new Error('Unexpected diagnostics: '+found);});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
