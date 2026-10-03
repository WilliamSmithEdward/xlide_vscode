// Run in the target checkout: node scripts/benchmark-parenthesis-validation.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-parenthesis-validation-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const result = buildSync({ stdin: { contents: `
        export { checkParentheses } from './src/analyzer/diagnostics/rules/parentheses';
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
for(const depth of [20,200,2000]) {
    const source='Sub Main()\nDim value As Long\nvalue = '+ '('.repeat(depth)+'1'+')'.repeat(depth)+'\nEnd Sub';
    const module=api.parseModule(source), symbols=api.buildModuleSymbols('Module','standard',source,{parsedModule:module});
    measure(depth+'-nested-parentheses',()=>{let found=0;api.checkParentheses(source,module,symbols,{},undefined,()=>{found++;});if(found!==0)throw new Error('Unexpected diagnostic');});
}
const source='Sub Main()\nDim value As Long\n'+'value = (1 + 2)\n'.repeat(1000)+'End Sub';
const module=api.parseModule(source), symbols=api.buildModuleSymbols('Module','standard',source,{parsedModule:module});
measure('1000-simple-groups',()=>{let found=0;api.checkParentheses(source,module,symbols,{},undefined,()=>{found++;});if(found!==0)throw new Error('Unexpected diagnostic');});
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
