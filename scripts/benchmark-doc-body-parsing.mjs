// Run in the target checkout: node scripts/benchmark-doc-body-parsing.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-doc-body-parsing-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
    const plugins = baseline ? [{name:'baseline',setup(builder) {
        builder.onLoad({filter:/docComment\.ts$/},()=>({
            contents:execFileSync('git',['show',baseline+':src/analyzer/docs/docComment.ts'],{encoding:'utf8'}),
            loader:'ts',resolveDir:join(root,'src/analyzer/docs'),
        }));
    }}] : [];
    const result = await build({ plugins, stdin: { contents:
        "export { parseDocBody } from './src/analyzer/docs/docComment';",
        resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
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
for (const count of [100,1000,10000]) {
    const cases=[
        ['unclosed-params',Array.from({length:count},(_,i)=>'<param name="P'+i+'">text').join('\n'),0],
        ['incomplete-params','<param name="A" '.repeat(count),0],
        ['unclosed-summaries','<summary>text\n'.repeat(count),0],
        ['paired-params',Array.from({length:count},(_,i)=>'<param name="P'+i+'">text</param>').join('\n'),count],
    ];
    for (const [name,body,expected] of cases) measure(count+'-'+name,()=>{
        const doc=api.parseDocBody(body,'inline');
        if(doc.params.length!==expected)throw new Error('Wrong parsing result');
    });
}
const ordinary='<summary>Computes tax.</summary><param name="Amount" type="Currency">The subtotal.</param><param name="Rate">Tax rate.</param><returns type="Currency">Tax owed.</returns><remarks>Rounds to cents.</remarks><example>Dim x As Long\nx = Tax(100, .2)</example>';
measure('1000-ordinary-docs',()=>{for(let i=0;i<1000;i++){const doc=api.parseDocBody(ordinary,'inline');if(doc.params.length!==2||doc.summary!=='Computes tax.')throw new Error('Wrong ordinary doc');}});
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
