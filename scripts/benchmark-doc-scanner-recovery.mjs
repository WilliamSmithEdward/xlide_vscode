// Run in the target checkout: node scripts/benchmark-doc-scanner-recovery.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-doc-scanner-recovery-'));
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
        "export { parseDocBody, scanDocTags } from './src/analyzer/docs/docComment';",
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
for(const count of [100,1000,10000]) {
 const body='<param name="Good" '+'x'.repeat(count)+' />';
 measure(count+'-unrecognized-attribute-word',()=>{const doc=api.parseDocBody(body,'inline');if(doc.params.length!==1)throw new Error('Wrong attribute recovery');});
 const lines=Array.from({length:count},(_,i)=>({start:i*16,directivesStart:i*16,textStart:i*16+4,text:'<param'}));
 measure(count+'-incomplete-diagnostic-openings',()=>{const tags=api.scanDocTags(lines);if(!tags||tags.length)throw new Error('Wrong opening recovery');});
}
const ordinary='<summary>Compute tax.</summary><param name="A" type="Long" unit="N">First.</param><param name="B">Second.</param><returns type="Long">Total.</returns>';
const lines=[{start:0,directivesStart:0,textStart:4,text:ordinary}];
measure('1000-ordinary-models-and-tag-scans',()=>{for(let i=0;i<1000;i++){const doc=api.parseDocBody(ordinary,'inline'),tags=api.scanDocTags(lines);if(doc.params.length!==2||tags.length!==4)throw new Error('Wrong ordinary result');}});
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
