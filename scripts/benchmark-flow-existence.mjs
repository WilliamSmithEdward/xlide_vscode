// Run in the target checkout: node scripts/benchmark-flow-existence.mjs [--rounds=15]
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
const scratch = mkdtempSync(join(tmpdir(), 'xlide-flow-existence-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
    const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
    const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/(procedureUnstructured|procedureLabels)\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/flow/'+args.path.split(/[\\/]/).pop()],{encoding:'utf8'}),loader:'ts',resolveDir:join(root,'src/analyzer/flow')}));}}]:[];
    const result = await build({ plugins, stdin: { contents:
        "export { procedureHasUnstructuredFlow } from './src/analyzer/flow/procedureUnstructured'; export { parseModule } from './src/analyzer/parser/parseModule';",
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
 for(const mode of ['none','early-reference','early-label','early-error','late-reference']) {
  const first=mode==='early-reference'?'GoTo Done\n':mode==='early-label'?'Entry:\n':mode==='early-error'?'On Error Resume Next\n':'';
  const source='Sub P()\n'+first+'x = x + 1\n'.repeat(count)+(mode==='late-reference'?'GoTo Done\n':'')+'End Sub';
  const proc=api.parseModule(source).members.find(member=>member.kind==='Procedure');
  measure(count+'-'+mode,()=>{const result=api.procedureHasUnstructuredFlow(source,{...proc});if(result!==(mode!=='none'))throw new Error('Wrong flow fact');});
  api.procedureHasUnstructuredFlow(source,proc);
  measure(count+'-'+mode+'-1000-cached-calls',()=>{for(let i=0;i<1000;i++)if(api.procedureHasUnstructuredFlow(source,proc)!==(mode!=='none'))throw new Error('Wrong cached fact');});
 }
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
