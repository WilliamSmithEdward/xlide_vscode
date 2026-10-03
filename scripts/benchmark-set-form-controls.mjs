// Run: node scripts/benchmark-set-form-controls.mjs [--baseline=<commit>] [--rounds=15]
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.slice(9) ?? 15);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw Error('rounds must be 3 to 100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-set-form-controls-'));
const bundle = join(scratch, 'analyzer.cjs');
let api;
try {
 const plugins = baseline ? [{ name: 'baseline', setup(builder) {
  builder.onLoad({ filter: /rules[\\/]assignments\.ts$/ }, () => ({ contents: execFileSync('git', ['show', baseline + ':src/analyzer/diagnostics/rules/assignments.ts'], { encoding: 'utf8' }), loader: 'ts', resolveDir: join(root, 'src/analyzer/diagnostics/rules') }));
 }}] : [];
 const result = await build({ plugins, stdin: { contents: "export { hostObjectModelForToken } from './src/analyzer/host/hostRegistry'; export { checkSetAssignments } from './src/analyzer/diagnostics/rules/assignments'; export { forEachStatementWithHeaders } from './src/analyzer/diagnostics/walker'; export { parseModule } from './src/analyzer/parser/parseModule'; export { buildModuleSymbols } from './src/analyzer/symbols/buildModuleSymbols';", resolveDir: root, loader: 'ts' }, bundle: true, platform: 'node', format: 'cjs', write: false });
 writeFileSync(bundle, result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch (error) { if (error.code !== 'ENOENT') throw error; } rmdirSync(scratch); }
const rows=[];
for(const count of [10,100,1000,3000])for(const mode of ['form-class-value','form-member-value','form-member-target','form-missing-value','scalar-control']) {
 const classes=mode==='form-class-value'?Array.from({length:count},(_,i)=>({name:'F'+i,kind:'userform',moduleName:'F'+i,members:[]})):[];
 const members=mode!=='form-class-value'?Array.from({length:count},(_,i)=>({name:'C'+i,kind:'property',moduleName:'Form1',returns:'MSForms.TextBox'})):[];
 members.push({name:'Answer',kind:'property',moduleName:'Form1',returns:'MSForms.Label'});classes.push({name:'Form1',kind:'userform',moduleName:'Form1',members});
 const line=mode==='form-member-target'?'Set Answer = Nothing\n':mode==='form-missing-value'?'Set target = UnknownControl\n':mode==='scalar-control'?'Set n = Nothing\n':'Set target = Answer\n';
 const source='Sub P()\nDim target As MSForms.TextBox\nDim n As Long\n'+line.repeat(1000)+'End Sub';const mod=api.parseModule(source);let previousRoot;
 function prepare(){const symbols=api.buildModuleSymbols('Form1','userform',source,{parsedModule:mod});if(symbols.root===previousRoot)throw Error('Reused root');previousRoot=symbols.root;const ctx=mode==='scalar-control'?{}:{meProjectType:'Form1',projectClassMembers:classes};return ()=>{let hits=0;const visitor=api.checkSetAssignments(source,symbols,undefined,ctx,()=>hits++);for(const proc of mod.members)if(proc.kind==='Procedure'){const visit=visitor(proc);if(visit)api.forEachStatementWithHeaders(source,proc.body,visit);}if(hits!==(mode==='form-missing-value'?0:1000))throw Error('Wrong '+mode+' hits '+hits);};}
 for(let i=0;i<3;i++)prepare()();const samples=[];for(let i=0;i<rounds;i++){const run=prepare(),start=performance.now();run();samples.push(performance.now()-start);}samples.sort((a,b)=>a-b);
 rows.push({name:count+'-'+mode,medianMs:+samples[Math.floor(rounds/2)].toFixed(3),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(3)});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
