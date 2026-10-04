// Run: node scripts/benchmark-call-site-bounds.mjs [--baseline=COMMIT] [--rounds=9]
import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { writeFileSync, mkdtempSync, unlinkSync, rmdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir, cpus } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const baseline = process.argv.find(arg => arg.startsWith('--baseline='))?.slice(11);
const rounds = Number(process.argv.find(arg => arg.startsWith('--rounds='))?.slice(9) ?? 9);
if (!Number.isInteger(rounds) || rounds < 3 || rounds > 100) throw new Error('rounds must be 3..100');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-call-site-bounds-'));
const bundle = join(scratch, 'calls.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(builder) {
  builder.onLoad({filter:/callSites\.ts$/},args => ({contents:execFileSync('git',['show',baseline+':src/analyzer/refactor/callSites.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));
 }}] : [];
 const result = await build({plugins,stdin:{contents:"export {callSitesOf} from './src/analyzer/refactor/callSites';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(bundle,result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch(error) { if(error.code !== 'ENOENT') throw error; } rmdirSync(scratch); }
const rows=[];
let salt=0;
for(const count of [1,100,1000]) for(const [ending,eol] of [['lf','\n'],['crlf','\r\n'],['cr','\r']]) for(const form of ['bare','bracketed']) for(const mode of ['cached','fresh']) {
 const body=['Sub Caller()',...Array(count).fill(form==='bare'?'Go 1':'Call Go(1)'), 'End Sub',''].join(eol);
 const cached='Option Explicit'+eol+body;
 const samples=[];
 let resultSignature;
 for(let round=-3;round<rounds;round++) {
  const source=mode==='cached'?cached:'Option Explicit'+eol+"' run "+(++salt)+eol+body;
  const start=performance.now();
  const sites=api.callSitesOf(source,'Go');
  const elapsed=performance.now()-start;
  if(sites.length!==count) throw new Error('Unexpected call count');
  // Offsets differ in fresh mode; compare complete relative insertion metadata and text.
  const serialized=sites.map(site=>({relativeStart:site.argumentInsert.start-site.offset,relativeEnd:site.argumentInsert.end-site.offset,newText:site.argumentText('3'),bracketed:site.bracketed,empty:site.empty}));
  const hash=createHash('sha256').update(JSON.stringify(serialized)).digest('hex');
  if(resultSignature && resultSignature!==hash)throw new Error('Unstable results');
  resultSignature=hash;
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);
 rows.push({name:[count,ending,form,mode].join('/'),count,ending,form,mode,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5),resultSignature});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
