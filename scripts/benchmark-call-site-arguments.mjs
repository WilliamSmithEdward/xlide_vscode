// Run: node scripts/benchmark-call-site-arguments.mjs [--baseline=COMMIT] [--rounds=9] [--date-literals] [--recursive-refactor]
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
const recursiveRefactor = process.argv.includes('--recursive-refactor');
const scratch = mkdtempSync(join(tmpdir(), 'xlide-call-site-bounds-'));
const bundle = join(scratch, 'calls.cjs');
let api;
try {
 const plugins = baseline ? [{name:'baseline',setup(builder) {
  builder.onLoad({filter:recursiveRefactor ? /(?:callSites|introduceParameter|memberAccess)\.ts$/ : /callSites\.ts$/},args => ({contents:execFileSync('git',['show',baseline+':'+args.path.slice(root.length+1).replaceAll('\\','/')],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));
 }}] : [];
 const result = await build({plugins,stdin:{contents:recursiveRefactor ? "export {introduceParameter} from './src/analyzer/refactor/introduceParameter';export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';" : "export {callSitesOf} from './src/analyzer/refactor/callSites';export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(bundle,result.outputFiles[0].contents);
 api = createRequire(import.meta.url)(bundle);
} finally { try { unlinkSync(bundle); } catch(error) { if(error.code !== 'ENOENT') throw error; } rmdirSync(scratch); }
if(recursiveRefactor) {
 const rows=[];
 let salt=0;
 for(const count of [1,100,1000]) for(const layout of ['bare','Me','WithMe']) for(const mode of ['cached','fresh']) {
  const call='If depth > 0 Then '+(layout==='Me'?'Me.':layout==='WithMe'?'.':'')+'Report depth - 1';
  const calls=Array(count).fill(call).join('\n');
  const statements=layout==='WithMe'?'With Me\n'+calls+'\nEnd With':calls;
  const expectedStatements=layout==='WithMe'?'With Me\n'+Array(count).fill(call+', 3').join('\n')+'\nEnd With':Array(count).fill(call+', 3').join('\n');
  const body='Public Sub Report(ByVal depth As Long)\nDim limit As Long\nlimit = 3\n'+statements+'\nEnd Sub\n';
  const expectedBody='Public Sub Report(ByVal depth As Long, ByVal limit As Long)\n'+expectedStatements+'\nEnd Sub\n';
  const samples=[];
  let result,source;
  for(let round=-3;round<rounds;round++) {
   source='Option Explicit\n'+(mode==='fresh'?"' run "+(++salt)+'\n':'')+body;
   const start=performance.now();
   result=api.introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',moduleKinds:{Module1:'class'}});
   const elapsed=performance.now()-start;
   if(!result.ok)throw new Error(result.reason);
   if(round>=0)samples.push(elapsed);
  }
  const rendered=api.applyVbaTextEdits(source,result.edits).slice(source.indexOf('Public Sub'));
  const correctRender=rendered===expectedBody;
  const actualCalls=result.edits.filter(edit=>edit.newText===', 3').length;
  if(!baseline&&(!correctRender||actualCalls!==count))throw new Error('Incorrect recursive refactor: '+layout+'/'+count);
  samples.sort((a,b)=>a-b);
  rows.push({name:[count,layout,mode].join('/'),count,layout,mode,actualCalls,correctRender,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
 }
 console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
} else {
const dateLiterals=process.argv.includes("--date-literals");
const layouts=dateLiterals ? ["month", "fileNumber", "number"] : ["number","string","bracketed","nested","colon","comments"];
const procedureName=dateLiterals ? "May" : "Go";
const rows=[];
let salt=0;
for(const count of [1,100,1000]) for(const layout of layouts) for(const mode of ['cached','fresh']) {
 const call=layout==='month'?'Debug.Print #May 1, 2000#\nMay 1':layout==='fileNumber'?'Print #channel, May(1)':dateLiterals?'May 1':layout==='string'?'Go "hello"':layout==='bracketed'?'Call Go(1)':'Go 1';
 const statement=layout==='nested'?'x = '+'Go('.repeat(count)+'1'+')'.repeat(count):layout==='colon'?Array(count).fill(call).join(': '):layout==='comments'?Array.from({length:count},(_,i)=>"Go 1 'note "+i).join('\n'):Array(count).fill(call).join('\n');
 const expectedStatement=layout==='nested'?'x = '+'Go('.repeat(count)+'1'+', 3)'.repeat(count):layout==='colon'?Array(count).fill('Go 1, 3').join(': '):layout==='comments'?Array.from({length:count},(_,i)=>"Go 1, 3 'note "+i).join('\n'):Array(count).fill(layout==='month'?'Debug.Print #May 1, 2000#\nMay 1, 3':layout==='fileNumber'?'Print #channel, May(1, 3)':dateLiterals?'May 1, 3':layout==='string'?'Go "hello", 3':layout==='bracketed'?'Call Go(1, 3)':'Go 1, 3').join('\n');
 const body='Sub Caller()\n'+statement+'\nEnd Sub\n';
 const expectedBody='Sub Caller()\n'+expectedStatement+'\nEnd Sub\n';
 const cached='Option Explicit\n'+body;
 const samples=[];
 let sites,source;
 for(let round=-3;round<rounds;round++) {
  source=mode==='cached'?cached:'Option Explicit\n'+"' run "+(++salt)+'\n'+body;
  const start=performance.now();sites=api.callSitesOf(source,procedureName);const elapsed=performance.now()-start;
  if(sites.length!==count && !(baseline&&dateLiterals))throw new Error('Unexpected call count');
  if(round>=0)samples.push(elapsed);
 }
 let rendered,renderError;
 try {rendered=api.applyVbaTextEdits(source,sites.map(site=>({span:site.argumentInsert,newText:site.argumentText('3','added')})));rendered=rendered.slice(rendered.indexOf('Sub Caller()'));}
 catch(error){renderError=error.message;}
 const correctRender=rendered===expectedBody;
 if(!baseline&&!correctRender)throw new Error('Candidate produces an incorrect edit: '+layout);
 const signature=text=>createHash('sha256').update(text??'render-error:'+renderError).digest('hex');
 samples.sort((a,b)=>a-b);
 rows.push({name:[count,layout,mode].join('/'),count,layout,mode,actualCalls:sites.length,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5),correctRender,renderError,renderedHash:signature(rendered),expectedHash:signature(expectedBody)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
}
