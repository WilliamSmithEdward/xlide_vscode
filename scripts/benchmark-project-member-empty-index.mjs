import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir, cpus } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { performance } from 'node:perf_hooks';
import assert from 'node:assert/strict';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const rounds=Number(process.argv.find(a=>a.startsWith('--rounds='))?.slice(9)??15);
if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3 to 100');
const onlyMetadata=process.argv.find(a=>a.startsWith('--metadata='))?.slice(11);
const onlyCount=process.argv.find(a=>a.startsWith('--references='))?.slice(13);
if(onlyMetadata&&!['absent','empty','ambiguous','class'].includes(onlyMetadata))throw Error('unknown metadata filter');
if(onlyCount&&!['1','100','1000'].includes(onlyCount))throw Error('unknown reference filter');
const dir=mkdtempSync(join(tmpdir(),'xlide-project-empty-')),file=join(dir,'api.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/memberAccess\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/completion/memberAccess.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export {projectTypeAt,projectClassMemberAt,privateMemberOwnerAt,projectClassMembersIndex} from './src/analyzer/completion/memberAccess';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';export {parseModule} from './src/analyzer/parser/parseModule';export {tokenizeCached} from './src/analyzer/lexer/tokenize';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,result.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
} finally {try{unlinkSync(file);}catch(e){if(e.code!=='ENOENT')throw e;}rmdirSync(dir);}
const surface=(name='C')=>({name,moduleName:name,kind:'class',members:[{name:'Value',kind:'property',returns:'Long',moduleName:name}],privateMembers:['Hidden']});
const metadata=name=>name==='absent'?undefined:name==='empty'?[]:name==='ambiguous'?[surface(),surface('c')]:name==='class'?[surface()]:[surface(),...Array.from({length:1000},(_,i)=>surface('Other'+i))];
const hash=v=>createHash('sha256').update(JSON.stringify(v)).digest('hex');
const rows=[];let salt=0;
for(const count of [1,100,1000].filter(n=>!onlyCount||String(n)===onlyCount))for(const name of ['absent','empty','ambiguous','class'].filter(n=>!onlyMetadata||n===onlyMetadata))for(const scope of ['direct','full','full-fresh']){
 const source='Option Explicit\nSub Main()\nDim c As '+(name==='class'?'C':'Range')+'\n'+Array.from({length:count},()=> 'Debug.Print c.Value\n').join('')+'End Sub\n';
 const offsets=[...source.matchAll(/c\.Value/g)].map(m=>m.index+2),module=api.parseModule(source),tokens=api.tokenizeCached(source).filter(t=>t.kind!=='comment');
 const samples=[];let signature;
 for(let round=-3;round<rounds;round++){
  const projectClassMembers=metadata(name);let result;
  const ctx={projectClassMembers,parsedModule:module,sourceTokens:tokens};
  if(scope==='direct')api.projectClassMembersIndex(projectClassMembers??[]);
  const start=performance.now();
  if(scope==='direct')result=offsets.map(at=>[api.projectTypeAt(source,at,ctx),api.projectClassMemberAt(source,at,'Value',ctx),api.privateMemberOwnerAt(source,at,'Hidden',ctx)]);
  else result=api.analyzeModule(scope==='full-fresh'?source+"' sample "+(++salt)+'\n':source,{knownIdentifiers:new Set(),projectClassMembers});
  const elapsed=performance.now()-start,current=hash(result);
  if(signature)assert.equal(current,signature);else signature=current;
  if(scope==='direct')for(const triple of result)assert.deepEqual(triple,name==='class'?[projectClassMembers[0],projectClassMembers[0].members[0],'C']:[undefined,undefined,undefined]);
  if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({count,metadata:name,scope,medianMs:samples[Math.floor(rounds/2)],p95Ms:samples[Math.ceil(rounds*.95)-1],resultSignature:signature});
}
for(const phase of onlyMetadata||onlyCount?[]:['cold','warm']){
 const source='Option Explicit\nSub Main()\nEnd Sub\n',samples=[];
 for(let round=-3;round<rounds;round++){
  const ctx={projectClassMembers:metadata('large')};if(phase==='warm')api.projectClassMembersIndex(ctx.projectClassMembers);
  const start=performance.now();const result=[api.projectTypeAt(source,0,ctx),api.projectClassMemberAt(source,0,'Value',ctx),api.privateMemberOwnerAt(source,0,'Hidden',ctx)];const elapsed=performance.now()-start;
  assert.deepEqual(result,[undefined,undefined,undefined]);if(round>=0)samples.push(elapsed);
 }
 samples.sort((a,b)=>a-b);rows.push({count:1,metadata:'1001classes-invalid-offset',scope:phase,medianMs:samples[Math.floor(rounds/2)],p95Ms:samples[Math.ceil(rounds*.95)-1]});
}
console.log(JSON.stringify({node:process.version,cpu:cpus()[0].model,baseline:baseline??null,rounds,warmups:3,scope:'Actual public project lookups with parsed AST/shared tokens and pre-indexed metadata; full analyzeModule has fresh metadata per round; full-fresh also changes the source key; absent/empty/ambiguous fixtures use host Range, positive class uses C; assertions outside clock. Cold control times first metadata index. No editor/Office timing.',rows},null,2));
