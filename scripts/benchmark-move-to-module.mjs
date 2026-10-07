// Run: node scripts/benchmark-move-to-module.mjs [--baseline=COMMIT] [--rounds=9]
import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {writeFileSync,mkdtempSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {tmpdir,cpus} from 'node:os';
import {dirname,join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {performance} from 'node:perf_hooks';
const root=dirname(dirname(fileURLToPath(import.meta.url)));
const baseline=process.argv.find(arg=>arg.startsWith('--baseline='))?.slice(11);
const rounds=Number(process.argv.find(arg=>arg.startsWith('--rounds='))?.slice(9)??9);
if(!Number.isInteger(rounds)||rounds<3||rounds>100)throw Error('rounds must be 3..100');
const scratch=mkdtempSync(join(tmpdir(),'xlide-move-module-'));
const file=join(scratch,'api.cjs');
let api;
try {
 const plugins=baseline?[{name:'baseline',setup(builder){builder.onLoad({filter:/moveToModule\.ts$/},args=>({contents:execFileSync('git',['show',baseline+':src/analyzer/refactor/moveToModule.ts'],{cwd:root,encoding:'utf8'}),loader:'ts',resolveDir:dirname(args.path)}));}}]:[];
 const built=await build({plugins,stdin:{contents:"export {moveToModule} from './src/analyzer/refactor/moveToModule';export {applyVbaTextEdits} from './src/analyzer/refactor/refactorTypes';",resolveDir:root,loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});
 writeFileSync(file,built.outputFiles[0].contents);api=createRequire(import.meta.url)(file);
}finally{try{unlinkSync(file);}catch(error){if(error.code!=='ENOENT')throw error;}rmdirSync(scratch);}
const rows=[];let salt=0;
for(const count of [1,100,1000])for(const layout of ['colon','multiline','comments'])for(const mode of ['cached','fresh']){
 const call='Reports.Build';
 const calls=layout==='comments'?Array(count).fill(call+" ' Reports.Build").join('\n'):Array(count).fill(call).join(layout==='colon'?': ':'\n');
 const expectedCalls=layout==='comments'?Array(count).fill('Helpers.Build'+" ' Reports.Build").join('\n'):Array(count).fill('Helpers.Build').join(layout==='colon'?': ':'\n');
 const source='Public Sub Build()\nEnd Sub\n';
 const body='Sub Caller()\n'+calls+'\nEnd Sub\n';
 const expectedBody='Sub Caller()\n'+expectedCalls+'\nEnd Sub\n';
 const samples=[];let result,caller;
 for(let round=-3;round<rounds;round++){
  caller=(mode==='fresh'?"' run "+(++salt)+'\n':'')+body;
  const start=performance.now();result=api.moveToModule({source,offset:source.indexOf('Build'),moduleName:'Reports',targetModuleName:'Helpers',otherModuleSources:{Helpers:'',Caller:caller}});const elapsed=performance.now()-start;
  if(!result.ok)throw Error(result.reason);if(round>=0)samples.push(elapsed);
 }
 const changes=result.otherModules.find(module=>module.moduleName==='Caller');
 const rendered=api.applyVbaTextEdits(caller,changes.edits).slice(caller.indexOf('Sub Caller()'));
 const target=result.otherModules.find(module=>module.moduleName==='Helpers');
 const correctRender=rendered===expectedBody&&api.applyVbaTextEdits(source,result.edits)===''&&api.applyVbaTextEdits('',target.edits)==='\n'+source;
 if(!correctRender)throw Error('Incorrect complete move '+layout);
 samples.sort((a,b)=>a-b);rows.push({name:[count,layout,mode].join('/'),count,layout,mode,correctRender,medianMs:+samples[Math.floor(rounds/2)].toFixed(5),p95Ms:+samples[Math.ceil(rounds*.95)-1].toFixed(5)});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds,rows},null,2));
