import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {mkdtempSync,writeFileSync,existsSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {dirname,join} from 'node:path';
import {tmpdir,cpus} from 'node:os';
import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-initializer-name-membership.mjs [--baseline=COMMIT]
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const scratch=mkdtempSync(join(tmpdir(),'xlide-initializer-names-')),path=join(scratch,'bundle.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]refactor[\\/]introduceParameter\.ts$/},a=>({contents:execFileSync('git',['show',baseline+':src/analyzer/refactor/introduceParameter.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(a.path)}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export {introduceParameter} from './src/analyzer/refactor/introduceParameter';export {parseModule} from './src/analyzer/parser/parseModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,result.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally {if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}

const workCounts=[],rows=[];
for(const count of [1,100,1000,2000]){
 const names=Array.from({length:count},(_,i)=>'v'+i),chunks=[];for(let i=0;i<count;i+=100)chunks.push(names.slice(i,i+100).join(' + '));
 const source=[...names.map(n=>'Private '+n+' As Long'),'Sub Report()','Dim limit As Long','limit = '+chunks.join(' + _\n'),'Debug.Print limit','End Sub',''].join('\n');assert.ok(Math.max(...source.split('\n').map(l=>l.length))<=1023);
 const input={source,offset:source.indexOf('limit As'),moduleName:'M'},expected={ok:false,reason:"The value '"+names.join(' + ')+"' names "+names.map(n=>"'"+n+"'").join(', ')+', which a caller in another module cannot see.'};api.parseModule(source);
 const original=Array.prototype.includes;let searchedSlots=0;Array.prototype.includes=function(value,...rest){if(typeof value==='string'&&/^v\d+$/.test(value)&&this.every(v=>typeof v==='string'&&/^v\d+$/.test(v)))searchedSlots+=this.length;return original.call(this,value,...rest);};let result;try{result=api.introduceParameter(input);}finally{Array.prototype.includes=original;}assert.deepEqual(result,expected);if(!baseline)assert.equal(searchedSlots,0);workCounts.push({count,searchedSlots});
 const samples=[];for(let round=-3;round<9;round++){const begin=performance.now(),actual=api.introduceParameter(input),elapsed=performance.now()-begin;assert.deepEqual(actual,expected);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({count,medianMs:+samples[4].toFixed(5),completeRefusalChecked:true});
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,workCounts,rows,scope:'Complete Introduce Parameter refusal, warmed AST. Valid continued physical lines; exact reason independently specified. Search slots count growing array lengths for distinct names, not heap bytes. No cold or editor claim.'},null,2));


