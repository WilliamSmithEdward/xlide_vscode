import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {mkdtempSync,writeFileSync,existsSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {dirname,join} from 'node:path';
import {tmpdir,cpus} from 'node:os';
import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-condition-operand-scan.mjs [--baseline=COMMIT]
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const scratch=mkdtempSync(join(tmpdir(),'xlide-condition-operands-')),path=join(scratch,'bundle.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]diagnostics[\\/]conditionOperands\.ts$/},a=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/conditionOperands.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(a.path)}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export {conditionOperands} from './src/analyzer/diagnostics/conditionOperands';export {tokenize} from './src/analyzer/lexer/tokenize';export {parseModule} from './src/analyzer/parser/parseModule';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,result.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally {if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}

const workCounts=[],rows=[];
for(const count of [1,100,1000]){
 const expression=Array(count).fill('Flag').join(' And '),tokens=api.tokenize('If '+expression+' Then');
 const expected=Array.from({length:count},(_,i)=>({index:1+2*i,form:count===1?'condition':'logical'}));
 const original=Array.prototype.some;let duplicateComparisons=0;Array.prototype.some=function(callback,thisArg){return original.call(this,(value,index,array)=>{if(value&&typeof value.index==='number'&&typeof value.form==='string')duplicateComparisons++;return callback.call(thisArg,value,index,array);});};
 let actual;try{actual=api.conditionOperands(tokens);}finally{Array.prototype.some=original;}assert.deepEqual(actual,expected);if(!baseline)assert.equal(duplicateComparisons,0);workCounts.push({count,duplicateComparisons});
 const chunks=[];for(let i=0;i<count;i+=100)chunks.push(Array(Math.min(100,count-i)).fill('Flag').join(' And '));
 const condition=chunks.join(' And _\n'),source='Option Explicit\nSub Go()\nDim Flag As Boolean\nDim unusedArray(2) As Long\nunusedArray(0) = 1\nIf '+condition+' Then\nDebug.Print 1\nEnd If\nEnd Sub\n';assert.ok(Math.max(...source.split('\n').map(l=>l.length))<=1023);api.parseModule(source);
 for(const scope of ['helper','complete-module-diagnostics']){const samples=[];for(let round=-3;round<9;round++){const failures=[],begin=performance.now();const result=scope==='helper'?api.conditionOperands(tokens):api.analyzeModule(source,{onInternalError:e=>failures.push(String(e))});const elapsed=performance.now()-begin;assert.deepEqual(result,scope==='helper'?expected:[]);assert.deepEqual(failures,[]);if(round>=0)samples.push(elapsed);}samples.sort((a,b)=>a-b);rows.push({scope,count,medianMs:+samples[4].toFixed(5),outputsChecked:true});}
}
console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,workCounts,rows,scope:'Pre-tokenized helper and warmed module AST. Complete diagnostics independently expected empty, with an array local enabling condition-value scanning. Continued lines within physical limits. Counters execute before timings; no heap/cold/renderer claim.'},null,2));

