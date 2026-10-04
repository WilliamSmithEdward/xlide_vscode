import {build} from 'esbuild';
import {execFileSync} from 'node:child_process';
import {mkdtempSync,writeFileSync,existsSync,unlinkSync,rmdirSync} from 'node:fs';
import {createRequire} from 'node:module';
import {dirname,join} from 'node:path';
import {tmpdir,cpus} from 'node:os';
import {performance} from 'node:perf_hooks';
import assert from 'node:assert/strict';
// node scripts/benchmark-null-operator-stack.mjs [--baseline=COMMIT]
const baseline=process.argv.find(a=>a.startsWith('--baseline='))?.slice(11);
const scratch=mkdtempSync(join(tmpdir(),'xlide-null-stack-')),path=join(scratch,'bundle.cjs');let api;
try {
 const plugins=baseline?[{name:'baseline',setup(b){b.onLoad({filter:/[\\/]diagnostics[\\/]nullOperators\.ts$/},a=>({contents:execFileSync('git',['show',baseline+':src/analyzer/diagnostics/nullOperators.ts'],{encoding:'utf8'}),loader:'ts',resolveDir:dirname(a.path)}));}}]:[];
 const result=await build({plugins,stdin:{contents:"export {operatorYieldsNull} from './src/analyzer/diagnostics/nullOperators';export {tokenize} from './src/analyzer/lexer/tokenize';export {parseModule} from './src/analyzer/parser/parseModule';export {analyzeModule} from './src/analyzer/diagnostics/analyzeModule';",resolveDir:process.cwd(),loader:'ts'},bundle:true,platform:'node',format:'cjs',write:false});writeFileSync(path,result.outputFiles[0].contents);api=createRequire(import.meta.url)(path);
} finally {if(existsSync(path))unlinkSync(path);rmdirSync(scratch);}
const workCounts=[];
for(const count of [100,1000]){
 const unary=api.tokenize('x='+'Not '.repeat(count)+'Null').slice(2),slice=Array.prototype.slice;let copiedTokenReferences=0;
 Array.prototype.slice=function(start,end){const result=slice.call(this,start,end);copiedTokenReferences+=result.length;return result;};
 let actual;try{actual=api.operatorYieldsNull(unary,t=>t.rawText.toLowerCase()==='null');}finally{Array.prototype.slice=slice;}
 assert.equal(actual,true);if(!baseline)assert.ok(copiedTokenReferences<=unary.length);
 const original=api.tokenize('x='+'1 + ('.repeat(count)+'Null'+')'.repeat(count)).slice(2);let rawTextReads=0;
 const arithmetic=original.map(t=>({...t,get rawText(){rawTextReads++;return t.rawText;}}));
 assert.equal(api.operatorYieldsNull(arithmetic,t=>t.rawText.toLowerCase()==='null'),true);
 if(!baseline)assert.ok(rawTextReads<=arithmetic.length*40);
 workCounts.push({count,unaryTokens:unary.length,copiedTokenReferences,arithmeticTokens:arithmetic.length,rawTextReads});
}
const rows=[];
for(const kind of ['unary','abs','arithmetic'])for(const count of [1,100,1000]){
 const expression=kind==='unary'?'Not '.repeat(count)+'Null':kind==='abs'?'Abs('.repeat(count)+'Null'+')'.repeat(count):'1 + ('.repeat(count)+'Null'+')'.repeat(count);
 const tokens=api.tokenize('x='+expression).slice(2),samples=[];
 for(let round=-3;round<9;round++){const begin=performance.now(),value=api.operatorYieldsNull(tokens,t=>t.rawText.toLowerCase()==='null'),elapsed=performance.now()-begin;assert.equal(value,true);if(round>=0)samples.push(elapsed);}
 samples.sort((a,b)=>a-b);rows.push({kind,count,scope:'Null-helper',medianMs:+samples[4].toFixed(5),outputsChecked:true});
}
for(const count of [1,100,1000]){
 const chunks=[];for(let remaining=count;remaining>0;remaining-=250)chunks.push('Not '.repeat(Math.min(remaining,250)).trimEnd());
 const expression=chunks.join(' _\n')+' Null',display='Not '.repeat(count)+'Null',source='Option Explicit\nSub Go()\nDim value As Long\nvalue = '+expression+'\nDebug.Print value\nEnd Sub\n',start=source.indexOf(expression);api.parseModule(source);
 assert.ok(Math.max(...source.split('\n').map(line=>line.length))<=1023);
 const expected=[{code:'assignment-type-mismatch',message:"Assignment to 'value' expects Long, but '"+display+"' is Null: an operator on Null gives Null. Null cannot be coerced to this scalar type. This will raise Run-time error '94': Invalid use of Null.",severity:'error',span:{start,end:start+expression.length},specReference:'MS-VBAL 5.4.3 / runtime type coercion and numeric overflow',origin:'run'}];
 const samples=[];for(let round=-3;round<9;round++){const failures=[],begin=performance.now(),actual=api.analyzeModule(source,{onInternalError:e=>failures.push(String(e))}),elapsed=performance.now()-begin;assert.deepEqual(failures,[]);assert.deepEqual(actual,expected);if(round>=0)samples.push(elapsed);}
 samples.sort((a,b)=>a-b);rows.push({kind:'unary',count,scope:'complete-module-diagnostics',medianMs:+samples[4].toFixed(5),completeOutputsChecked:true});
}
assert.equal(rows.length,12);
assert.equal(new Set(rows.map(row=>row.scope+':'+row.kind+':'+row.count)).size,12);

console.log(JSON.stringify({baseline:baseline??null,node:process.version,cpu:cpus()[0]?.model,rounds:9,warmups:3,workCounts,rows,scope:'Helper tokenization and module AST warming outside clock. Every helper result and independently specified complete diagnostic objects/failures checked; module sources use continued physical lines within line limits. Baseline successful depths only; deep failure regression is independently tested. No cold-parser, heap-byte or renderer claim.'},null,2));
