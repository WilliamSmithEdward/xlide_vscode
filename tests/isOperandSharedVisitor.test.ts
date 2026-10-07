import { afterEach, expect, it, vi } from 'vitest';
import { checkIsOperandsInConditions } from '../src/analyzer/diagnostics/rules/typeOfIs';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import * as walk from '../src/analyzer/diagnostics/exprWalk';
afterEach(() => vi.restoreAllMocks());
function fixture(body: string, type = 'Long') {
 const source = ['Option Explicit','Sub Go()','Dim n As '+type,body,'End Sub',''].join('\n');
 const mod = parseModule(source), symbols = buildModuleSymbols('M','standard',source,{parsedModule:mod});
 return {source,mod,symbols};
}
function run(f: ReturnType<typeof fixture>) { const out: unknown[][]=[]; checkIsOperandsInConditions(f.source,f.mod,f.symbols,undefined,(...v)=>out.push(v));return out; }
function expected(f: ReturnType<typeof fixture>, count: number) {
 let from = 0; return Array.from({length:count},()=> { const start = f.source.indexOf('n Is Nothing',from);from=start+1;return ['isOperatorNonObject',"The 'Is' operator requires object operands, but 'n' is declared As Long, which is not an object.",{start,end:start+1}];});
}
for (const n of [10,100,1000]) it(`reuses one visitor and removes dead comment filtering for ${n} conditions`, () => {
 const f=fixture(Array(n).fill("If n Is Nothing Then Debug.Print n ' tail").join('\n'));
 const spy=vi.spyOn(walk,'forEachSubExpression'), original=Array.prototype.filter;let elements=0;
 try {
  Array.prototype.filter=function(callback:any,thisArg?:any):any {const stack=new Error().stack??'';if(/\.kind\s*!==\s*['"]comment/.test(String(callback))&&stack.includes('checkIsOperandsInConditions')&&!/\bstatementTokens(?:Cached)?\b/.test(stack))elements+=this.length;return original.call(this,callback,thisArg);};
  expect(run(f)).toEqual(expected(f,n));
 } finally {Array.prototype.filter=original;}
 expect(spy).toHaveBeenCalledTimes(n);expect.soft(new Set(spy.mock.calls.map(call=>call[1])).size).toBe(1);expect.soft(elements).toBe(0);
});
for(const body of ['If n Is Nothing Then Debug.Print n','Do While n Is Nothing\nExit Do\nLoop','Do Until n Is Nothing\nExit Do\nLoop','Do\nDebug.Print n\nLoop While n Is Nothing','Do\nDebug.Print n\nLoop Until n Is Nothing','While n Is Nothing\nDebug.Print n\nWend'])
 it(`preserves exact scalar diagnostic for ${body.split('\n')[0]}`,()=>{const f=fixture(body);expect(run(f)).toEqual(expected(f,1));});
for(const type of ['Object','Variant','vbLong']) it(`preserves quiet ${type} controls`,()=>expect(run(fixture('If n Is Nothing Then Debug.Print n',type))).toEqual([]));
it('keeps visitor environments separate across procedures',()=>{
 const source='Sub A()\nDim n As Long\nIf n Is Nothing Then Debug.Print n\nEnd Sub\nSub B()\nDim n As Object\nIf n Is Nothing Then Debug.Print n\nEnd Sub';
 const mod=parseModule(source),symbols=buildModuleSymbols('M','standard',source,{parsedModule:mod});
 expect(run({source,mod,symbols})).toEqual(expected({source,mod,symbols},1));
});
