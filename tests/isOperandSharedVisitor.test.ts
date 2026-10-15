import { afterEach, expect, it, vi } from 'vitest';
import { checkIsOperatorOperands } from '../src/analyzer/diagnostics/rules/typeOfIs';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import * as walk from '../src/analyzer/diagnostics/exprWalk';
afterEach(() => vi.restoreAllMocks());
function fixture(body: string, type = 'Long') {
 const source = ['Option Explicit','Sub Go()','Dim n As '+type,body,'End Sub',''].join('\n');
 const mod = parseModule(source), symbols = buildModuleSymbols('M','standard',source,{parsedModule:mod});
 return {source,mod,symbols};
}
function run(f: ReturnType<typeof fixture>) { const out: unknown[][]=[]; walk.walkProcedureExpressions(f.mod,undefined,[checkIsOperatorOperands(f.symbols,(...v)=>out.push(v))],f.source);return out; }
function expected(f: ReturnType<typeof fixture>, count: number) {
 let from = 0; return Array.from({length:count},()=> { const start = f.source.indexOf('n Is Nothing',from);from=start+1;return ['isOperatorNonObject',"The 'Is' operator requires object operands, but 'n' has type Long, which is not an object.",{start,end:start+1}];});
}
for (const n of [10,100,1000]) it(`shares one visitor across ${n} conditions and print items`, () => {
 const f=fixture(Array(n).fill("If n Is Nothing Then Debug.Print n ' tail").join('\n'));
 const out: unknown[][]=[];
 const rule=checkIsOperatorOperands(f.symbols,(...v)=>out.push(v));
 const visits=vi.fn();
 const factory=vi.fn((member: Parameters<typeof rule>[0])=>{const check=rule(member);return (expr: Parameters<typeof check>[0])=>{visits(expr);check(expr);};});
 walk.walkProcedureExpressions(f.mod,undefined,[factory],f.source);
 expect(out).toEqual(expected(f,n));
 expect(factory).toHaveBeenCalledTimes(1);
 expect(visits).toHaveBeenCalledTimes(n*4);
 expect(visits.mock.calls.filter(([expr])=>expr.exprKind==='BinaryExpr')).toHaveLength(n);
});
for(const body of ['If n Is Nothing Then Debug.Print n','Do While n Is Nothing\nExit Do\nLoop','Do Until n Is Nothing\nExit Do\nLoop','Do\nDebug.Print n\nLoop While n Is Nothing','Do\nDebug.Print n\nLoop Until n Is Nothing','While n Is Nothing\nDebug.Print n\nWend'])
 it(`preserves exact scalar diagnostic for ${body.split('\n')[0]}`,()=>{const f=fixture(body);expect(run(f)).toEqual(expected(f,1));});
for(const type of ['Object','Variant','vbLong']) it(`preserves quiet ${type} controls`,()=>expect(run(fixture('If n Is Nothing Then Debug.Print n',type))).toEqual([]));
it('keeps visitor environments separate across procedures',()=>{
 const source='Sub A()\nDim n As Long\nIf n Is Nothing Then Debug.Print n\nEnd Sub\nSub B()\nDim n As Object\nIf n Is Nothing Then Debug.Print n\nEnd Sub';
 const mod=parseModule(source),symbols=buildModuleSymbols('M','standard',source,{parsedModule:mod});
 expect(run({source,mod,symbols})).toEqual(expected({source,mod,symbols},1));
});
