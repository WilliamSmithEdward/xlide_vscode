import {expect,it,vi,afterEach} from 'vitest';
import * as inference from '../src/analyzer/diagnostics/typeInference';
afterEach(()=>vi.restoreAllMocks());
import {analyzeProjectModule} from './diagnostics/helpers';
const fn='Function MakeFlags() As Boolean()\nDim values(1) As Boolean\nMakeFlags = values\nEnd Function';
const setter='Public Property Let State(ByVal value As Boolean)\nEnd Property';
function errors(line:string,bare=false) {
 const source='Option Explicit\n'+fn+'\n'+(bare?setter+'\n':'')+'Sub T(ByVal item As Widget)\nDim values(1) As Boolean\n'+line+'\nEnd Sub';
 return analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:setter},{moduleName:'Library',source:fn}],'Caller').filter(d=>d.severity==='error');
}
it.each([true,false])('rejects typed arrays for scalar setters, bare %s',bare=>{
 for(const rhs of ['values','[values]','(values)','MakeFlags()','(MakeFlags())','Library.MakeFlags()']) {
  expect(errors((bare?'State':'item.State')+' = '+rhs,bare),rhs).toEqual([expect.objectContaining({code:'array-assignment-to-scalar'})]);
 }
 expect(errors((bare?'State':'item.State')+' = values(0)',bare)).toEqual([]);
});
it('checks With and If assignments',()=>{
 expect(errors('With item\n.State = MakeFlags()\nEnd With')).toEqual([expect.objectContaining({code:'array-assignment-to-scalar'})]);
 expect(errors('If True Then item.State = MakeFlags()')).toEqual([expect.objectContaining({code:'array-assignment-to-scalar'})]);
});

it('checks a qualified whole array but preserves its element',()=>{
 const library='Public values(1) As Boolean';
 for(const rhs of ['Library.values','(Library.[values])']) {
  const source='Sub T(ByVal item As Widget)\nitem.State = '+rhs+'\nEnd Sub';
  expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:setter},{moduleName:'Library',source:library}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'array-assignment-to-scalar'})]);
 }
 const source='Sub T(ByVal item As Widget)\nitem.State = Library.values(0)\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:setter},{moduleName:'Library',source:library}],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});

it('infers a scalar function result once for shape and coercion checks',()=>{
 const source='Function MakeFlag() As Boolean\nMakeFlag = True\nEnd Function\nSub T(ByVal item As Widget)\nitem.State = MakeFlag()\nEnd Sub';
 const infer=vi.spyOn(inference,'inferArgumentType');
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:setter}],'Caller').filter(d=>d.severity==='error')).toEqual([]);
 expect(infer.mock.calls.filter(call=>call[0].map(token=>token.rawText).join('')==='MakeFlag()')).toHaveLength(1);
});

it('keeps the host Boolean property failure at runtime',()=>{
 const source=fn+'\nSub T(ByVal sheet As Worksheet)\nDim values(1) As Boolean\nsheet.EnableCalculation = MakeFlags()\nsheet.EnableCalculation = values\nEnd Sub';
 const found=analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error');
 expect(found).toHaveLength(2);
 expect(found.every(d=>d.code==='assignment-type-mismatch' && d.message.includes("Run-time error '13'"))).toBe(true);
});
it('checks scalar variables without mistaking an element for an array',()=>{
 const source=fn+'\nSub T()\nDim flag As Boolean\nDim values(1) As Boolean\nflag = MakeFlags()\nflag = values\nflag = values(0)\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'array-assignment-to-scalar'}),expect.objectContaining({code:'array-assignment-to-scalar'})]);
});

it('preserves whole Byte-array conversion to a String setter',()=>{
 const source='Function Bytes() As Byte()\nDim data(1) As Byte\nBytes = data\nEnd Function\nProperty Let Text(ByVal value As String)\nEnd Property\nSub T()\nDim data(1) As Byte\nText = data\nText = Bytes()\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
