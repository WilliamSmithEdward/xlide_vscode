import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
const setter=(passing='ByVal',type='Long')=>`Public Property Let State(${passing} index${type?' As '+type:''}, ByVal value As Boolean)\nEnd Property`;
function errors(line:string,decl=setter(),prelude='Dim i As Integer',get=false) {
 const body=(get?'Public Property Get State(ByVal index As Long) As Boolean\nState = True\nEnd Property\n':'')+decl;
 return analyzeProjectModule('Option Explicit\nSub T(ByVal item As Widget)\n'+prelude+'\n'+line+'\nEnd Sub',[{moduleName:'Widget',moduleKind:'class',source:body}],'Caller').filter(d=>d.severity==='error');
}
it.each([false,true])('checks scalar setter indices with getter %s',get=>{
 for(const line of ['item.State("bad") = True','item.State(index:="bad") = True','With item\n.State("bad") = True\nEnd With','If True Then item.State("bad") = True']) {
  expect(errors(line,setter(),undefined,get)).toEqual([expect.objectContaining({code:'argument-type-mismatch'})]);
 }
 expect(errors('item.State(1) = True',setter(),undefined,get)).toEqual([]);
});
it('retains ByRef exactness and parentheses conversion',()=>{
 expect(errors('item.State(i) = True',setter('ByRef'))).toEqual([expect.objectContaining({code:'byref-argument-type-mismatch'})]);
 expect(errors('item.State((i)) = True',setter('ByRef'))).toEqual([]);
 expect(errors('item.State(i) = True')).toEqual([]);
});
it('prioritizes invalid counts over index and value mismatches',()=>{
 expect(errors('item.State("bad",2) = "bad"')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('checks array shape for setter indices',()=>{
 expect(errors('item.State(values) = True',setter(),'Dim values(1) As Long')).toEqual([expect.objectContaining({code:'argument-shape-mismatch'})]);
});
it('uses the owning class default parameter type',()=>{
 expect(errors('item.State(i) = True','DefLng I\n'+setter('ByRef',''))).toEqual([expect.objectContaining({code:'byref-argument-type-mismatch'})]);
});
it.each(['State','Library.State'])('checks exported indices for %s with owning default types',name=>{
 const source='Option Explicit\nSub T()\nDim i As Integer\n'+name+'(i) = True\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Library',source:'DefLng I\n'+setter('ByRef','')}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'byref-argument-type-mismatch'})]);
});
it('uses default types and passing flags for a same-module bare setter',()=>{
 const source='Option Explicit\nDefLng I\n'+setter('ByRef','')+'\nSub T()\nDim i As Integer\nState(i) = True\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'byref-argument-type-mismatch'})]);
});

it('keeps named count errors separate from type errors',()=>{
 expect(errors('item.State(unknown:="bad") = True')).toEqual([expect.objectContaining({code:'argument-count'})]);
 expect(errors('item.State(index:=1,index:="bad") = True')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('handles omitted required and optional middle indices',()=>{
 const required='Public Property Let State(ByVal first As Long, ByVal second As Long, ByVal third As Long, ByVal value As Boolean)\nEnd Property';
 expect(errors('item.State(1,,3) = True',required)).toEqual([expect.objectContaining({code:'argument-count'})]);
 const optional='Public Property Let State(ByVal first As Long, Optional ByVal second As Long = 2, Optional ByVal third As Long = 3, ByVal value As Boolean)\nEnd Property';
 expect(errors('item.State(1,,3) = True',optional)).toEqual([]);
});

it('accepts array indices of the declared shape and rejects scalar indices',()=>{
 const decl='Public Property Let State(ByRef index() As Long, ByVal value As Boolean)\nEnd Property';
 expect(errors('item.State(values) = True',decl,'Dim values(1) As Long')).toEqual([]);
 expect(errors('item.State(1) = True',decl)).toEqual([expect.objectContaining({code:'argument-shape-mismatch'})]);
});
it('checks scalar indices of a Set accessor',()=>{
 const decl='Public Property Set State(ByVal index As Long, ByVal value As Worksheet)\nEnd Property';
 expect(errors('Set item.State("bad") = ThisWorkbook.Worksheets(1)',decl)).toEqual([expect.objectContaining({code:'argument-type-mismatch'})]);
});
