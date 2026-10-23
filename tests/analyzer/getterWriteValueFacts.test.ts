import {describe,expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
const literal='Public Property Get Height() As Variant\nHeight = 20\nEnd Property';
function found(statement:string,getter=literal) {
 const source=`Option Explicit\nSub T(ByVal item As Widget)\n${statement}\nEnd Sub`;
 return analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:getter}],'Caller').filter(d=>d.severity==='error');
}
describe('Let through getter-only Variant results',()=>{
 it.each(['item.Height = 30','With item\n.Height = 30\nEnd With','If True Then item.Height = 30'])('reports a known scalar result: %s',statement=>{
  const result=found(statement);
  expect(result).toEqual([expect.objectContaining({code:'variant-value-misuse'})]);
  expect(result[0].message).toContain("Run-time error '424'");
 });
 it('checks the result after an indexed receiver chain',()=>{
  const getter=literal+'\nPublic Function At(ByVal index As Long) As Widget\nSet At = Me\nEnd Function';
  expect(found('item.At(1).Height = 30',getter)).toEqual([expect.objectContaining({code:'variant-value-misuse'})]);
 });
 it.each(['Height = 20','Height = index'])('checks indexed scalar getters: %s',body=>{
  expect(found('item.Height(1) = 30',`Public Property Get Height(ByVal index As Long) As Variant\n${body}\nEnd Property`)).toEqual([expect.objectContaining({code:'variant-value-misuse'})]);
 });
 it('checks a scalar field return',()=>{
  expect(found('item.Height = 30','Private stored As Long\n'+literal.replace('Height = 20','Height = stored'))).toEqual([expect.objectContaining({code:'variant-value-misuse'})]);
 });
 it('keeps an object-valued result writable through its default property',()=>{
  expect(found('item.Height = 30',literal.replace('Height = 20','Set Height = ThisWorkbook.Worksheets(1).Range("A1")'))).toEqual([]);
 });
 it('does not infer a scalar from a Variant parameter',()=>{
  expect(found('item.Height(ThisWorkbook.Worksheets(1).Range("A1")) = 30','Public Property Get Height(ByVal value As Variant) As Variant\nSet Height = value\nEnd Property')).toEqual([]);
 });
 it('does not apply getter facts to a real setter',()=>{
  const getter=literal+'\nPublic Property Let Height(ByVal value As Variant)\nEnd Property';
  expect(found('item.Height = 30',getter)).toEqual([]);
 });
 it('retains compile-time readonly errors on array getters',()=>{
  const getter='Public Property Get Height() As Boolean()\nDim result(1) As Boolean\nHeight = result\nEnd Property';
  expect(found('item.Height(0) = True',getter)).toEqual([expect.objectContaining({code:'readonly-member-assignment'})]);
 });
});

it('reads an indexed scalar getter without treating its argument as result indexing',()=>{
 const getter='Public Property Get Height(ByVal index As Long) As Variant\nHeight = index\nEnd Property';
 expect(found('Debug.Print item.Height(1)',getter)).toEqual([]);
 expect(found('item.Height(1).Value = 30',getter)).toEqual([expect.objectContaining({code:'variant-value-misuse'})]);
});

it.each([false,true])('checks a %s indexed standard-module getter',indexed=>{
 const source=`Option Explicit\nProperty Get Height(${indexed?'ByVal index As Long':''}) As Variant\nHeight = 20\nEnd Property\nSub T()\nHeight${indexed?'(1)':''} = 30\nEnd Sub`;
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'variant-value-misuse'})]);
});
it('checks an exported getter in bare and qualified form',()=>{
 for(const name of ['Height','Library.Height']) {
  const source=`Option Explicit\nSub T()\n${name} = 30\nEnd Sub`;
  expect(analyzeProjectModule(source,[{moduleName:'Library',source:literal}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'variant-value-misuse'})]);
 }
});
it.each(['Variant','Widget'])('keeps %s array getters read-only',type=>{
 const getter=`Public Property Get Height() As ${type}()\nDim values(1) As ${type}\nHeight = values\nEnd Property`;
 expect(found('item.Height = 30',getter)).toEqual([expect.objectContaining({code:'readonly-member-assignment'})]);
});
it('does not apply bare getter rules to a local variable shadow',()=>{
 const source='Option Explicit\n'+literal+'\nSub T()\nDim Height As Variant\nHeight = 30\nDebug.Print Height\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
it('preserves optional argument defaults containing a comma',()=>{
 const getter='Public Property Get Height(Optional ByVal index As String = "1,2") As Variant\nSet Height = ThisWorkbook.Worksheets(1).Range("A1")\nEnd Property';
 expect(found('item.Height = 30',getter)).toEqual([]);
 expect(found('item.Height = 30',literal.replace('Height()', 'Height(ByVal index As Long)'))).toEqual([expect.objectContaining({code:'argument-count'})]);
});

it('leaves a bare missing getter argument to one compile diagnostic',()=>{
 const source='Option Explicit\nProperty Get Height(ByVal index As Long) As Variant\nHeight = 20\nEnd Property\nSub T()\nHeight = 30\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
