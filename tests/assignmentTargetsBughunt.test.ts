import { afterEach, describe, expect, it, vi } from 'vitest';
import { analyzeProjectModule } from './diagnostics/helpers';
import { resolveAssignmentValueCompletion } from '../src/analyzer/completion/assignmentValueCompletion';
function values(statement:string,prelude='') {
 const source=`Option Explicit\n${prelude}\nSub T()\n${statement}\nEnd Sub`;
 return resolveAssignmentValueCompletion(source,source.indexOf(statement)+statement.length);
}
describe('assignment target bug hunt',()=>{
 it.each(['IsNumeric("abc") = ', 'IsNumeric = ', 'Ready() = ', 'Ready = '])('does not offer values for a scalar function target %s',statement=>{
  expect(values(statement,'Function Ready() As Boolean\nEnd Function')).toBeUndefined();
 });
 it.each(['enabled(1) = ', 'enabled(1, 2) = '])('completes Boolean array elements: %s',statement=>{
  expect(values(statement,'Dim enabled(1,2) As Boolean')?.enumName).toBe('Boolean');
 });
 it('completes an enum array element',()=>{
  expect(values('facing(1) = ', 'Enum Direction\nNorth = 1\nEnd Enum\nDim facing(1) As Direction')?.constants.map(c=>c.name)).toEqual(['North']);
 });
 it('does not suggest scalar constants for a whole array assignment',()=>{
  expect(values('enabled = ', 'Dim enabled(1) As Boolean')).toBeUndefined();
 });
 it('respects a local array shadowing a runtime function',()=>{
  expect(values('IsNumeric(1) = ', 'Dim IsNumeric(1) As Boolean')?.enumName).toBe('Boolean');
 });
 it('rejects indexing a scalar variable',()=>{
  expect(values('enabled(1) = ', 'Dim enabled As Boolean')).toBeUndefined();
 });
 it('keeps assignment to the current function return variable',()=>{
  const source='Function Ready() As Boolean\nReady = ';
  expect(resolveAssignmentValueCompletion(source,source.length)?.enumName).toBe('Boolean');
 });
 it('uses the value type of an indexed setter-only property',()=>{
  const source='Sub T()\nDim item As Widget\nitem.State(1) = ';
  expect(resolveAssignmentValueCompletion(source,source.length,{projectClassMembers:[{name:'Widget',moduleName:'Widget',kind:'class',members:[{name:'State',kind:'property',writeType:'Boolean',writable:true,letAccessor:true}]}]})?.enumName).toBe('Boolean');
 });
 it('uses a bare source Property Let value type',()=>{
  const source='Property Let State(ByVal index As Long, ByVal value As Boolean)\nEnd Property\nSub T()\nState(1) = ';
  expect(resolveAssignmentValueCompletion(source,source.length)?.enumName).toBe('Boolean');
 });
 it('does not offer a getter-only bare source property',()=>{
  expect(values('State = ', 'Property Get State() As Boolean\nEnd Property')).toBeUndefined();
 });
});

const widget='Private held As Boolean\nPublic Property Get State(ByVal index As Long) As Boolean\nState = held\nEnd Property\nPublic Property Let State(ByVal index As Long, ByVal value As Boolean)\nheld = value\nEnd Property\n';
function projectDiagnostics(statement:string,cls=widget) {
 const source='Option Explicit\nSub T(ByVal item As Widget)\n'+statement+'\nEnd Sub';
 return analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:cls}],'Caller');
}
describe('indexed property assignment diagnostics',()=>{
 it.each(['item.State(1) = "nonsense"', 'With item\n.State(1) = "nonsense"\nEnd With', 'If True Then item.State(1) = "nonsense"'])('flags the actual setter value: %s',statement=>{
  expect(projectDiagnostics(statement)).toEqual([expect.objectContaining({code:'assignment-type-mismatch',severity:'error'})]);
 });
 it.each(['item.State(1) = True', 'item.State(1) = "True"', 'item.State(1) = 1'])('allows valid conversions: %s',statement=>{
  expect(projectDiagnostics(statement)).toEqual([]);
 });
 it('flags the indexed getter-only property once',()=>{
  const cls='Public Property Get State(ByVal index As Long) As Boolean\nState = True\nEnd Property';
  expect(projectDiagnostics('item.State(1) = True',cls)).toEqual([expect.objectContaining({code:'readonly-member-assignment',severity:'error'})]);
 });
 it('flags an explicit empty call to a scalar getter-only property',()=>{
  const cls='Public Property Get State() As Boolean\nState = True\nEnd Property';
  expect(projectDiagnostics('item.State() = True',cls)).toEqual([expect.objectContaining({code:'readonly-member-assignment'})]);
 });
 it('does not duplicate the existing read-only String subscript finding',()=>{
  const cls='Public Property Get State() As String\nState = "text"\nEnd Property';
  expect(projectDiagnostics('item.State(1) = "a"',cls).filter(d=>d.code==='readonly-member-assignment')).toHaveLength(1);
 });
});


describe('getter-only object result negative controls',()=>{
 for (const type of ['Range','Variant']) {
  for (const indexed of [false,true]) {
   it(`allows a Let through a ${type} returned by ${indexed?'an indexed':'a bare'} getter`,()=>{
    const cls=`Public Property Get State(${indexed?'ByVal index As Long':''}) As ${type}\nSet State = ThisWorkbook.Worksheets(1).Range("A1")\nEnd Property`;
    expect(projectDiagnostics(`item.State${indexed?'(1)':''} = 20`,cls)).toEqual([]);
   });
  }
  it(`allows an unqualified standard-module getter returning ${type}`,()=>{
   const source=`Option Explicit\nProperty Get State() As ${type}\nSet State = ThisWorkbook.Worksheets(1).Range("A1")\nEnd Property\nSub T()\nState = 20\nEnd Sub`;
   expect(analyzeProjectModule(source,[],'Caller')).toEqual([]);
  });
 }
});


afterEach(()=>vi.restoreAllMocks());
it('reuses the bound source graph across repeated array-value requests',async()=>{
 const builder=await import('../src/analyzer/symbols/buildModuleSymbols');
 const source='Option Explicit\nSub CachedArrayProbe()\nDim flags(1) As Boolean\n'+'Debug.Print 1\n'.repeat(4000)+'flags(1) = True\nflags(1) = False\nEnd Sub';
 const first=source.indexOf('flags(1) =')+'flags(1) = '.length;
 const second=source.lastIndexOf('flags(1) =')+'flags(1) = '.length;
 expect(resolveAssignmentValueCompletion(source,first)?.enumName).toBe('Boolean');
 const build=vi.spyOn(builder,'buildModuleSymbols');
 for(let i=0;i<20;i++) expect(resolveAssignmentValueCompletion(source,i%2?first:second)?.enumName).toBe('Boolean');
 expect(build).not.toHaveBeenCalled();
});
it('keeps DefBool-typed arrays usable without an explicit As clause',()=>{
 expect(values('flags(1) = ', 'DefBool A-Z\nDim flags(1)')?.enumName).toBe('Boolean');
});

it('preserves the read-only error on a DefBool-typed standard-module getter',()=>{
 const source='Option Explicit\nDefBool S\nProperty Get State()\nState = True\nEnd Property\nSub T()\nState = True\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller')).toContainEqual(expect.objectContaining({code:'readonly-member-assignment'}));
});
