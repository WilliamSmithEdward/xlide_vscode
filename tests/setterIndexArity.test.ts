import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
const setter='Public Property Let State(ByVal index As Long, ByVal value As Boolean)\nEnd Property';
const getter='Public Property Get State(ByVal index As Long) As Boolean\nState = True\nEnd Property\n';
function errors(line:string,body=setter) {
 return analyzeProjectModule('Option Explicit\nSub T(ByVal item As Widget)\n'+line+'\nEnd Sub',[{moduleName:'Widget',moduleKind:'class',source:body}],'Caller').filter(d=>d.severity==='error');
}
it.each([false,true])('checks index counts with getter %s',hasGetter=>{
 for(const line of ['item.State = True','item.State() = True','item.State(1,2) = True','With item\n.State = True\nEnd With','If True Then item.State() = True','Let item.State = True','10 item.State() = True']) {
  expect(errors(line,(hasGetter?getter:'')+setter),line).toEqual([expect.objectContaining({code:'argument-count'})]);
 }
 expect(errors('item.State(1) = True',(hasGetter?getter:'')+setter)).toEqual([]);
});
it('prioritizes invalid index counts over setter value runtime errors',()=>{
 expect(errors('item.State() = "nonsense"')).toEqual([expect.objectContaining({code:'argument-count'})]);
 expect(errors('item.State(1) = "nonsense"')).toEqual([expect.objectContaining({code:'assignment-type-mismatch'})]);
});
it('checks bare and module-qualified standard setters',()=>{
 for(const name of ['State','Library.State']) {
  const source='Option Explicit\nSub T()\n'+name+' = True\nEnd Sub';
  expect(analyzeProjectModule(source,[{moduleName:'Library',source:setter}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
 }
 const source=setter+'\nSub T()\nState() = True\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('checks Set accessor indices and excludes the assigned object from the count',()=>{
 const body='Public Property Set State(ByVal index As Long, ByVal value As Worksheet)\nEnd Property';
 expect(errors('Set item.State = ThisWorkbook.Worksheets(1)',body)).toEqual([expect.objectContaining({code:'argument-count'})]);
 expect(errors('Set item.State(1) = ThisWorkbook.Worksheets(1)',body)).toEqual([]);
});
it('accepts omitted optional indices and counts quoted defaults correctly',()=>{
 const body='Public Property Let State(Optional ByVal index As String = "1,2", ByVal value As Boolean)\nEnd Property';
 expect(errors('item.State = True',body)).toEqual([]);
 expect(errors('item.State() = True',body)).toEqual([]);
});

it.each(['item.State = True','item.State() = True'])('inserts a missing index into valid assignment syntax: %s',line=>{
 const finding=errors(line)[0];
 const edit=finding.data!.missingRequiredArgumentPlaceholder!.edit;
 const source='Option Explicit\nSub T(ByVal item As Widget)\n'+line+'\nEnd Sub';
 const corrected=source.slice(0,edit.span.start)+edit.newText.replace('TODO_index','1')+source.slice(edit.span.end);
 expect(analyzeProjectModule(corrected,[{moduleName:'Widget',moduleKind:'class',source:setter}],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
it('does not fall back to a VBA function for a source setter',()=>{
 const source='Property Let Left(ByVal value As Long)\nEnd Property\nSub T()\nLeft(1,2) = 3\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
