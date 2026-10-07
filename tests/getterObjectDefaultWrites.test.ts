import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
function errors(line:string,type='Collection',extra='') {
 const body=`Public Property Get Child() As ${type}\nSet Child = ${type==='Collection'?'New Collection':type==='Range'?'ThisWorkbook.Worksheets(1).Range("A1")':type==='Worksheet'?'ThisWorkbook.Worksheets(1)':'New '+type}\nEnd Property`;
 return analyzeProjectModule('Option Explicit\nSub T(ByVal item As Widget)\n'+line+'\nEnd Sub',[{moduleName:'Widget',moduleKind:'class',source:body},...(extra?[{moduleName:'Leaf',moduleKind:'class' as const,source:extra}]:[])],'Caller').filter(d=>d.severity==='error');
}
it.each(['item.Child = 20','item.Child() = 20','With item\n.Child = 20\nEnd With','If True Then item.Child = 20'])('requires the returned Collection index: %s',line=>{
 expect(errors(line)).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('keeps an indexed Collection item dynamic and a Range result writable',()=>{
 expect(errors('item.Child(1) = 20')).toEqual([]);
 expect(errors('item.Child = 20','Range')).toEqual([]);
});
it.each(['Worksheet','Widget'])('reports a returned %s without a default',type=>{
 expect(errors('item.Child = 20',type)).toEqual([expect.objectContaining({code:'runtime-member-not-found'})]);
});
it.each(['Long','Variant'])('rejects a readonly source default returning %s',type=>{
 const leaf=`Public Property Get Item() As ${type}\nAttribute Item.VB_UserMemId = 0\nItem = 1\nEnd Property`;
 expect(errors('item.Child = 20','Leaf',leaf)).toEqual([expect.objectContaining({code:'invalid-property-use'})]);
});
it('prioritizes a readonly source default over its missing index',()=>{
 const leaf='Public Property Get Item(ByVal index As Long) As Variant\nAttribute Item.VB_UserMemId = 0\nItem = index\nEnd Property';
 expect(errors('item.Child = 20','Leaf',leaf)).toEqual([expect.objectContaining({code:'invalid-property-use'})]);
});
it('checks a bare standard-module getter returning a Collection',()=>{
 const source='Property Get Child() As Collection\nSet Child = New Collection\nEnd Property\nSub T()\nChild = 20\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
});

it('checks the returned default after consuming the getter argument',()=>{
 const source='Option Explicit\nSub T(ByVal item As Widget)\nitem.Child("key") = 20\nEnd Sub';
 const getter='Public Property Get Child(ByVal key As String) As Collection\nSet Child = New Collection\nEnd Property';
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:getter}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('does not duplicate invalid getter argument counts',()=>{
 const source='Option Explicit\nSub T(ByVal item As Widget)\nitem.Child(1,2) = 20\nEnd Sub';
 const getter='Public Property Get Child(ByVal key As String) As Collection\nSet Child = New Collection\nEnd Property';
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:getter}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('does not replace a valid indexed source-default object write with a readonly error',()=>{
 const leaf='Public Property Get Item(ByVal index As Long) As Variant\nAttribute Item.VB_UserMemId = 0\nSet Item = ThisWorkbook.Worksheets(1).Range("A1")\nEnd Property';
 expect(errors('item.Child(1) = 20','Leaf',leaf)).toEqual([]);
});

it('allows a writable default with optional indices',()=>{
 const leaf='Private stored As Long\nPublic Property Get Item(Optional ByVal index As Long = 1) As Long\nAttribute Item.VB_UserMemId = 0\nItem = stored\nEnd Property\nPublic Property Let Item(Optional ByVal index As Long = 1, ByVal value As Long)\nAttribute Item.VB_UserMemId = 0\nstored = value\nEnd Property';
 expect(errors('item.Child = 20','Leaf',leaf)).toEqual([]);
});
