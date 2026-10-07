import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
const leaf={moduleName:'Leaf',moduleKind:'class' as const,source:'Option Explicit'};
function diagnostics(type:string,line='Factory() = 20',qualified=false){
 const factory='Function Factory() As '+type+'\nEnd Function';
 const source=(qualified?'':factory+'\n')+'Sub T()\n'+line+'\nEnd Sub';
 return analyzeProjectModule(source,[leaf,...(qualified?[{moduleName:'Library',source:factory}]:[])],'Caller').filter(d=>d.severity==='error');
}
it.each(['Worksheet','Leaf'])('reports no default on a %s function result',type=>{
 expect(diagnostics(type)).toEqual([expect.objectContaining({code:'runtime-member-not-found'})]);
});
it('requires the default Collection index',()=>{
 expect(diagnostics('Collection')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('keeps a Range result writable and an indexed Collection result dynamic',()=>{
 expect(diagnostics('Range')).toEqual([]);
 expect(diagnostics('Collection','Factory(1) = 20')).toEqual([]);
});

it.each(['Worksheet','Leaf','Collection'])('checks a qualified %s factory result',type=>{
 expect(diagnostics(type,'Library.Factory() = 20',true)).toEqual([expect.objectContaining({code:type==='Collection'?'argument-count':'runtime-member-not-found'})]);
});
it.each(['item.Factory() = 20','With item\n.Factory() = 20\nEnd With','If True Then item.Factory() = 20'])('checks member factory defaults: %s',line=>{
 const source='Sub T(ByVal item As Widget)\n'+line+'\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:'Function Factory() As Collection\nEnd Function'}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
it('preserves the factory return assignment and a local shadow',()=>{
 const source='Function Factory() As Range\nSet Factory = ThisWorkbook.Worksheets(1).Range("A1")\nEnd Function\nSub T()\nDim Factory As Long\nFactory = 20\nEnd Sub';
 expect(analyzeProjectModule(source,[],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
