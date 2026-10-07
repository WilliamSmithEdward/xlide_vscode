import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
function diagnostics(rhs:string,type='Long') {
 const source='Sub T()\nDim values() As Boolean\nvalues = '+rhs+'\nEnd Sub';
 return analyzeProjectModule(source,[{moduleName:'Library',source:'Public Function MakeFlags() As '+type+'()\nDim result(1) As '+type+'\nMakeFlags = result\nEnd Function'}],'Caller').filter(d=>d.severity==='error');
}
it.each(['Library.MakeFlags()','(Library.MakeFlags())','Library.[MakeFlags]()'])('rejects incompatible qualified array results: %s',rhs=>{
 expect(diagnostics(rhs)).toEqual([expect.objectContaining({code:'array-target-assignment'})]);
 expect(diagnostics(rhs,'Boolean')).toEqual([]);
});

it.each([['XlHAlign','Excel.XlHAlign'],['Excel.XlHAlign','XlHAlign'],['XlHAlign','Long'],['Direction','Long'],['Long','Direction']])('uses enum storage identity: %s / %s',(target,returned)=>{
 const library=(target==='Direction'||returned==='Direction'?'Public Enum Direction\nNorth = 1\nEnd Enum\n':'')+'Function Factory() As '+returned+'()\nDim data(1) As '+returned+'\nFactory = data\nEnd Function';
 const source='Sub T()\nDim values() As '+target+'\nvalues = Library.Factory()\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Library',source:library}],'Caller').filter(d=>d.severity==='error')).toEqual([]);
});
it.each([['Object','Widget'],['Widget','Object'],['Widget','Widget']])('checks class array identity: %s / %s',(target,returned)=>{
 const library='Function Factory() As '+returned+'()\nDim data(1) As '+returned+'\nFactory = data\nEnd Function';
 const modules=[{moduleName:'Library',source:library},{moduleName:'Widget',moduleKind:'class' as const,source:'Option Explicit'}];
 const source='Sub T()\nDim values() As '+target+'\nvalues = Library.Factory()\nEnd Sub';
 const errors=analyzeProjectModule(source,modules,'Caller').filter(d=>d.severity==='error');
 expect(errors).toEqual(target===returned?[]:[expect.objectContaining({code:'array-target-assignment'})]);
 expect(analyzeProjectModule(library,modules,'Library').filter(d=>d.severity==='error')).toEqual([]);
});

it.each(['Long','Boolean'])('preserves qualified source array field shape: %s',type=>{
 const library='Public values(1) As '+type;
 const source='Sub T()\nDim copy() As Boolean\ncopy = Library.[values]\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Library',source:library}],'Caller').filter(d=>d.severity==='error')).toEqual(type==='Boolean'?[]:[expect.objectContaining({code:'array-target-assignment'})]);
});

it.each([['Object','Widget',false],['Widget','Widget',true],['Direction','Long',false],['XlHAlign','Excel.XlHAlign',true]])('keeps setter ByRef identity distinct from array-copy storage: %s / %s',(target,returned,valid)=>{
 const library=(target==='Direction'?'Public Enum Direction\nNorth = 1\nEnd Enum\n':'')+'Function Factory() As '+returned+'()\nDim data(1) As '+returned+'\nFactory = data\nEnd Function';
 const source='Sub T(ByVal item As Receiver)\nitem.Flags = Library.Factory()\nEnd Sub';
 const modules=[{moduleName:'Library',source:library},{moduleName:'Widget',moduleKind:'class' as const,source:'Option Explicit'},{moduleName:'Receiver',moduleKind:'class' as const,source:'Property Let Flags(ByRef value() As '+target+')\nEnd Property'}];
 expect(analyzeProjectModule(source,modules,'Caller').filter(d=>d.severity==='error')).toEqual(valid?[]:[expect.objectContaining({code:'argument-shape-mismatch'})]);
});
