import {expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
import {buildLiveVbaProjectIndex,projectEditorSymbolContextForModule} from '../src/vbaProjectAnalysis';
import {memberExpressionCalls} from '../src/analyzer/diagnostics/typeInference';
function declaration(kind='Property Get',passing='ByRef',implicit=false) {
 return (implicit?'DefLng I\n':'')+`Public ${kind} At(${passing} index${implicit?'':' As Long'}) As Worksheet\nSet At = ThisWorkbook.Worksheets(1)\nEnd ${kind==='Function'?'Function':'Property'}`;
}
function errors(line:string,decl=declaration()) {
 const source='Option Explicit\nSub T(ByVal item As Widget)\nDim i As Integer\n'+line+'\nEnd Sub';
 return analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:decl}],'Caller').filter(d=>d.severity==='error');
}
it.each(['Property Get','Function'])('checks source %s receiver arguments',kind=>{
 for(const line of ['item.At(i).EnableCalculation = True','With item.At(i)\n.EnableCalculation = True\nEnd With','If True Then item.At(i).EnableCalculation = True']) {
  expect(errors(line,declaration(kind)),line).toEqual([expect.objectContaining({code:'byref-argument-type-mismatch'})]);
 }
 expect(errors('item.At((i)).EnableCalculation = True',declaration(kind))).toEqual([]);
 expect(errors('item.At(i).EnableCalculation = True',declaration(kind,'ByVal'))).toEqual([]);
});
it.each(['Property Get','Function'])('uses owning default types for %s',kind=>{
 expect(errors('item.At(i).EnableCalculation = True',declaration(kind,'ByRef',true))).toEqual([expect.objectContaining({code:'byref-argument-type-mismatch'})]);
});
it('checks a parenless source method with its passing mode',()=>{
 expect(errors('item.Take i','Public Sub Take(ByRef index As Long)\nEnd Sub')).toEqual([expect.objectContaining({code:'byref-argument-type-mismatch'})]);
 expect(errors('item.Take(i)','Public Sub Take(ByRef index As Long)\nEnd Sub')).toEqual([]);
 expect(errors('item.Take (i)','Public Sub Take(ByRef index As Long)\nEnd Sub')).toEqual([]);
});
it('reuses immutable source parameter contracts across expression calls',()=>{
 const source='Sub T(ByVal item As Widget)\nitem.At(1).EnableCalculation = True\nEnd Sub';
 const ctx=projectEditorSymbolContextForModule(buildLiveVbaProjectIndex([{moduleName:'Caller',moduleKind:'standard',source},{moduleName:'Widget',moduleKind:'class',source:declaration()}]),'Caller');
 const options={moduleName:'Caller',projectClassMembers:ctx.analysisOptions.projectClassMembers};
 const span={start:source.indexOf('item.At'),end:source.indexOf('\nEnd Sub')};
 const first=memberExpressionCalls(source,span,options)[0].signature;
 for(let i=0;i<20;i++) { expect(memberExpressionCalls(source,span,options)[0].signature.params).toBe(first.params); }
});

it('checks array shape in a getter receiver and keeps a typed scalar valid',()=>{
 const source='Option Explicit\nSub T(ByVal item As Widget)\nDim values(1) As Long\nitem.At(values).EnableCalculation = True\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:declaration('Property Get','ByVal')}],'Caller').filter(d=>d.severity==='error')).toEqual([expect.objectContaining({code:'argument-shape-mismatch'})]);
});

it('does not treat a source member named Left as a VBA call',()=>{
 expect(errors('Call item.Left("abc",1)','Public Sub Left()\nEnd Sub')).toEqual([expect.objectContaining({code:'argument-count'})]);
});
