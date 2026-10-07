import { describe, expect, it, vi } from 'vitest';
const work=vi.hoisted(()=>({builds:new Map<string,number>(),statementReads:0}));
vi.mock('../src/analyzer/symbols/buildModuleSymbols',async original=>{
 const actual=await original<typeof import('../src/analyzer/symbols/buildModuleSymbols')>();
 return {...actual,buildModuleSymbols:(...args:Parameters<typeof actual.buildModuleSymbols>)=>{work.builds.set(args[0],(work.builds.get(args[0])??0)+1);return actual.buildModuleSymbols(...args);}};
});
vi.mock('../src/analyzer/parser/parseModule',async original=>{
 const actual=await original<typeof import('../src/analyzer/parser/parseModule')>();
 function freeze<T>(value:T):T {if(value&&typeof value==='object'&&!Object.isFrozen(value)){for(const child of Object.values(value))freeze(child);Object.freeze(value);}return value;}
 return {...actual,parseModule:(source:string)=>freeze(actual.parseModule(source))};
});
vi.mock('../src/analyzer/lexer/tokenHelpers',async original=>{
 const actual=await original<typeof import('../src/analyzer/lexer/tokenHelpers')>();
 return {...actual,statementTokensCached:(...args:Parameters<typeof actual.statementTokensCached>)=>{work.statementReads++;return actual.statementTokensCached(...args);}};
});
import { introduceParameter } from '../src/analyzer/refactor/introduceParameter';
import { applyVbaTextEdits } from '../src/analyzer/refactor/refactorTypes';
for(const count of [1,100,1000]) {
 for(const layout of ['bare','Me','WithMe']) {
  it('builds each binding module once with frozen ASTs at '+count+' calls, layout='+layout,()=>{
   const call='If depth > 0 Then '+(layout==='Me'?'Me.':layout==='WithMe'?'.':'')+'Report depth - 1';
   const statements=[...(layout==='WithMe'?['With Me']:[]),...Array(count).fill(call),...(layout==='WithMe'?['End With']:[])];
   const source=['Public Sub Report(ByVal depth As Long)','Dim limit As Long','limit = 3',...statements,'End Sub',''].join('\n');
   work.builds.clear();work.statementReads=0;
   const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',moduleKinds:{Module1:'class',Other:'class'},otherModuleSources:{Other:'Public Sub Another()\nEnd Sub\n'}});
   if(!result.ok)throw new Error(result.reason);
   expect(work.builds.get('Module1')??0).toBeLessThanOrEqual(1);
   expect(work.builds.get('Other')??0).toBeLessThanOrEqual(layout==='bare'?0:1);
   expect(work.statementReads).toBeLessThanOrEqual(count * 3 + 10);
   expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub Report(ByVal depth As Long, ByVal limit As Long)',...statements.map(statement=>statement===call?call+', 3':statement),'End Sub',''].join('\n'));
  });
 }
}

for(const count of [1,100,1000]) for(const qualified of [false,true]) {
 it('shares the project between recursion and external calls at '+count+', qualified='+qualified,()=>{
  const call='If depth > 0 Then '+(qualified?'Module1.':'')+'Report depth - 1';
  const source=['Public Sub Report(ByVal depth As Long)','Dim limit As Long','limit = 3',...Array(count).fill(call),'End Sub',''].join('\n');
  const caller=['Sub Caller()',...Array(count).fill('Module1.Report 2'), 'Other.Report 2','End Sub',''].join('\n');
  const other='Public Sub Report(ByVal depth As Long)\nEnd Sub\n';
  work.builds.clear();work.statementReads=0;
  const result=introduceParameter({source,offset:source.indexOf('limit'),moduleName:'Module1',otherModuleSources:{Caller:caller,Other:other}});
  if(!result.ok)throw new Error(result.reason);
  for(const name of ['Module1','Caller','Other'])expect(work.builds.get(name)??0).toBeLessThanOrEqual(1);
  expect(applyVbaTextEdits(source,result.edits)).toBe(['Public Sub Report(ByVal depth As Long, ByVal limit As Long)',...Array(count).fill(call+', 3'),'End Sub',''].join('\n'));
  expect(result.otherModules).toHaveLength(1);
  expect(result.otherModules![0].moduleName).toBe('Caller');
  expect(applyVbaTextEdits(caller,result.otherModules![0].edits)).toBe(['Sub Caller()',...Array(count).fill('Module1.Report 2, 3'),'Other.Report 2','End Sub',''].join('\n'));
 });
}
