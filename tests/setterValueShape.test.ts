import {describe,expect,it} from 'vitest';
import {analyzeProjectModule} from './diagnostics/helpers';
import {resolveAssignmentValueCompletion} from '../src/analyzer/completion/assignmentValueCompletion';
import {buildLiveVbaProjectIndex,projectEditorSymbolContextForModule} from '../src/vbaProjectAnalysis';
function resolve(source:string,marker:string,modules:{moduleName:string;moduleKind:'standard'|'class';source:string}[]=[]) {
 const context=projectEditorSymbolContextForModule(buildLiveVbaProjectIndex([{moduleName:'Caller',moduleKind:'standard',source},...modules]),'Caller');
 return resolveAssignmentValueCompletion(source,source.indexOf(marker)+marker.length,{moduleName:'Caller',projectClassMembers:context.analysisOptions.projectClassMembers,projectSymbols:context.externalProjectSymbols});
}
describe('setter value shape and default typing',()=>{
 it('does not suggest Boolean scalars for a bare array setter',()=>{
  const source='Property Let Flags(ByRef value() As Boolean)\nEnd Property\nSub T()\nFlags = ';
  expect(resolve(source,'Flags = ')).toBeUndefined();
 });
 it('does not suggest scalars for an indexed member whose value parameter is an array',()=>{
  const source='Sub T(ByVal item As Widget)\nitem.Flags(1) = ';
  expect(resolve(source,'item.Flags(1) = ',[{moduleName:'Widget',moduleKind:'class',source:'Property Let Flags(ByVal index As Long, ByRef value() As Boolean)\nEnd Property'}])).toBeUndefined();
 });
 it('uses DefBool for a bare local setter parameter',()=>{
  const source='DefBool V\nProperty Let State(ByVal value)\nEnd Property\nSub T()\nState = ';
  expect(resolve(source,'State = ')?.enumName).toBe('Boolean');
 });
 it('uses the owning class DefBool for a dotted setter parameter',()=>{
  const source='Sub T(ByVal item As Widget)\nitem.State = ';
  expect(resolve(source,'item.State = ',[{moduleName:'Widget',moduleKind:'class',source:'DefBool V\nProperty Let State(ByVal value)\nEnd Property'}])?.enumName).toBe('Boolean');
 });
 it('uses the owning module DefBool for an exported bare setter',()=>{
  const source='Sub T()\nState = ';
  expect(resolve(source,'State = ',[{moduleName:'Library',moduleKind:'standard',source:'DefBool V\nPublic Property Let State(ByVal value)\nEnd Property'}])?.enumName).toBe('Boolean');
 });
});

it('retains a scalar read-only error for an implicitly typed class getter',()=>{
 const source='Option Explicit\nSub T(ByVal item As Widget)\nitem.State = True\nEnd Sub';
 expect(analyzeProjectModule(source,[{moduleName:'Widget',moduleKind:'class',source:'DefBool S\nProperty Get State()\nState = True\nEnd Property'}],'Caller')).toContainEqual(expect.objectContaining({code:'readonly-member-assignment'}));
});

it('reuses exported setter lookup metadata for repeated requests',()=>{
 const source='Sub T()\nState = ';
 const project=buildLiveVbaProjectIndex([{moduleName:'Caller',moduleKind:'standard',source},{moduleName:'Library',moduleKind:'standard',source:'DefBool V\nPublic Property Let State(ByVal value)\nEnd Property'}]);
 const ctx=projectEditorSymbolContextForModule(project,'Caller');
 const surfaces=ctx.analysisOptions.projectClassMembers!; let reads=0;
 const surface=surfaces.find(s=>s.moduleName==='Library'&&s.kind==='standardModule')!;
 const member=surface.members.find(m=>m.name==='State')!; const name=member.name;
 Object.defineProperty(member,'name',{get(){reads++;return name;}});
 const options={moduleName:'Caller',projectClassMembers:surfaces,projectSymbols:ctx.externalProjectSymbols};
 expect(resolveAssignmentValueCompletion(source,source.length,options)?.enumName).toBe('Boolean');
 reads=0;
 for(let i=0;i<20;i++) expect(resolveAssignmentValueCompletion(source,source.length,options)?.enumName).toBe('Boolean');
 expect(reads).toBe(0);
});
