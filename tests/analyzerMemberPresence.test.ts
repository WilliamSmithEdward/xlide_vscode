import { beforeEach, describe, expect, it, vi } from 'vitest';
const calls = vi.hoisted(() => ({ resolve: vi.fn() }));
vi.mock('../src/analyzer/completion/memberAccess', async importOriginal => {
 const real = await importOriginal<typeof import('../src/analyzer/completion/memberAccess')>();
 return { ...real, privateMemberOwnerAt: calls.resolve.mockImplementation(real.privateMemberOwnerAt) };
});
import { checkMemberNotFound } from '../src/analyzer/diagnostics/rules/undeclared';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { walkProcedureStatements } from '../src/analyzer/diagnostics/walker';
import { resolveMemberPresenceSurfaceAt, resolveExhaustiveMemberSurfaceAt, type MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
const type = (name='C', privateMembers: string[] = [], kind: VbaProjectClassMembers['kind']='document'): VbaProjectClassMembers => ({name,moduleName:name,kind,members:[],privateMembers});
function run(body:string,ctx:MemberCompletionContext){
 const source='Option Explicit\nSub Main()\nDim r As Range\nDim c As C\nDim result As Variant\n'+body+'\nEnd Sub\n';
 const push=vi.fn();walkProcedureStatements(parseModule(source),undefined,[checkMemberNotFound(source,ctx,push)],undefined,{source,takes:[true]});return push.mock.calls;
}
beforeEach(()=>{ calls.resolve.mockClear(); });
describe('known public member presence',()=>{
 it.each([[],[type()],[type('C',['Hidden'])],[type('C',['Hidden'],'class')],[type('C',['Hidden'],'userform')]].map(types=>({types})))('skips private-owner lookup for a known host member with metadata %j',({types})=>{
  expect(run(Array.from({length:1000},()=> 'result = r.Value').join('\n'),{projectClassMembers:types})).toEqual([]);
  expect(calls.resolve.mock.calls.length).toBe(0);
 });
 it.each(['class','document','userform','standardModule'] as const)('keeps private ownership for %s candidates',kind=>{
  const diagnostics=run('result = c.hIdDeN',{projectClassMembers:[type('C',['Hidden'],kind)]});
  expect(calls.resolve.mock.calls.length).toBe(kind==='class'?0:1);
  if(kind==='class') expect(diagnostics.map(args=>args[1])).toEqual(["Method or data member not found: 'C.hIdDeN'."]);
  else if(kind==='standardModule') expect(diagnostics).toEqual([]);
  else expect(diagnostics.map(args=>args[1])).toEqual([expect.stringContaining('Private to C')]);
 });
 it('retains private resolution without scanning unrelated private lists',()=>{
  let reads=0;const other=type('Other');Object.defineProperty(other,'privateMembers',{get(){reads++;return ['Unused'];}});
  expect(run('result = c.Hidden\nresult = c.Hidden',{projectClassMembers:[type('C',['Hidden']),other]})).toHaveLength(2);
  expect(reads).toBe(0);
 });
 it('resolves private names around known host members',()=>{
  const diagnostics=run('result = c.First\nresult = r.Value\nresult = c.Second\nresult = c.First',{projectClassMembers:[type('Other',['Unrelated']),type('C',['First','Second'])]});
  expect(diagnostics).toHaveLength(3);expect(calls.resolve.mock.calls.length).toBe(3);
 });
 it('uses current mutable private metadata on each invocation',()=>{
  const names:string[]=[],ctx={projectClassMembers:[type('C',names)]};
  expect(run('result = c.Hidden',ctx)).toEqual([]);names.push('Hidden');
  expect(run('result = c.Hidden',ctx)).toHaveLength(1);names.length=0;
  expect(run('result = c.Hidden',ctx)).toEqual([]);
 });
 it('preserves a public member that shadows the private name',()=>{
  const c=type('C',['Hidden']);c.members.push({name:'Hidden',kind:'property',moduleName:'C',returns:'Long'});
  expect(run('result = c.Hidden',{projectClassMembers:[c]})).toEqual([]);
 });
 it('preserves ambiguous project types',()=>{
  expect(run('result = c.Hidden',{projectClassMembers:[type('C',['Hidden']),type('c',['Hidden'])]})).toEqual([]);
 });
});

it('retains public presence when a project surface is not exhaustive',()=>{
 const c=type('C');c.members.push({name:'Value',kind:'property',returns:'Long',moduleName:'C'});
 const source='Sub Main()\nDim c As C\nresult = c.Value\nEnd Sub',offset=source.indexOf('c.Value')+2,ctx={projectClassMembers:[c]};
 expect(resolveExhaustiveMemberSurfaceAt(source,offset,ctx)).toBeUndefined();
 const presence=resolveMemberPresenceSurfaceAt(source,offset,ctx)!;
 expect(presence.owner).toBe('C');expect(presence.exhaustive).toBe(false);
 expect(presence.hasMember('vAlUe')).toBe(true);expect(presence.hasMember('Absent')).toBe(false);
});
it('preserves exhaustive API shape and negative class-member checks',()=>{
 const source='Sub Main()\nDim c As C\nresult = c.Value\nEnd Sub',offset=source.indexOf('c.Value')+2,ctx={projectClassMembers:[type('C',[],'class')]};
 const old=resolveExhaustiveMemberSurfaceAt(source,offset,ctx)!;
 expect(Object.keys(old).sort()).toEqual(['hasMember','owner']);expect(old.hasMember('Value')).toBe(false);
 const presence=resolveMemberPresenceSurfaceAt(source,offset,ctx)!;
 expect(presence.exhaustive).toBe(true);expect(presence.hasMember('Value')).toBe(false);
});
it('keeps unknown receivers unresolved',()=>{
 const source='Sub Main()\nDim c As Object\nresult = c.Value\nEnd Sub',offset=source.indexOf('c.Value')+2;
 expect(resolveMemberPresenceSurfaceAt(source,offset)).toBeUndefined();
 expect(resolveExhaustiveMemberSurfaceAt(source,offset)).toBeUndefined();
});
it('checks frozen public metadata without materializing completion documentation',()=>{
 let reads=0;const member={name:'Value',kind:'property' as const,moduleName:'C',get doc(){reads++;return undefined;}};
 const c=Object.freeze({...type('C'),members:Object.freeze([Object.freeze(member)])});
 const source='Sub Main()\nDim c As C\nresult = c.Value\nEnd Sub',offset=source.indexOf('c.Value')+2;
 const surface=resolveMemberPresenceSurfaceAt(source,offset,{projectClassMembers:Object.freeze([c])})!;
 expect(surface.hasMember('Value')).toBe(true);expect(reads).toBe(0);
});
