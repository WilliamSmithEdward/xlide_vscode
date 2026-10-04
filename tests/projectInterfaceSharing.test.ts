import { describe, expect, it } from 'vitest';
import { checkSetAssignments } from '../src/analyzer/diagnostics/rules/assignments';
import * as typeInference from '../src/analyzer/diagnostics/typeInference';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkSetAssignments>[4]>;
function surface(name: string, kind: VbaProjectClassMembers['kind'] = 'class'): VbaProjectClassMembers { return { name, kind, moduleName: name, members: [], exhaustive: true }; }
function prepare(source: string) {
 const mod = parseModule(source), sourceTokens = tokenizeCached(source).filter(token => token.kind !== 'comment');
 return (surfaces: VbaProjectClassMembers[]): Hit[] => {
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), hits: Hit[] = [];
  const ctx = { projectClassMembers: surfaces, parsedModule: mod, sourceTokens, withScanCache: new Map(), receiverTypeCache: new Map(), receiverChainCache: new Map(), memberSurfaceCache: new Map(), allowSetAssignmentRefinement: false };
  const visitor = checkSetAssignments(source, symbols, undefined, ctx, (...hit) => { hits.push(hit); });
  for (const proc of mod.members) if (proc.kind === 'Procedure') { const visit = visitor(proc); if (visit) forEachStatementWithHeaders(source, proc.body, visit); }
  return hits;
 };
}
describe('project interface sharing', () => {
 it.each(['first', 'last', 'missing'])('bounds actual Implements reads across procedures (%s)', (mode) => {
  let reads = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => surface('K' + i)); surfaces.push(surface('Box'));
  if (mode !== 'missing') surfaces[mode === 'first' ? 'unshift' : 'push']({ ...surface('Shared'), implements: ['Box', 'K0'] });
  for (const item of surfaces) { const implemented = item.implements; Object.defineProperty(item, 'implements', { enumerable: true, get() { reads++; return implemented; } }); }
  const body = 'Dim c As Box\n' + 'Set c = New K0\n'.repeat(100);
  expect(prepare('Sub P()\n' + body + 'End Sub\nSub Q()\n' + body + 'End Sub')(surfaces)).toHaveLength(mode === 'missing' ? 200 : 0);
  expect(reads).toBeLessThan(200);
 });
 it('bounds reads for distinct and reversed pairs, retaining misses and earlier hits', () => {
  let reads = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => surface('K' + i)); surfaces.push(surface('Box'));
  for (const item of surfaces) Object.defineProperty(item, 'implements', { enumerable: true, get() { reads++; return []; } });
  const body = Array.from({ length: 100 }, (_, i) => 'Dim c' + i + ' As K' + i + '\nSet c' + i + ' = New Box\n').join('');
  expect(prepare('Sub P()\n' + body + 'End Sub')(surfaces)).toHaveLength(100); expect(reads).toBeLessThan(300);
  const lookup = typeInference.createProjectInterfaceSharingLookup({ projectClassMembers: [{ ...surface('First'), implements: ['A', 'B'] }, { ...surface('Last'), implements: ['B', 'C'] }] });
  for (const [a, b, expected] of [['a', 'b', true], ['b', 'c', true], ['a', 'missing', false], ['b', 'a', true], ['c', 'a', false], ['c', 'b', true]] as const) expect(lookup(a, b)).toBe(expected);
 });
 it('keeps all metadata kinds, full lowercased names and any common surface', () => {
  for (const kind of ['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const) {
   const lookup = typeInference.createProjectInterfaceSharingLookup({ projectClassMembers: [{ ...surface('Shared', kind), implements: ['A', 'b'] }] });
   expect(lookup('a', 'b')).toBe(true); expect(lookup('project.a', 'b')).toBe(false);
  }
  const lookup = typeInference.createProjectInterfaceSharingLookup({ projectClassMembers: [{ ...surface('Same'), implements: [' a ', 'b'] }, { ...surface('Same'), implements: ['a', 'b'] }] });
  expect(lookup('a', 'b')).toBe(true);
  const split = typeInference.createProjectInterfaceSharingLookup({ projectClassMembers: [{ ...surface('Same'), implements: ['a'] }, { ...surface('Same'), implements: ['b'] }] });
  expect(split('a', 'b')).toBe(false);
  expect(typeInference.createProjectInterfaceSharingLookup({ projectClassMembers: [{ ...surface('Trim'), implements: [' a ', 'b'] }] })('a', 'b')).toBe(false);
 });
 it('does not enumerate all pairs of a many-interface surface', () => {
  let reads = 0;
  const item = surface('Shared'), names = Array.from({ length: 2000 }, (_, i) => 'I' + i);
  Object.defineProperty(item, 'implements', { get() { reads++; return names; } });
  const lookup = typeInference.createProjectInterfaceSharingLookup({ projectClassMembers: [item] });
  for (let i = 0; i < 100; i++) expect(lookup('i' + i, 'i1999')).toBe(true);
  expect(lookup('i1999', 'i0')).toBe(true); expect(lookup('missing', 'i0')).toBe(false); expect(reads).toBe(1);
 });
 it('preserves direct casts, same types and generic/host short circuits', () => {
  for (const implementedOn of ['Box', 'K0']) {
   const surfaces = [surface('Box'), surface('K0'), surface('Noise')];
   surfaces[implementedOn === 'Box' ? 0 : 1].implements = [implementedOn === 'Box' ? 'K0' : 'Box'];
   Object.defineProperty(surfaces[2], 'implements', { get() { throw Error('Unexpected sharing scan'); } });
   expect(prepare('Sub P()\nDim c As Box\nSet c = New K0\nEnd Sub')(surfaces)).toEqual([]);
  }
  const item = surface('Noise'); Object.defineProperty(item, 'implements', { get() { throw Error('Unexpected sharing scan'); } });
  expect(prepare('Sub P()\nDim c As Object\nSet c = New Worksheet\nDim w As Worksheet\nSet w = New Worksheet\nEnd Sub')([item])).toEqual([]);
 });
 it('refreshes mutable Implements arrays and appended metadata on the next pass', () => {
  const surfaces = [surface('Box'), surface('K0'), { ...surface('Shared'), implements: ['Box', 'K0'] }], run = prepare('Sub P()\nDim c As Box\nSet c = New K0\nEnd Sub');
  expect(run(surfaces)).toEqual([]); surfaces[2].implements!.pop(); expect(run(surfaces)).toHaveLength(1);
  surfaces[2].implements!.push('k0'); expect(run(surfaces)).toEqual([]);
  surfaces[2].implements = []; expect(run(surfaces)).toHaveLength(1);
  surfaces.push({ ...surface('Later'), implements: ['box', 'k0'] }); expect(run(surfaces)).toEqual([]);
 });
 it('matches default compatibility while actual classes and pair order change', () => {
  const ctx = { projectClassMembers: [surface('Box'), surface('K0'), surface('K1'), { ...surface('Shared', 'enum'), implements: ['Box', 'K0'] }] }, resolve = typeInference.createObjectAssignmentTypeResolver(ctx), sharing = typeInference.createProjectInterfaceSharingLookup(ctx);
  for (const expected of ['Box', 'K0', 'K1', 'Object', 'Worksheet', 'Collection', 'Missing']) for (const type of ['Box', 'K0', 'K1', 'Nothing', 'Object', 'Long', 'Worksheet']) {
   const actual = { type, label: type };
   expect(typeInference.objectAssignmentIncompatibilityReason(expected, actual, ctx, resolve, sharing)).toBe(typeInference.objectAssignmentIncompatibilityReason(expected, actual, ctx));
  }
  const body = 'Dim c As Box\nDim o As Object\nSet o = New K0\nSet c = o\nSet o = New K1\nSet c = o\nSet o = New K0\nSet c = o\n';
  expect(prepare('Sub P()\n' + body + 'End Sub')(ctx.projectClassMembers)).toHaveLength(1);
 });
});
