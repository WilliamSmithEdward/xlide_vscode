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
describe('Set object type resolution', () => {
 it.each(['Nothing', 'New Box', 'ActiveSheet'])('bounds actual class metadata reads across procedures for %s', (rhs) => {
  let names = 0, kinds = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => surface('K' + i)); surfaces.push(surface('Box'));
  for (const item of surfaces) { const name = item.name, kind = item.kind; Object.defineProperty(item, 'name', { enumerable: true, get() { names++; return name; } }); Object.defineProperty(item, 'kind', { enumerable: true, get() { kinds++; return kind; } }); }
  const body = 'Dim c As Box\n' + ('Set c = ' + rhs + '\n').repeat(100);
  expect(prepare('Sub P()\n' + body + 'End Sub\nSub Q()\n' + body + 'End Sub')(surfaces)).toHaveLength(rhs === 'ActiveSheet' ? 200 : 0);
  expect(names).toBeLessThan(500); expect(kinds).toBeLessThan(500);
 });
 it('bounds metadata reads for distinct targets and missing queries', () => {
  let reads = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => surface('K' + i));
  for (const item of surfaces) { const name = item.name; Object.defineProperty(item, 'name', { enumerable: true, get() { reads++; return name; } }); }
  const source = 'Sub P()\n' + Array.from({ length: 100 }, (_, i) => 'Dim c' + i + ' As K' + i + '\nSet c' + i + ' = Nothing\n').join('') + 'Dim m As Missing\n' + 'Set m = Nothing\n'.repeat(100) + 'End Sub';
  expect(prepare(source)(surfaces)).toEqual([]); expect(reads).toBeLessThan(400);
 });
 it('preserves unique matches, all eligible kinds and duplicate ambiguity', () => {
  for (const kind of ['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const) {
   const ctx = { projectClassMembers: [surface('bOX', kind)] }, resolve = typeInference.createObjectAssignmentTypeResolver(ctx);
   expect(resolve('BOX')).toEqual(typeInference.resolveKnownObjectAssignmentType('BOX', ctx));
   expect(resolve('BOX')).toBe(resolve('BOX'));
  }
  for (const surfaces of [[surface('Box'), surface('BOX')], [surface('Box'), surface('BOX'), surface('box')], [surface('Box', 'enum'), surface('BOX')]]) {
   const ctx = { projectClassMembers: surfaces };
   expect(typeInference.createObjectAssignmentTypeResolver(ctx)('Box')).toEqual(typeInference.resolveKnownObjectAssignmentType('Box', ctx));
  }
  const same = surface('Box'); expect(typeInference.createObjectAssignmentTypeResolver({ projectClassMembers: [same, same] })('Box')).toBeUndefined();
 });
 it('keeps raw generic/host display spellings and canonical project display/implements', () => {
  const item = { ...surface('bOX'), implements: ['Collection', 'IFoo'] }, ctx = { projectClassMembers: [item] }, resolve = typeInference.createObjectAssignmentTypeResolver(ctx);
  for (const type of ['Object', 'object', ' Object ', 'Collection', 'COLLECTION', 'Worksheet', 'Excel.Worksheet', 'DAO.Recordset', 'Box', 'BOX', 'Project.Box', 'Box()', 'Long', 'Variant', 'Missing', '', undefined]) expect(resolve(type)).toEqual(typeInference.resolveKnownObjectAssignmentType(type, ctx));
  expect(resolve('Object')?.display).toBe('Object'); expect(resolve('object')?.display).toBe('object');
  expect(resolve('Box')).toEqual({ kind: 'project', display: 'bOX', key: 'box', implements: item.implements });
 });
 it('keeps generic, scalar, host and library queries from inspecting project metadata', () => {
  const item = surface('Box'); Object.defineProperty(item, 'kind', { get() { throw Error('Unexpected project scan'); } });
  const ctx = { projectClassMembers: [item] }, resolve = typeInference.createObjectAssignmentTypeResolver(ctx);
  for (const type of ['Object', 'Collection', 'Long', 'Variant', undefined, '', 'Worksheet', 'Excel.Worksheet', 'DAO.Recordset']) expect(resolve(type)).toEqual(typeInference.resolveKnownObjectAssignmentType(type, ctx));
 });
 it('refreshes misses, names, kinds, ambiguity and implements on subsequent Set passes', () => {
  const surfaces = [surface('Other')], run = prepare('Sub P()\nDim c As Box\nSet c = ActiveSheet\nEnd Sub');
  expect(run(surfaces)).toEqual([]); surfaces[0].name = 'Box'; expect(run(surfaces)).toHaveLength(1);
  surfaces[0].kind = 'enum'; expect(run(surfaces)).toEqual([]);
  surfaces.push(surface('BOX')); expect(run(surfaces)).toHaveLength(1);
  surfaces.push(surface('Box')); expect(run(surfaces)).toEqual([]);
  const item = { ...surface('Square'), implements: ['Collection'] }, ctx = { projectClassMembers: [item] };
  expect(typeInference.createObjectAssignmentTypeResolver(ctx)('Square')).toMatchObject({ implements: ['Collection'] });
  item.implements = ['IFoo']; expect(typeInference.createObjectAssignmentTypeResolver(ctx)('Square')).toMatchObject({ implements: ['IFoo'] });
 });
 it('uses the cached resolver for compatibility without caching statement-specific actual values', () => {
  const ctx = { projectClassMembers: [surface('Box'), { ...surface('Square'), implements: ['Collection'] }] }, resolve = typeInference.createObjectAssignmentTypeResolver(ctx);
  for (const expected of ['Box', 'Object', 'Collection', 'Missing']) for (const type of ['Box', 'Square', 'Object', 'Nothing', 'Long', 'Worksheet', 'Scripting.Dictionary', 'Worksheet Or Chart']) {
   const actual = { type, label: type };
   expect(typeInference.objectAssignmentIncompatibilityReason(expected, actual, ctx, resolve)).toBe(typeInference.objectAssignmentIncompatibilityReason(expected, actual, ctx));
  }
  const source = 'Sub P()\nDim c As Box\nDim o As Object\nSet o = New Box\nSet c = o\nSet o = New Square\nSet c = o\nEnd Sub';
  expect(prepare(source)(ctx.projectClassMembers)).toHaveLength(1);
 });
});
