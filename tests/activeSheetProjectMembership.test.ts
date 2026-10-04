import { afterEach, describe, expect, it, vi } from 'vitest';
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
function isolateGuard() {
 vi.spyOn(typeInference, 'createObjectAssignmentTypeResolver').mockReturnValue(() => ({ kind: 'generic', display: 'Object', key: 'object' }));
 vi.spyOn(typeInference, 'objectAssignmentIncompatibilityReason').mockReturnValue(undefined);
}
afterEach(() => { vi.restoreAllMocks(); });
describe('ActiveSheet project class membership', () => {
 it.each(['Object', 'Worksheet'])('bounds actual metadata reads across procedures for %s misses', (type) => {
  let names = 0, kinds = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => surface('K' + i)); surfaces.push(surface('Box'));
  for (const item of surfaces) { const name = item.name, kind = item.kind; Object.defineProperty(item, 'name', { enumerable: true, get() { names++; return name; } }); Object.defineProperty(item, 'kind', { enumerable: true, get() { kinds++; return kind; } }); }
  const body = 'Dim c As ' + type + '\n' + 'Set c = ActiveSheet\n'.repeat(100);
  expect(prepare('Sub P()\n' + body + 'End Sub\nSub Q()\n' + body + 'End Sub')(surfaces)).toEqual([]);
  expect(names).toBeLessThan(500); expect(kinds).toBeLessThan(500);
 });
 it('keeps real class and Collection mismatches', () => {
  for (const type of ['Box', 'Collection']) {
   const hits = prepare('Sub P()\nDim c As ' + type + '\nSet c = ActiveSheet\nEnd Sub')([surface('Box')]);
   expect(hits).toHaveLength(1); expect(hits[0][1]).toContain('ActiveSheet holds a sheet');
  }
 });
 it('preserves any matching class, case folding, and exact full metadata names', () => {
  isolateGuard();
  const run = prepare('Sub P()\nDim c As Box\nSet c = ActiveSheet\nEnd Sub');
  expect(run([surface('bOX')])).toHaveLength(1);
  expect(run([surface('Box', 'userform'), surface('BOX')])).toHaveLength(1);
  expect(run([surface('Box'), surface('BOX')])).toHaveLength(1);
  expect(run([surface('Box', 'document'), surface('Box', 'userform')])).toEqual([]);
  expect(run([surface('Project.Box'), surface(' Box ')])).toEqual([]);
 });
 it('retains prior names and resumes lookup after hits and misses', () => {
  isolateGuard();
  const source = 'Sub P()\nDim a As A\nDim b As B\nDim c As C\nDim o As Object\nSet a = ActiveSheet\nSet b = ActiveSheet\nSet a = ActiveSheet\nSet o = ActiveSheet\nSet c = ActiveSheet\nSet b = ActiveSheet\nEnd Sub';
  expect(prepare(source)([surface('A'), surface('Noise', 'document'), surface('B'), surface('C')])).toHaveLength(5);
 });
 it('refreshes misses, changed names, kinds, and appended classes between invocations', () => {
  isolateGuard();
  const surfaces = [surface('Other')], run = prepare('Sub P()\nDim c As Box\nSet c = ActiveSheet\nEnd Sub');
  expect(run(surfaces)).toEqual([]);
  surfaces[0].name = 'Box'; expect(run(surfaces)).toHaveLength(1);
  surfaces[0].kind = 'userform'; expect(run(surfaces)).toEqual([]);
  surfaces.push(surface('BOX')); expect(run(surfaces)).toHaveLength(1);
 });
 it.each(['collection', 'shadow', 'other-value', 'qualified', 'existing-reason'])('does not inspect metadata when the guard is bypassed (%s)', (mode) => {
  isolateGuard();
  vi.spyOn(typeInference, 'inferArgumentType').mockReturnValue({ type: 'Object', label: 'object expression' });
  if (mode === 'existing-reason') vi.spyOn(typeInference, 'objectAssignmentIncompatibilityReason').mockReturnValue('Earlier mismatch');
  let reads = 0;
  const item = surface('Object');
  Object.defineProperty(item, 'name', { enumerable: true, get() { reads++; return 'Object'; } });
  Object.defineProperty(item, 'kind', { enumerable: true, get() { reads++; return 'class'; } });
  const source = 'Sub P()\nDim c As ' + (mode === 'collection' ? 'Collection' : 'Object') + '\n' + (mode === 'shadow' ? 'Dim ActiveSheet As Object\n' : '') + 'Set c = ' + (mode === 'other-value' ? 'Nothing' : mode === 'qualified' ? 'Application.ActiveSheet' : 'ActiveSheet') + '\nEnd Sub';
  const hits = prepare(source)([item]);
  expect(hits).toHaveLength(mode === 'collection' || mode === 'existing-reason' ? 1 : 0);
  expect(reads).toBe(0);
 });
});
