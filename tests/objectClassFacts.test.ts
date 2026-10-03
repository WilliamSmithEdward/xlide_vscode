import { describe, expect, it } from 'vitest';
import { checkObjectDefaultValues } from '../src/analyzer/diagnostics/rules/objectValues';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkObjectDefaultValues>[3]>;
function surface(members: VbaProjectClassMember[] = [], extra: Partial<VbaProjectClassMembers> = {}): VbaProjectClassMembers {
 return { name: 'Target', moduleName: 'Target', kind: 'class', exhaustive: true, members, ...extra };
}
function run(surfaces: VbaProjectClassMembers[], body: string, declaredType = 'Target'): Hit[] {
 const source = 'Sub P()\nDim c As ' + declaredType + '\nDim v As Variant\n' + body + '\nEnd Sub';
 const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const hits: Hit[] = [], factory = checkObjectDefaultValues(source, symbols, { projectClassMembers: surfaces }, (...hit) => { hits.push(hit); });
 for (const proc of mod.members) { if (proc.kind === 'Procedure') { const visit = factory(proc); if (visit) forEachStatementWithHeaders(source, proc.body, visit); } }
 return hits;
}
const item = (): VbaProjectClassMember => ({ name: 'Item', kind: 'property', moduleName: 'Target', defaultMember: true, signature: 'Item(i As Long)', returns: 'Long' });
describe('object class read facts', () => {
 it('bounds repeated class-name and default-member metadata reads', () => {
  let names = 0, defaults = 0;
  const others = Array.from({ length: 100 }, (_, i) => {
   const cls = surface([], { name: 'K' + i });
   Object.defineProperty(cls, 'name', { enumerable: true, get() { names++; return 'K' + i; } });
   return cls;
  });
  const members = Array.from({ length: 100 }, (_, i): VbaProjectClassMember => ({ name: 'M' + i, kind: 'property', moduleName: 'Target', get defaultMember() { defaults++; return false; } }));
  members.push(item());
  expect(run([...others, surface(members)], 'v = c\n'.repeat(100))).toHaveLength(100);
  expect(names).toBeLessThan(1000);
  expect(defaults).toBeLessThan(500);
 });
 it('keeps the first matching class and ignores other surface kinds', () => {
  const incomplete = surface([], { exhaustive: false });
  const complete = surface([item()], { name: 'TARGET' });
  const nonClass = surface([], { kind: 'userType' });
  expect(run([incomplete, complete], 'v = c')).toEqual([]);
  expect(run([nonClass, complete], 'v = c')).toHaveLength(1);
  expect(run([complete, incomplete], 'v = c')).toHaveLength(1);
 });
 it('rechecks defaults and missing enumerators when metadata changes between invocations', () => {
  const target = surface(), surfaces = [target];
  expect(run(surfaces, 'v = c')).toHaveLength(1);
  expect(run(surfaces, 'For Each v In c\nNext')).toHaveLength(1);
  target.members.push({ name: '_NewEnum', kind: 'property', moduleName: 'Target', returns: 'Object', attributes: [{ name: 'VB_UserMemId', valueRaw: '-4', span: { start: 0, end: 0 } }] });
  expect(run(surfaces, 'For Each v In c\nNext')).toEqual([]);
  target.members.push(item());
  expect(run(surfaces, 'v = c')).toHaveLength(1);
  target.members[target.members.length - 1].signature = 'Item()';
  expect(run(surfaces, 'v = c')).toEqual([]);
 });
 it('retains qualified/case-insensitive types and separate repeated enumerator reports', () => {
  const target = surface([{ name: '_NewEnum', kind: 'property', moduleName: 'Target', returns: 'Collection', attributes: [{ name: 'VB_UserMemId', valueRaw: '-4', span: { start: 0, end: 0 } }] }]);
  const hits = run([target], 'For Each v In C\nNext\nFor Each v In c\nNext', 'Library.TARGET');
  expect(hits).toHaveLength(2);
  expect(hits.every(hit => hit[1].includes("'451'"))).toBe(true);
  expect(hits[0][2].start).toBeLessThan(hits[1][2].start);
 });
});
