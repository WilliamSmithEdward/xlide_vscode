import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { hostObjectModelForToken } from '../src/analyzer/host/hostRegistry';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkAssignmentTypes>[6]>;
function surface(members: VbaProjectClassMember[] = []): VbaProjectClassMembers { return { name: 'Target', kind: 'class', moduleName: 'Target', exhaustive: true, members }; }
function item(): VbaProjectClassMember { return { name: 'Item', kind: 'property', moduleName: 'Target', defaultMember: true, signature: 'Item(i As Long)', returns: 'Long' }; }
function prepare(source: string) {
 const mod = parseModule(source);
 return (surfaces: VbaProjectClassMembers[], host = 'excel', vba7?: boolean): Hit[] => {
  const environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
  const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
  const hits: Hit[] = [];
  checkAssignmentTypes(source, mod, symbols, undefined, { projectClassMembers: surfaces, model: hostObjectModelForToken(host) }, activity, (...hit) => { hits.push(hit); });
  return hits;
 };
}
describe('assignment object type facts', () => {
 it('bounds repeated class-name and default-member metadata reads', () => {
  let names = 0, defaults = 0;
  const others = Array.from({ length: 100 }, (_, i) => {
   const cls = surface();
   Object.defineProperty(cls, 'name', { enumerable: true, get() { names++; return 'K' + i; } });
   return cls;
  });
  const members = Array.from({ length: 100 }, (_, i): VbaProjectClassMember => ({ name: 'M' + i, kind: 'property', moduleName: 'Target', get defaultMember() { defaults++; return false; } }));
  members.push({ ...item(), signature: 'Item()' });
  const run = prepare('Sub P()\nDim c As Target\n' + 'c = 5\n'.repeat(100) + 'End Sub');
  expect(run([...others, surface(members)])).toHaveLength(100);
  expect(names).toBeLessThan(1000);
  expect(defaults).toBeLessThan(500);
 });
 it('refreshes type/default facts after updates to the same metadata array and object', () => {
  const member = item(), surfaces = [surface([member])];
  const run = prepare('Sub P()\nDim c As Target\nc = 5\nEnd Sub');
  expect(run(surfaces).map(hit => hit[0])).toEqual(['setRequired']);
  member.signature = 'Item()';
  expect(run(surfaces).map(hit => hit[0])).toEqual(['readonlyMemberAssignment']);
  member.writable = true; member.letAccessor = true;
  expect(run(surfaces)).toEqual([]);
 });
 it('keeps array-element and scalar default rules distinct for the same declared type', () => {
  const member = item(), surfaces = [surface([member])];
  const run = prepare('Sub P()\nDim c As Target\nDim a(0) As Target\nc = 5\na(0) = 5\nEnd Sub');
  expect(run(surfaces).map(hit => hit[0])).toEqual(['setRequired', 'setRequired']);
  member.signature = 'Item()';
  expect(run(surfaces).map(hit => hit[0])).toEqual(['readonlyMemberAssignment']);
 });
 it('still checks statement-specific object state for a type with no default', () => {
  const hits = prepare('Sub P()\nDim c As Target\nc = 5\nSet c = New Target\nc = 5\nEnd Sub')([surface()]);
  expect(hits).toHaveLength(2);
  expect(hits[0][1]).toContain("'91'");
  expect(hits[1][1]).toContain("'438'");
 });
 it('preserves activity-specific local shadowing on reused parsed nodes', () => {
  const run = prepare('Dim c As Target\nSub P()\n#If VBA7 Then\nDim c As Long\nc = 1\n#Else\nc = 5\n#End If\nEnd Sub');
  const surfaces = [surface([{ ...item(), signature: 'Item()' }])];
  expect(run(surfaces, 'excel', true)).toEqual([]);
  expect(run(surfaces, 'excel', false).map(hit => hit[0])).toEqual(['readonlyMemberAssignment']);
  expect(run(surfaces, 'excel', true)).toEqual([]);
 });
 it('retains host defaults holding objects and read-only host defaults', () => {
  const hits = prepare('Sub P()\nDim p As Paragraph\nDim d As Document\np = 5\nd = 5\nEnd Sub')([], 'word');
  expect(hits.map(hit => hit[0])).toEqual(['invalidPropertyUse', 'readonlyMemberAssignment']);
 });
});
