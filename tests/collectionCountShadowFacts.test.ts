import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkAssignmentTypes>[6]>;
function surface(name: string, kind: VbaProjectClassMembers['kind'] = 'class'): VbaProjectClassMembers { return { name, kind, moduleName: name, exhaustive: true, members: [{ name: 'Count', kind: 'property', moduleName: name, returns: 'Long', writable: true, letAccessor: true }] }; }
function prepare(source: string) {
 const mod = parseModule(source), sourceTokens = tokenizeCached(source).filter(token => token.kind !== 'comment');
 return (ctx: MemberCompletionContext): Hit[] => {
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), hits: Hit[] = [];
  checkAssignmentTypes(source, mod, symbols, undefined, { ...ctx, parsedModule: mod, sourceTokens, withScanCache: new Map(), receiverTypeCache: new Map(), receiverChainCache: new Map(), memberSurfaceCache: new Map(), allowSetAssignmentRefinement: false }, undefined, (...hit) => { hits.push(hit); });
  return hits;
 };
}
describe('Collection.Count project-name predicate', () => {
 it.each([false, true])('bounds project name reads across procedures (shadow=%s)', (shadow) => {
  let reads = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => surface('K' + i)); if (shadow) surfaces.push(surface('CoLlEcTiOn'));
  for (const item of surfaces) { const name = item.name; Object.defineProperty(item, 'name', { enumerable: true, get() { reads++; return name; } }); }
  const body = 'Dim c As Collection\n' + 'c.Count = 1\n'.repeat(100);
  expect(prepare('Sub P()\n' + body + 'End Sub\nSub Q()\n' + body + 'End Sub')({ projectClassMembers: surfaces })).toHaveLength(shadow ? 0 : 200);
  expect(reads).toBeLessThan(500);
 });
 it.each(['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const)('retains any-kind matching (%s)', (kind) => {
  const hits = prepare('Sub P()\nDim c As Collection\nc.Count = 1\nEnd Sub')({ projectClassMembers: [surface('cOlLeCtIoN', kind)] });
  expect(hits).toEqual([]);
 });
 it('does not treat qualified or whitespace names as Collection shadows', () => {
  const run = prepare('Sub P()\nDim c As Collection\nc.Count = 1\nEnd Sub');
  for (const name of ['Project.Collection', ' Collection ']) {
   const hits = run({ projectClassMembers: [surface(name)] });
   expect(hits).toHaveLength(1); expect(hits[0][1]).toContain("a Collection's Count is a Function");
  }
 });
 it('does not inspect unused project names for ordinary scalar assignments', () => {
  const unused = surface('Unused'); Object.defineProperty(unused, 'name', { get() { throw Error('Unused name'); } });
  expect(prepare('Sub P()\nDim n As Long\nn = 1\nEnd Sub')({ projectClassMembers: [unused] })).toEqual([]);
 });
 it('refreshes missing and matching results across invocations with the same metadata', () => {
  const surfaces = [surface('Other')], ctx = { projectClassMembers: surfaces }, run = prepare('Sub P()\nDim c As Collection\nc.Count = 1\nEnd Sub');
  expect(run(ctx)).toHaveLength(1);
  surfaces[0].name = 'Collection'; expect(run(ctx)).toEqual([]);
  surfaces[0].name = 'Other'; expect(run(ctx)).toHaveLength(1);
  surfaces.push(surface('COLLECTION')); expect(run(ctx)).toEqual([]);
 });
 it('retains receiver type and argument guards', () => {
  const source = 'Sub P()\nDim c As Long\nc.Count = 1\nEnd Sub';
  expect(prepare(source)({ projectClassMembers: [] })).toEqual([]);
  const hits = prepare('Sub P()\nDim c As Collection\nc.Count(1) = 1\nEnd Sub')({ projectClassMembers: [] });
  expect(hits.some(hit => hit[1].includes("a Collection's Count is a Function"))).toBe(false);
 });
});
