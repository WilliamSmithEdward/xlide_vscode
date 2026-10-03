import { afterEach, describe, expect, it, vi } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import * as typeInference from '../src/analyzer/diagnostics/typeInference';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import type { MemberCompletion } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkAssignmentTypes>[6]>;
function form(name: string, kind: VbaProjectClassMembers['kind'] = 'userform', returns = 'MSForms.Label'): VbaProjectClassMembers { return { name, kind, moduleName: name, members: [{ name: 'Answer', kind: 'property', moduleName: name, returns }] }; }
function prepare(source: string) {
 const mod = parseModule(source), sourceTokens = tokenizeCached(source).filter(token => token.kind !== 'comment');
 return (surfaces: VbaProjectClassMembers[]): Hit[] => {
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), hits: Hit[] = [];
  const ctx = { projectClassMembers: surfaces, parsedModule: mod, sourceTokens, withScanCache: new Map(), receiverTypeCache: new Map(), receiverChainCache: new Map(), memberSurfaceCache: new Map(), allowSetAssignmentRefinement: false };
  checkAssignmentTypes(source, mod, symbols, undefined, ctx, undefined, (...hit) => { hits.push(hit); }); return hits;
 };
}
function target(owner: string): MemberCompletion { return { name: 'Answer', kind: 'property', owner, returns: 'MSForms.Label', access: 'read-write' }; }
afterEach(() => { vi.restoreAllMocks(); });
describe('qualified form control owner membership', () => {
 it.each(['last', 'first'])('bounds actual metadata reads across procedures (%s)', (mode) => {
  let names = 0, kinds = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => form('F' + i)); surfaces[mode === 'first' ? 'unshift' : 'push'](form('Form1'));
  for (const surface of surfaces) { const name = surface.name, kind = surface.kind; Object.defineProperty(surface, 'name', { enumerable: true, get() { names++; return name; } }); Object.defineProperty(surface, 'kind', { enumerable: true, get() { kinds++; return kind; } }); }
  const body = 'Dim f As Form1\n' + 'Set f.Answer = Nothing\n'.repeat(100);
  expect(prepare('Sub P()\n' + body + 'End Sub\nSub Q()\n' + body + 'End Sub')(surfaces)).toHaveLength(200);
  expect(names).toBeLessThan(1000); expect(kinds).toBeLessThan(1000);
 });
 it('preserves exact case-sensitive owner names and the selected kind', () => {
  vi.spyOn(typeInference, 'resolveExactMemberCompletion').mockReturnValue(target('Form1'));
  const run = prepare('Sub P()\nDim f As Object\nSet f.Answer = Nothing\nEnd Sub');
  expect(run([form('Form1')])).toHaveLength(1);
  expect(run([form('form1')])).toEqual([]);
  expect(run([form('Form1', 'class')])).toEqual([]);
  expect(run([form('Project.Form1'), form(' Form1 ')])).toEqual([]);
  expect(run([form('Form1', 'class'), form('Form1')])).toHaveLength(1);
 });
 it('resumes for new owners and retains earlier names after a miss', () => {
  const resolve = vi.spyOn(typeInference, 'resolveExactMemberCompletion');
  for (const owner of ['A', 'B', 'A', 'Missing', 'B', 'C', 'A']) resolve.mockReturnValueOnce(target(owner));
  const hits = prepare('Sub P()\nDim f As Object\n' + 'Set f.Answer = Nothing\n'.repeat(7) + 'End Sub')([form('A'), form('Noise', 'class'), form('B'), form('C')]);
  expect(hits.map(hit => /on the form (\w+)/.exec(hit[1])?.[1])).toEqual(['A', 'B', 'A', 'B', 'C', 'A']);
 });
 it('refreshes owner misses and changed kinds in the same metadata array between invocations', () => {
  vi.spyOn(typeInference, 'resolveExactMemberCompletion').mockReturnValue(target('Form1'));
  const surfaces = [form('Other')], run = prepare('Sub P()\nDim f As Object\nSet f.Answer = Nothing\nEnd Sub');
  expect(run(surfaces)).toEqual([]);
  surfaces[0].name = 'Form1'; expect(run(surfaces)).toHaveLength(1);
  surfaces[0].kind = 'class'; expect(run(surfaces)).toEqual([]);
  surfaces.push(form('Form1')); expect(run(surfaces)).toHaveLength(1);
 });
 it.each(['let', 'other-return'])('leaves unused predicate controls unchanged (%s)', (mode) => {
  const source = 'Sub P()\nDim f As Form1\n' + (mode === 'let' ? 'f.Answer = 1\n' : 'Set f.Answer = Nothing\n').repeat(100) + 'End Sub';
  expect(prepare(source)([form('Form1', 'userform', mode === 'other-return' ? 'Object' : 'MSForms.Label')])).toEqual([]);
 });
});
