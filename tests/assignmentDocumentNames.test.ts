import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes, checkSetAssignments } from '../src/analyzer/diagnostics/rules/assignments';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { hostObjectModelForToken } from '../src/analyzer/host/hostRegistry';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkSetAssignments>[4]>;
function surface(name: string, kind: VbaProjectClassMembers['kind'] = 'document'): VbaProjectClassMembers { return { name, kind, moduleName: name, members: [] }; }
function prepare(source: string, set: boolean) {
 const mod = parseModule(source);
 return (ctx: MemberCompletionContext): Hit[] => {
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), hits: Hit[] = [];
  const push = (...hit: Hit) => { hits.push(hit); };
  if (set) { const visitor = checkSetAssignments(source, symbols, undefined, ctx, push); for (const proc of mod.members) if (proc.kind === 'Procedure') { const visit = visitor(proc); if (visit) forEachStatementWithHeaders(source, proc.body, visit); } }
  else checkAssignmentTypes(source, mod, symbols, undefined, ctx, undefined, push);
  return hits;
 };
}
for (const set of [false, true]) describe(set ? 'Set document names' : 'Let document names', () => {
 const assignment = (name: string) => (set ? 'Set ' : '') + name + (set ? ' = Nothing\n' : ' = 1\n');
 it.each(['Sheet1', 'Missing', 'distinct'])('examines each surface once across procedures (%s)', (target) => {
  let kinds = 0, names = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => surface('K' + i, i % 2 === 0 ? 'class' : 'document')); surfaces.push(surface('Sheet1'));
  for (const item of surfaces) { const kind = item.kind, name = item.name; Object.defineProperty(item, 'kind', { enumerable: true, get() { kinds++; return kind; } }); Object.defineProperty(item, 'name', { enumerable: true, get() { names++; return name; } }); }
  const body = Array.from({ length: 100 }, (_, i) => assignment(target === 'distinct' ? 'Missing' + i : target)).join('');
  const hits = prepare('Sub P()\n' + body + 'End Sub\nSub Q()\n' + body + 'End Sub', set)({ projectClassMembers: surfaces });
  expect(hits).toHaveLength(target === 'Sheet1' ? 200 : 0);
  expect(kinds).toBe(101); expect(names).toBe(51);
 });
 it('stops at an early hit without inspecting the unused tail', () => {
  const unused = surface('Unused'); Object.defineProperty(unused, 'kind', { get() { throw Error('Unused metadata'); } });
  expect(prepare('Sub P()\n' + assignment('sheet1').repeat(100) + 'End Sub', set)({ projectClassMembers: [surface('Sheet1'), unused] })).toHaveLength(100);
 });
 it('resumes for later names and retains earlier names after a miss exhausts the surfaces', () => {
  const source = 'Sub P()\n' + ['A', 'B', 'A', 'Missing', 'B', 'C', 'A'].map(assignment).join('') + 'End Sub';
  const hits = prepare(source, set)({ projectClassMembers: [surface('A'), surface('Noise', 'class'), surface('B'), surface('C')] });
  expect(hits.map(hit => source.slice(hit[2].start, hit[2].end))).toEqual(['A', 'B', 'A', 'B', 'C', 'A']);
 });
 it('uses any document match despite duplicate non-document names, with exact case folding', () => {
  const source = 'Sub P()\n' + assignment('sHeEt1') + assignment('Other') + 'End Sub';
  const hits = prepare(source, set)({ projectClassMembers: [surface('Sheet1', 'class'), surface('SHEET1', 'document'), surface('Other', 'userform')] });
  expect(hits).toHaveLength(1); expect(source.slice(hits[0][2].start, hits[0][2].end)).toBe('sHeEt1');
  expect(prepare('Sub P()\n' + assignment('Sheet1') + 'End Sub', set)({ projectClassMembers: [surface('Project.Sheet1'), surface(' Sheet1 ')] })).toEqual([]);
 });
 it('skips document metadata for declared targets and non-assignment statements', () => {
  const unused = surface('Sheet1'); Object.defineProperty(unused, 'kind', { get() { throw Error('Unused metadata'); } });
  const source = 'Sub P()\nDim Sheet1 As ' + (set ? 'Object' : 'Long') + '\n' + assignment('Sheet1') + 'End Sub';
  expect(prepare(source, set)({ projectClassMembers: [unused] })).toEqual([]);
  expect(prepare('Sub P()\nBeep\nEnd Sub', set)({ projectClassMembers: [unused] })).toEqual([]);
 });
 it('refreshes the same mutable context and array on subsequent invocations', () => {
  const surfaces = [surface('Sheet1', 'class')], ctx = { projectClassMembers: surfaces }, run = prepare('Sub P()\n' + assignment('Sheet1') + 'End Sub', set);
  expect(run(ctx)).toEqual([]);
  surfaces[0].kind = 'document'; expect(run(ctx)).toHaveLength(1);
  surfaces[0].name = 'Other'; expect(run(ctx)).toEqual([]);
  surfaces.push(surface('SHEET1')); expect(run(ctx)).toHaveLength(1);
 });
 it('retains host-specific document diagnostics', () => {
  const run = prepare('Sub P()\n' + assignment('ThisDocument') + 'End Sub', set), projectClassMembers = [surface('ThisDocument')];
  const excel = run({ projectClassMembers, model: hostObjectModelForToken('excel') });
  const word = run({ projectClassMembers, model: hostObjectModelForToken('word') });
  expect(excel).toHaveLength(1); expect(word).toHaveLength(1);
  expect(word[0][1]).toContain('Invalid use of property');
  if (!set) { expect(excel[0][0]).toBe('setRequired'); expect(word[0][0]).toBe('setRequiresObject'); }
 });
});
