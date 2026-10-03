import { describe, expect, it } from 'vitest';
import { checkSetAssignments } from '../src/analyzer/diagnostics/rules/assignments';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkSetAssignments>[4]>;
function control(name = 'Answer', returns = 'MSForms.Label'): VbaProjectClassMember { return { name, returns, kind: 'property', moduleName: 'Form1' }; }
function form(members: VbaProjectClassMember[] = []): VbaProjectClassMembers { return { name: 'Form1', kind: 'userform', moduleName: 'Form1', members }; }
function prepare(source: string) {
 const mod = parseModule(source);
 return (ctx: MemberCompletionContext): Hit[] => {
  const symbols = buildModuleSymbols('Form1', 'userform', source, { parsedModule: mod });
  const hits: Hit[] = [], visitor = checkSetAssignments(source, symbols, undefined, ctx, (...hit) => { hits.push(hit); });
  for (const proc of mod.members) if (proc.kind === 'Procedure') { const visit = visitor(proc); if (visit) forEachStatementWithHeaders(source, proc.body, visit); }
  return hits;
 };
}
describe('Set current-form/control facts', () => {
 it('bounds repeated surface and member reads across target/value queries and procedures', () => {
  let names = 0, members = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => { const f = form(); Object.defineProperty(f, 'name', { enumerable: true, get() { names++; return 'F' + i; } }); return f; });
  const controls = Array.from({ length: 100 }, (_, i) => ({ ...control(), get name() { members++; return 'C' + i; } }));
  controls.push(control()); surfaces.push(form(controls));
  const body = 'Set Answer = Nothing\nSet t = Answer\n'.repeat(50);
  const hits = prepare('Sub P()\nDim t As MSForms.TextBox\n' + body + 'End Sub\nSub Q()\nDim t As MSForms.TextBox\n' + body + 'End Sub')({ meProjectType: 'FORM1', projectClassMembers: surfaces });
  expect(hits).toHaveLength(200);
  expect(names).toBeLessThan(1000); expect(members).toBeLessThan(500);
 });
 it('caches missing controls and missing form surfaces', () => {
  let names = 0;
  const controls = Array.from({ length: 100 }, (_, i) => ({ ...control(), get name() { names++; return 'C' + i; } }));
  const run = prepare('Sub P()\n' + 'Set Missing = Nothing\n'.repeat(100) + 'End Sub');
  expect(run({ meProjectType: 'Form1', projectClassMembers: [form(controls)] })).toEqual([]);
  expect(names).toBeLessThan(500);
  names = 0;
  const surfaces = Array.from({ length: 100 }, (_, i) => { const f = form(); Object.defineProperty(f, 'name', { enumerable: true, get() { names++; return 'Other' + i; } }); return f; });
  expect(run({ meProjectType: 'Form1', projectClassMembers: surfaces })).toEqual([]);
  expect(names).toBeLessThan(500);
 });
 it('keeps the first matching userform even when it has no matching control', () => {
  const run = prepare('Sub P()\nSet Answer = Nothing\nEnd Sub');
  const earlierClass = { ...form([control()]), kind: 'class' as const };
  expect(run({ meProjectType: 'Form1', projectClassMembers: [earlierClass, form([]), form([control()])] })).toEqual([]);
  expect(run({ meProjectType: 'Form1', projectClassMembers: [earlierClass, { ...form([control()]), exhaustive: false }, form([])] })).toHaveLength(1);
 });
 it('skips same-name non-control members and retains the first MSForms match', () => {
  const hits = prepare('Sub P()\nDim t As MSForms.TextBox\nSet t = aNsWeR\nEnd Sub')({ meProjectType: 'form1', projectClassMembers: [form([control('Answer', 'String'), control('ANSWER', 'MSForms.Label'), control('Answer', 'MSForms.TextBox')])] });
  expect(hits).toHaveLength(1); expect(hits[0][1]).toContain('MSForms.Label');
 });
 it('refreshes misses and changed metadata between invocations using the same array/object', () => {
  const f = form(), surfaces = [f], ctx = { meProjectType: 'Form1', projectClassMembers: surfaces };
  const run = prepare('Sub P()\nSet Answer = Nothing\nEnd Sub');
  expect(run(ctx)).toEqual([]);
  f.members.push(control()); expect(run(ctx)).toHaveLength(1);
  f.members[0].returns = 'String'; expect(run(ctx)).toEqual([]);
  f.members[0].returns = 'MSForms.Label'; f.name = 'Other'; expect(run(ctx)).toEqual([]);
  f.name = 'FORM1'; expect(run(ctx)).toHaveLength(1);
 });
 it('retains local value shadows and target declaration guards', () => {
  const run = prepare('Sub P()\nDim t As MSForms.TextBox\nDim Answer As MSForms.TextBox\nSet t = Answer\nEnd Sub\nSub Q()\nDim Answer As Long\nSet Answer = Nothing\nEnd Sub');
  const hits = run({ meProjectType: 'Form1', projectClassMembers: [form([control()])] });
  expect(hits).toHaveLength(1); expect(hits[0][1]).toContain('declared as Long');
 });
 it('does not inspect unused control metadata without a form context', () => {
  const f = form(); Object.defineProperty(f, 'members', { get() { throw Error('Unused control metadata read'); } });
  expect(prepare('Sub P()\nDim n As Long\nSet n = Nothing\nEnd Sub')({ projectClassMembers: [f] })).toHaveLength(1);
 });
});
