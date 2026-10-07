import { describe, expect, it, vi } from 'vitest';
import { checkMemberNotFound } from '../src/analyzer/diagnostics/rules/undeclared';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { walkProcedureStatements } from '../src/analyzer/diagnostics/walker';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

const controls = [{ name: 'T1', type: 'MSForms.TextBox' }, { name: 'T2', type: 'MSForms.TextBox' }];
function form(name = 'F1'): VbaProjectClassMembers {
 return { name, moduleName: name, kind: 'userform', exhaustive: true,
  members: controls.map(c => ({ name: c.name, kind: 'property', returns: c.type, moduleName: name })) };
}
function run(source: string, context: MemberCompletionContext) {
 const parsedModule = parseModule(source), push = vi.fn();
 const ctx = { ...context, parsedModule, sourceTokens: tokenizeCached(source).filter(t => t.kind !== 'comment'),
  receiverTypeCache: new Map(), receiverChainCache: new Map(), memberSurfaceCache: new Map(), withScanCache: new Map() };
 walkProcedureStatements(parsedModule, undefined, [checkMemberNotFound(source, ctx, push)], undefined, { source, takes: [true] });
 return push.mock.calls;
}
function context(meProjectType: string | undefined = 'F1'): MemberCompletionContext {
 return { projectClassMembers: [form()], meProjectType, meType: 'VBA.UserForm', implicitMembers: controls };
}
const fixtures = [
 { label: 'local with a current form', header: 'Sub Main()', declaration: 'Dim r As Range', expression: 'r.Value', me: 'F1' },
 { label: 'local without a current form', header: 'Sub Main()', declaration: 'Dim r As Range', expression: 'r.Value', me: undefined },
 { label: 'parameter shadow', header: 'Sub Main(ByVal r As Range)', declaration: '', expression: 'r.Value', me: 'F1' },
 { label: 'case-insensitive local shadow', header: 'Sub Main()', declaration: 'Dim r As Range', expression: 'R.Value', me: 'F1' },
 { label: 'host receiver without a current form', header: 'Sub Main()', declaration: '', expression: 'ActiveSheet.Name', me: undefined },
 { label: 'nested local declaration', header: 'Sub Main()', declaration: 'If True Then\nDim r As Range\nEnd If', expression: 'r.Value', me: 'F1' },
];
describe.each(fixtures)('$label', fixture => {
 it.each([1, 1000])('avoids scanning %i unrelated project kinds for 1000 references', count => {
  let reads = 0;
  const projectClassMembers: VbaProjectClassMembers[] = Array.from({ length: count }, (_, i) => ({
   name: 'Other' + i, moduleName: 'Other' + i, get kind() { reads++; return 'document' as const; }, members: [],
  }));
  projectClassMembers.push(form());
  const source = 'Option Explicit\n' + fixture.header + '\nDim result As Variant\n' + fixture.declaration + '\n'
   + Array.from({ length: 1000 }, () => 'result = ' + fixture.expression).join('\n') + '\nEnd Sub\n';
  expect(run(source, { ...context(), meProjectType: fixture.me, projectClassMembers })).toEqual([]);
  expect(reads).toBe(0);
 });
});
describe('real form control binding', () => {
 it.each(['T1.Nope', 't1.Nope', 'T2.Nope', 'Me.T1.Nope', 'f.T1.Nope', 'F1.T1.Nope'])('keeps the missing member diagnostic for %s', expression => {
  const source = 'Sub Main()\nDim f As F1\nDim result As Variant\nresult = ' + expression + '\nEnd Sub';
  const diagnostics = run(source, context());
  expect(diagnostics).toHaveLength(1);expect(diagnostics[0][0]).toBe('memberNotFound');
  expect(diagnostics[0][1]).toContain('Nope');
 });
 it.each(['Dim t1 As Object', 'If True Then\nDim t1 As Object\nEnd If'])('retains local control-name shadowing for %s', declaration => {
  expect(run('Sub Main()\n' + declaration + '\nDim result As Variant\nresult = T1.Nope\nEnd Sub', context())).toEqual([]);
 });
 it('retains parameter control-name shadowing', () => {
  expect(run('Sub Main(ByVal t1 As Object)\nDim result As Variant\nresult = T1.Nope\nEnd Sub', context())).toEqual([]);
 });
 it('does not apply local shadow or absent-current-form guards to a qualified control', () => {
  const source = 'Sub Main(ByVal T1 As Object)\nDim f As F1\nDim result As Variant\nresult = f.T1.Nope\nEnd Sub';
  expect(run(source, { ...context(), meProjectType: undefined })).toHaveLength(1);
 });
 it('retains valid public controls and case-insensitive form matching', () => {
  expect(run('Sub Main()\nDim result As Variant\nresult = T1.Text\nEnd Sub', context('f1'))).toEqual([]);
  expect(run('Sub Main()\nDim result As Variant\nresult = T1.Nope\nEnd Sub', context('f1'))).toHaveLength(1);
 });
 it('preserves an explicitly empty form name rather than treating it as absent', () => {
  const ctx = { ...context(''), projectClassMembers: [form('')] };
  expect(run('Sub Main()\nDim result As Variant\nresult = T1.Nope\nEnd Sub', ctx)).toHaveLength(1);
 });
 it('uses current control metadata on each invocation', () => {
  const ctx = context(), source = 'Sub Main()\nDim result As Variant\nresult = T1.Nope\nEnd Sub';
  expect(run(source, ctx)).toHaveLength(1);
  ctx.projectClassMembers![0].members[0].returns = 'MSForms.Frame';
  expect(run(source, ctx)).toEqual([]);
 });
});
