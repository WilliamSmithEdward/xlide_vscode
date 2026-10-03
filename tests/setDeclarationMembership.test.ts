import { describe, expect, it } from 'vitest';
import { checkSetAssignments } from '../src/analyzer/diagnostics/rules/assignments';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import type { VbaProjectClassMembers, VbaSymbol } from '../src/analyzer/symbols/symbolModel';
import type { MemberCompletionContext } from '../src/analyzer/completion/memberAccess';
type Hit = Parameters<Parameters<typeof checkSetAssignments>[4]>;
const form: VbaProjectClassMembers = { name: 'Form1', kind: 'userform', moduleName: 'Form1', members: [{ name: 'Answer', kind: 'property', moduleName: 'Form1', returns: 'MSForms.TextBox' }] };
const ctx: MemberCompletionContext = { meProjectType: 'Form1', projectClassMembers: [form] };
function run(source: string, context: MemberCompletionContext = {}, project?: VbaSymbol[], vba7?: boolean, countNames?: (symbols: ReturnType<typeof buildModuleSymbols>) => void): Hit[] {
 const mod = parseModule(source), environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
 countNames?.(symbols);
 const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
 const hits: Hit[] = [], visitor = checkSetAssignments(source, symbols, project, context, (...hit) => { hits.push(hit); }, activity);
 for (const proc of mod.members) if (proc.kind === 'Procedure') { const visit = visitor(proc); if (visit) forEachStatementWithHeaders(source, proc.body, visit, activity); }
 return hits;
}
describe('Set declaration membership', () => {
 it.each([false, true])('bounds declaration reads across repeated Set statements (form=%s)', (inForm) => {
  const source = 'Sub P()\n' + Array.from({ length: 100 }, (_, i) => `Dim k${i} As Long\n`).join('') + 'Dim target As Long\n' + 'Set target = Nothing\n'.repeat(100) + 'End Sub';
  let reads = 0;
  const hits = run(source, inForm ? ctx : {}, undefined, undefined, (symbols) => {
   for (const symbol of symbols.all) { const name = symbol.name; Object.defineProperty(symbol, 'name', { enumerable: true, get() { reads++; return name; } }); }
  });
  expect(hits).toHaveLength(100);
  expect(reads).toBeLessThan(2000);
 });
 it('preserves module/local/parameter shadowing and isolates procedure locals', () => {
  expect(run('Dim Answer As Long\nSub P()\nSet aNsWeR = Nothing\nEnd Sub', ctx)[0][1]).toContain('declared as Long');
  const hits = run('Sub P(Answer As Long)\nSet Answer = Nothing\nEnd Sub\nSub Q()\nSet Answer = Nothing\nEnd Sub', ctx);
  expect(hits.map(h => h[1].includes('control on this form'))).toEqual([false, true]);
 });
 it('keeps project-visible symbols and enum members outside direct declaration shadowing', () => {
  const project = buildModuleSymbols('Other', 'standard', 'Public Answer As Long').root.children;
  expect(run('Sub P()\nSet Answer = Nothing\nEnd Sub', ctx, project)[0][1]).toContain('control on this form');
  expect(run('Enum E\nAnswer = 1\nEnd Enum\nSub P()\nSet Answer = Nothing\nEnd Sub', ctx)[0][1]).toContain('control on this form');
 });
 it('refreshes membership for conditional declarations between invocations', () => {
  const source = '#If VBA7 Then\nDim Answer As Long\n#End If\nSub P()\nSet Answer = Nothing\nEnd Sub';
  expect(run(source, ctx, undefined, true)[0][1]).toContain('declared as Long');
  expect(run(source, ctx, undefined, false)[0][1]).toContain('control on this form');
  expect(run(source, ctx, undefined, true)[0][1]).toContain('declared as Long');
 });
 it('still validates ordinary scalar/array/control and no-form assignments', () => {
  expect(run('Sub P()\nDim n As Long\nDim a(0) As Long\nSet n = Nothing\nSet a(0) = Nothing\nSet Answer = Nothing\nEnd Sub').map(h => h[0])).toEqual(['setRequiresObject']);
  expect(run('Sub P()\nSet Answer = Nothing\nEnd Sub', ctx)[0][1]).toContain('control on this form');
 });
});
