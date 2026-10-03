import { afterEach, describe, expect, it, vi } from 'vitest';
import * as heldFacts from '../src/analyzer/diagnostics/heldObjects';
import { checkArgumentTypes } from '../src/analyzer/diagnostics/rules/argumentTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
type Hit = Parameters<Parameters<typeof checkArgumentTypes>[5]>;
function fixture(source: string) {
 const mod = parseModule(source);
 return (vba7?: boolean) => {
  const environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
  const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
  const hits: Hit[] = [];
  const factory = checkArgumentTypes(source, symbols, undefined, undefined, {}, (...hit) => { hits.push(hit); }, activity);
  for (const member of mod.members) {
   if (member.kind === 'Procedure') {
    const visit = factory(member);
    if (visit) forEachStatementWithHeaders(source, member.body, visit, activity);
   }
  }
  return hits;
 };
}
afterEach(() => { vi.restoreAllMocks(); });
describe('argument held-object preparation', () => {
 it('does no held-object walk when there are no calls', () => {
  const run = fixture('Sub P()\nDim c As New Collection\n' + 'x = 1\n'.repeat(1000) + 'End Sub');
  const spy = vi.spyOn(heldFacts, 'heldObjectsAt');
  expect(run()).toEqual([]);
  expect(spy).not.toHaveBeenCalled();
 });
 it('does no held-object walk for literal-only arguments', () => {
  const run = fixture('Sub TakeLong(ByVal n As Long)\nEnd Sub\nSub P()\nTakeLong 1\nTakeLong "bad"\nEnd Sub');
  const spy = vi.spyOn(heldFacts, 'heldObjectsAt');
  const hits = run();
  expect(hits).toHaveLength(1);
  expect(spy).not.toHaveBeenCalled();
 });
 it('builds object facts once for repeated named arguments', () => {
  const run = fixture('Sub TakeCollection(ByVal o As Collection)\nEnd Sub\nSub P()\nDim o As Object\nSet o = New Collection\n' + 'TakeCollection o\n'.repeat(100) + 'End Sub');
  const spy = vi.spyOn(heldFacts, 'heldObjectsAt');
  expect(run()).toEqual([]);
  expect(spy).toHaveBeenCalledTimes(1);
 });
 it('keeps the Variant holding a scalar mismatch', () => {
  const run = fixture('Sub TakeCollection(ByVal o As Collection)\nEnd Sub\nSub P()\nDim v As Variant\nv = 1\nTakeCollection v\nEnd Sub');
  expect(run()).toHaveLength(1);
 });
 it('keeps conditional activity separate when parsed nodes are reused', () => {
  const run = fixture('Sub TakeCollection(ByVal o As Collection)\nEnd Sub\nSub P()\nDim v As Variant\n#If VBA7 Then\nv = 1\n#Else\nSet v = New Collection\n#End If\nTakeCollection v\nEnd Sub');
  expect(run(true)).toHaveLength(1);
  expect(run(false)).toHaveLength(0);
  expect(run(true)).toHaveLength(1);
 });
});
