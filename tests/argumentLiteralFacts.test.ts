import { afterEach, describe, expect, it, vi } from 'vitest';
import * as inference from '../src/analyzer/diagnostics/typeInference';
import { checkArgumentTypes } from '../src/analyzer/diagnostics/rules/argumentTypes';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
type Hit = Parameters<Parameters<typeof checkArgumentTypes>[5]>;
function prepare(source: string) {
 const mod = parseModule(source);
 return (vba7?: boolean) => {
  const environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
  const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
  const hits: Hit[] = [];
  const visitor = checkArgumentTypes(source, symbols, undefined, undefined, {}, (...hit) => { hits.push(hit); }, activity);
  for (const member of mod.members) {
   if (member.kind !== 'Procedure') continue;
   const visit = visitor(member);
   if (visit) forEachStatementWithHeaders(source, member.body, visit, activity);
  }
  return hits;
 };
}
afterEach(() => { vi.restoreAllMocks(); });
describe('argument literal fact preparation', () => {
 it('does not request literal facts when there are no calls', () => {
  const run = prepare('Sub P()\nDim v As Long\n' + 'v = 1\n'.repeat(1000) + 'End Sub');
  const spy = vi.spyOn(inference, 'knownLocalLiteralValuesAt');
  expect(run()).toEqual([]);
  expect(spy).not.toHaveBeenCalled();
 });
 it('does not request literal facts for a literal-only mismatch', () => {
  const run = prepare('Sub TakeLong(ByVal n As Long)\nEnd Sub\nSub P()\nTakeLong "bad"\nEnd Sub');
  const spy = vi.spyOn(inference, 'knownLocalLiteralValuesAt');
  expect(run()).toHaveLength(1);
  expect(spy).not.toHaveBeenCalled();
 });
 it('requests facts once for repeated known Variant values and preserves scalar-to-object mismatches', () => {
  const run = prepare('Sub TakeCollection(ByVal n As Collection)\nEnd Sub\nSub P()\nDim v As Variant\nv = 300\nTakeCollection v\nTakeCollection v\nEnd Sub');
  const spy = vi.spyOn(inference, 'knownLocalLiteralValuesAt');
  expect(run()).toHaveLength(2);
  expect(spy).toHaveBeenCalledTimes(1);
 });
 it('keeps activity-specific values on reused parsed nodes', () => {
  const run = prepare('Sub TakeCollection(ByVal n As Collection)\nEnd Sub\nSub P()\nDim v As Variant\n#If VBA7 Then\nv = 1\n#Else\nSet v = New Collection\n#End If\nTakeCollection v\nEnd Sub');
  expect(run(true)).toHaveLength(1);
  expect(run(false)).toHaveLength(0);
  expect(run(true)).toHaveLength(1);
 });
});
