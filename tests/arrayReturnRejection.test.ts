import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
function prepare(body: string) {
 const source = 'Function F() As Variant\nDim n As Long\n' + body + '\nEnd Function\nSub P()\nDim a() As Long\na = F()\nEnd Sub';
 const mod = parseModule(source);
 return (vba7?: boolean) => {
  const environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
  const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
  const hits: string[] = [];
  checkAssignmentTypes(source, mod, symbols, undefined, {}, activity, code => { hits.push(code); });
  return hits;
 };
}
describe('array return rejection', () => {
 it('keeps a scalar return rejected even when later returns are arrays', () => {
  expect(prepare('F = 1\n' + 'Debug.Print n\n'.repeat(1000) + 'F = Array(1)')()).toEqual([]);
 });
 it('keeps self-reads and rejected single-line conditions from being restored by array branches', () => {
  expect(prepare('Debug.Print F\nF = Array(1)')()).toEqual([]);
  expect(prepare('If F Then F = Array(1) Else F = Array(2)')()).toEqual([]);
 });
 it('still reports array-only and never-assigned Empty returns into a Long array', () => {
  expect(prepare('If True Then F = Array(1) Else F = Array(2)')()).toEqual(['assignmentTypeMismatch']);
  expect(prepare('Debug.Print n')()).toEqual(['assignmentTypeMismatch']);
 });
 it('keeps active conditional rejection separate on reused parsed nodes', () => {
  const run = prepare('#If VBA7 Then\nF = 1\n#Else\nF = Array(1)\n#End If\nDebug.Print n');
  expect(run(true)).toEqual([]);
  expect(run(false)).toEqual(['assignmentTypeMismatch']);
  expect(run(true)).toEqual([]);
 });
});
