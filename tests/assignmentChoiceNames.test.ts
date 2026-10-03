import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
type Hit = Parameters<Parameters<typeof checkAssignmentTypes>[6]>;
function prepare(source: string, vba7?: boolean) {
 const mod = parseModule(source), environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
 const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
 return { symbols, run() { const hits: Hit[] = []; checkAssignmentTypes(source, mod, symbols, undefined, {}, activity, (...hit) => { hits.push(hit); }); return hits; } };
}
const expressions = ['Choose(0, 1)', 'Switch(False, 1)', 'IIf(False, 1, Null)'];
describe('assignment Null-choice module names', () => {
 it.each(expressions)('bounds repeated module-name reads across procedures (%s)', (expression) => {
  const declarations = Array.from({ length: 100 }, (_, i) => 'Dim k' + i + ' As Long\n').join('');
  const body = 'Dim n As Long\n' + ('n = ' + expression + '\n').repeat(100);
  const { symbols, run } = prepare(declarations + 'Sub P()\n' + body + 'End Sub\nSub Q()\n' + body + 'End Sub');
  let reads = 0;
  for (const child of symbols.root.children ?? []) { const name = child.name; Object.defineProperty(child, 'name', { enumerable: true, get() { reads++; return name; } }); }
  expect(run()).toHaveLength(200);
  expect(reads).toBeLessThan(1500);
 });
 it('preserves module shadows regardless of case or declaration kind and bypasses them for VBA calls', () => {
  for (const declaration of ['Dim ChOoSe As Variant', 'Const ChOoSe = 1', 'Function ChOoSe() As Variant\nEnd Function', 'Sub ChOoSe()\nEnd Sub']) {
   const source = declaration + '\nSub P()\nDim n As Long\nn = Choose(0, 1)\nn = VBA.Choose(0, 1)\nEnd Sub';
   const hits = prepare(source).run();
   expect(hits).toHaveLength(1);
   expect(source.slice(hits[0][2].start, hits[0][2].end)).toBe('VBA.Choose(0, 1)');
  }
 });
 it('does not substitute local/runtime shadow membership for direct module membership', () => {
  const hits = prepare('Sub P()\nDim Choose As Variant\nDim n As Long\nn = Choose(0, 1)\nEnd Sub').run();
  expect(hits).toHaveLength(1);
  expect(hits[0][1]).toContain('names no choice');
 });
 it('refreshes queries on each invocation even when the bound root is reused', () => {
  const { symbols, run } = prepare('Dim k As Long\nSub P()\nDim n As Long\nn = Choose(0, 1)\nEnd Sub');
  expect(run()).toHaveLength(1);
  const child = symbols.root.children!.find(symbol => symbol.name === 'k')!;
  child.name = 'ChOoSe';
  expect(run()).toEqual([]);
  child.name = 'k';
  expect(run()).toHaveLength(1);
 });
 it('refreshes conditional module shadows while reusing parsed source', () => {
  const source = '#If VBA7 Then\nDim ChOoSe As Variant\n#Else\nDim k As Variant\n#End If\nSub P()\nDim n As Long\nn = Choose(0, 1)\nEnd Sub';
  expect(prepare(source, true).run()).toEqual([]);
  expect(prepare(source, false).run()).toHaveLength(1);
  expect(prepare(source, true).run()).toEqual([]);
 });
});
