import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
type Hit = Parameters<Parameters<typeof checkAssignmentTypes>[6]>;
function run(source: string, vba7?: boolean, watch?: (symbols: ReturnType<typeof buildModuleSymbols>) => void): Hit[] {
 const mod = parseModule(source), environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
 const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
 watch?.(symbols);
 const hits: Hit[] = [];
 checkAssignmentTypes(source, mod, symbols, undefined, {}, activity, (...hit) => { hits.push(hit); });
 return hits;
}
describe('assignment direct local facts', () => {
 it.each(['null', 'empty', 'fixed', 'known'])('bounds repeated direct declaration name reads (%s)', (mode) => {
  const locals = Array.from({ length: 100 }, (_, i) => `Dim k${i} As Long\n`).join('');
  const declaration = mode === 'fixed' ? 'Dim v As String * 3\n' : mode === 'known' ? 'Dim v As String\n' : 'Dim v As Variant\n';
  const initial = mode === 'null' ? 'v = Null\n' : mode === 'known' ? 'v = "bad"\n' : '';
  let reads = 0;
  const hits = run('Sub P()\n' + locals + declaration + 'Dim n As Long\n' + initial + 'n = v\n'.repeat(100) + 'End Sub', undefined, (symbols) => {
   for (const symbol of symbols.all) { const name = symbol.name; Object.defineProperty(symbol, 'name', { enumerable: true, get() { reads++; return name; } }); }
  });
  expect(hits).toHaveLength(mode === 'null' || mode === 'known' ? 100 : 0);
  expect(reads).toBeLessThan(3000);
 });
 it('caches missing direct locals without treating module declarations as locals', () => {
  let reads = 0;
  const locals = Array.from({ length: 100 }, (_, i) => `Dim k${i} As Long\n`).join('');
  const source = 'Dim v As Variant\nSub P()\n' + locals + 'Dim n As Long\n' + 'n = v\n'.repeat(100) + 'End Sub';
  expect(run(source, undefined, (symbols) => {
   for (const symbol of symbols.all) { const name = symbol.name; Object.defineProperty(symbol, 'name', { enumerable: true, get() { reads++; return name; } }); }
  })).toEqual([]);
  expect(reads).toBeLessThan(3000);
 });
 it('keeps held values specific to the statement', () => {
  const hits = run('Sub P()\nDim v As Variant\nDim n As Long\nv = Null\nn = v\nv = 5\nn = v\nv = "bad"\nn = v\nEnd Sub');
  expect(hits.map(h => h[0])).toEqual(['assignmentTypeMismatch', 'assignmentTypeMismatch']);
  expect(hits[0][1]).toContain('Null'); expect(hits[1][1]).toContain('bad');
 });
 it('keeps Static and parameter guards and first matching child behavior', () => {
  expect(run('Sub P(v As Variant)\nDim n As Long\nv = Null\nn = v\nEnd Sub')).toEqual([]);
  expect(run('Sub P()\nStatic v As Variant\nDim v As Variant\nDim n As Long\nv = Null\nn = v\nEnd Sub')).toEqual([]);
  expect(run('Sub P()\nDim v As Variant\nStatic v As Variant\nDim n As Long\nv = Null\nn = v\nEnd Sub').map(h => h[0])).toEqual(['assignmentTypeMismatch']);
 });
 it('isolates the same local name in separate procedures', () => {
  const hits = run('Sub P()\nDim v As Variant\nDim n As Long\nv = Null\nn = v\nEnd Sub\nSub Q()\nDim v As Long\nDim n As Long\nv = 1\nn = v\nEnd Sub');
  expect(hits.map(h => h[0])).toEqual(['assignmentTypeMismatch']);
 });
 it('refreshes declaration choices across conditional activity', () => {
  const source = 'Sub P()\n#If VBA7 Then\nStatic v As Variant\n#Else\nDim v As Variant\n#End If\nDim n As Long\nv = Null\nn = v\nEnd Sub';
  expect(run(source, true)).toEqual([]);
  expect(run(source, false).map(h => h[0])).toEqual(['assignmentTypeMismatch']);
  expect(run(source, true)).toEqual([]);
 });
 it('retains array and fixed-string value diagnostics', () => {
  expect(run('Sub P()\nDim v As Variant\nDim a() As Long\nv = Array(1, 2)\na = v\nEnd Sub').map(h => h[0])).toEqual(['assignmentTypeMismatch']);
  expect(run('Sub P()\nDim v As String * 3\nDim n As Long\nn = v\nEnd Sub').map(h => h[0])).toEqual(['assignmentTypeMismatch']);
 });
});
