import { afterEach, describe, expect, it, vi } from 'vitest';
import * as arrays from '../src/analyzer/diagnostics/rules/arrays';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
type Hit = Parameters<Parameters<typeof checkAssignmentTypes>[6]>;
function prepare(source: string) {
 const mod = parseModule(source);
 return (vba7?: boolean): Hit[] => {
  const environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
  const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
  const hits: Hit[] = [];
  checkAssignmentTypes(source, mod, symbols, undefined, {}, activity, (...hit) => { hits.push(hit); });
  return hits;
 };
}
afterEach(() => { vi.restoreAllMocks(); });
describe('assignment module Option Base', () => {
 it.each(['', 'Option Base 0\n', 'Option Base 1\n'])('resolves the option once across repeated array reads and procedures (%s)', (option) => {
  const base = option.includes('1') ? 1 : 0;
  const proc = (name: string) => `Sub ${name}()\nDim v As Variant\nDim n As Long\nv = Array("bad", "1")\n` + `n = v(${base})\n`.repeat(100) + 'End Sub\n';
  const spy = vi.spyOn(arrays, 'moduleOptionBase');
  expect(prepare(option + proc('P') + proc('Q'))()).toHaveLength(200);
  expect(spy).toHaveBeenCalledTimes(1);
 });
 it('refreshes active options for reused parsed nodes between invocations', () => {
  const source = '#If VBA7 Then\nOption Base 1\n#Else\nOption Base 0\n#End If\nSub P()\nDim v As Variant\nDim n As Long\nv = Array("bad", "1")\nn = v(0)\nn = v(1)\nEnd Sub';
  const run = prepare(source);
  expect(source.slice(run(true)[0][2].start, run(true)[0][2].end)).toBe('v(1)');
  expect(source.slice(run(false)[0][2].start, run(false)[0][2].end)).toBe('v(0)');
  expect(run(true)).toHaveLength(1); expect(run(true)[0][1]).toContain('bad');
  expect(run(false)).toHaveLength(1); expect(run(false)[0][1]).toContain('bad');
  expect(run(true)).toHaveLength(1);
 });
 it('retains VBA.Array zero-base behavior alongside module-based Array', () => {
  const source = 'Option Base 1\nSub P()\nDim a As Variant\nDim b As Variant\nDim n As Long\na = Array("bad", "1")\nb = VBA.Array("bad", "1")\nn = a(1)\nn = b(0)\nn = b(1)\nEnd Sub';
  const hits = prepare(source)();
  expect(hits.map(h => source.slice(h[2].start, h[2].end))).toEqual(['a(1)', 'b(0)']);
  expect(hits).toHaveLength(2); expect(hits.every(h => h[1].includes('bad'))).toBe(true);
 });
 it('keeps the first active Base option', () => {
  expect(prepare('Option Base 1\nOption Base 0\nSub P()\nDim v As Variant\nDim n As Long\nv = Array("bad", "1")\nn = v(1)\nEnd Sub')()).toHaveLength(1);
 });
 it('does not resolve array options for ordinary scalar assignments', () => {
  const spy = vi.spyOn(arrays, 'moduleOptionBase');
  expect(prepare('Sub P()\nDim n As Long\n' + 'n = 1\n'.repeat(100) + 'End Sub')()).toEqual([]);
  expect(spy).not.toHaveBeenCalled();
 });
});
