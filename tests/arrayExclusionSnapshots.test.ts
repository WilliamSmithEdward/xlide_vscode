import { afterEach, describe, expect, it, vi } from 'vitest';
import { checkFixedArraySubscriptBounds } from '../src/analyzer/diagnostics/rules/arrays';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
type Hit = Parameters<Parameters<typeof checkFixedArraySubscriptBounds>[4]>;
function fixture(body: string, count = 2) {
 const source = 'Sub P()\n' + Array.from({ length: count }, (_, i) => `Dim a${i}() As Long\nReDim a${i}(3)`).join('\n') + '\n' + body + '\nEnd Sub';
 const mod = parseModule(source);
 const run = (vba7?: boolean) => {
  const environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment });
  const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
  const hits: Hit[] = [];
  checkFixedArraySubscriptBounds(source, mod, symbols, activity, (...hit) => { hits.push(hit); });
  return hits;
 };
 return { run, source };
}
afterEach(() => { vi.restoreAllMocks(); });
describe('fixed-array exclusion snapshots', () => {
 it('does not rescan ReDim targets for every statement sharing a shape map', () => {
  const { run } = fixture('x = a0(0)\n'.repeat(200) + 'x = a0(7)', 100);
  const original = Set.prototype[Symbol.iterator];
  let reads = 0;
  vi.spyOn(Set.prototype, Symbol.iterator).mockImplementation(function* (this: Set<string>) {
   for (const value of original.call(this)) {
    if (this.size === 100 && value === 'a0') reads++;
    yield value;
   }
  });
  expect(run()).toHaveLength(1);
  expect(reads).toBeLessThan(40);
 });
 it('uses new exclusions after an erasure and a later ReDim', () => {
  const { run, source } = fixture('x = a0(7)\nErase a0\nx = a0(7)\nx = a1(7)\nReDim a0(1)\nx = a0(7)');
  const hits = run();
  expect(hits).toHaveLength(3);
  expect(hits.map(hit => source.slice(hit[2].start, hit[2].end))).toEqual(['7', '7', '7']);
  expect(hits[2][1]).toContain('upper bound 1');
 });
 it('keeps conditional activity separate when reusing parsed nodes', () => {
  const { run } = fixture('#If VBA7 Then\nErase a0\n#End If\nx = a0(7)\nx = a1(7)');
  expect(run(false)).toHaveLength(2);
  expect(run(true)).toHaveLength(1);
  expect(run(false)).toHaveLength(2);
 });
 it('restores branch-entry exclusions before checking another arm', () => {
  const { run } = fixture('If flag Then\nErase a0\nx = a0(7)\nElse\nx = a0(7)\nEnd If\nx = a1(7)');
  expect(run()).toHaveLength(2);
 });
});
