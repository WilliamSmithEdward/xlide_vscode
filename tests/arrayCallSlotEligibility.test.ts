import { describe, expect, it } from 'vitest';
import { checkFixedArraySubscriptBounds } from '../src/analyzer/diagnostics/rules/arrays';
import { statementTokens } from '../src/analyzer/diagnostics/analysisContext';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
type Hit = Parameters<Parameters<typeof checkFixedArraySubscriptBounds>[4]>;
function fixture(source: string) {
 const mod = parseModule(source);
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const run = () => { const hits: Hit[] = []; checkFixedArraySubscriptBounds(source, mod, symbols, undefined, (...hit) => { hits.push(hit); }); return hits; };
 return { mod, run };
}
describe('fixed-array call slot eligibility', () => {
 it('does not repeatedly scan nested ordinary-call arguments without counter facts', () => {
  const depth = 200;
  const source = 'Sub P()\nDim a(3) As Long\nx = ' + 'F('.repeat(depth) + '1' + ')'.repeat(depth) + '\nEnd Sub';
  const { mod, run } = fixture(source);
  const proc = mod.members.find(member => member.kind === 'Procedure');
  if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
  const statement = proc.body.find(node => node.kind === 'Statement');
  if (!statement) throw new Error('Missing assignment');
  let reads = 0;
  for (const token of statementTokens(source, statement.span)) {
   const raw = token.rawText;
   Object.defineProperty(token, 'rawText', { get: () => { reads++; return raw; } });
  }
  expect(run()).toEqual([]);
  expect(reads).toBeLessThan(depth * 100);
 });
 it('continues into unrelated calls to check a nested declared array', () => {
  const source = 'Sub P()\nDim a(3) As Long\nx = ' + 'F('.repeat(100) + 'a(5)' + ')'.repeat(100) + '\nEnd Sub';
  const hits = fixture(source).run();
  expect(hits).toHaveLength(1);
  expect(hits[0][1]).toContain('upper bound 3');
 });
 it('keeps symbolic bounds checks when only loop counters prove an array shape', () => {
  const source = 'Sub P(b() As Long)\nDim a(3) As Long\nDim i As Long\nFor i = 0 To UBound(b) + 1\nx = F(b(i))\nNext\nEnd Sub';
  const hits = fixture(source).run();
  expect(hits).toHaveLength(1);
  expect(hits[0][1]).toContain('above its upper bound');
 });
 it('keeps wrong-dimension reports inside an unknown call', () => {
  const hits = fixture('Sub P()\nDim a(3, 3) As Long\nx = F(a(1))\nEnd Sub').run();
  expect(hits).toHaveLength(1);
  expect(hits[0][0]).toBe('wrongNumberOfDimensions');
 });
});
