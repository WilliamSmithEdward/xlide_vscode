import { afterEach, describe, expect, it, vi } from 'vitest';
import * as inference from '../src/analyzer/diagnostics/typeInference';
import { checkObjectDefaultValues } from '../src/analyzer/diagnostics/rules/objectValues';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { hostObjectModelForToken } from '../src/analyzer/host/hostRegistry';
type Hit = Parameters<Parameters<typeof checkObjectDefaultValues>[3]>;
function run(source: string, host = 'excel'): Hit[] {
 const mod = parseModule(source);
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const hits: Hit[] = [];
 const factory = checkObjectDefaultValues(source, symbols, { model: hostObjectModelForToken(host) }, (...hit) => { hits.push(hit); });
 for (const proc of mod.members) {
  if (proc.kind !== 'Procedure') continue;
  const visitor = factory(proc);
  if (visitor) forEachStatementWithHeaders(source, proc.body, visitor);
 }
 return hits;
}
afterEach(() => { vi.restoreAllMocks(); });
describe('object read indexes', () => {
 it('does not normalize every host type again for each unrelated statement', () => {
  const spy = vi.spyOn(inference, 'normalizeType');
  const source = 'Sub P()\nDim n As Long\n' + Array.from({ length: 100 }, (_, i) => 'Dim x' + i + ' As Application').join('\n') + '\n' + 'n = 1\n'.repeat(1000) + 'End Sub';
  expect(run(source)).toEqual([]);
  expect(spy.mock.calls.length).toBeLessThan(500);
 });
 it('keeps local host types separate across procedures that shadow a module variable', () => {
  const source = 'Dim x As Long\nSub A()\nDim x As Application\nSet x = Application\nDebug.Print X + 1\nEnd Sub\nSub B()\nDim x As Long\nx = 2\nDebug.Print x + 1\nEnd Sub';
  const hits = run(source);
  expect(hits).toHaveLength(1);
  expect(hits[0][0]).toBe('assignmentTypeMismatch');
  expect(source.slice(hits[0][2].start, hits[0][2].end)).toBe('X');
 });
 it('retains late-bound collection reads and Word document reads in single-line branches', () => {
  const collection = run('Sub P()\nDim x As Object\nSet x = New Collection\nDebug.Print X + 1\nEnd Sub');
  expect(collection.some(hit => hit[0] === 'objectDefaultValue' && hit[1].includes("'450'"))).toBe(true);
  const document = run('Sub P()\nDim x As Document\nSet x = ActiveDocument\nIf X Then Debug.Print 1 Else Debug.Print 2\nEnd Sub', 'word');
  expect(document).toHaveLength(1);
  expect(document[0][0]).toBe('assignmentTypeMismatch');
 });
});
