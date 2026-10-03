import { describe, expect, it } from 'vitest';
import { checkObjectDefaultValues } from '../src/analyzer/diagnostics/rules/objectValues';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
type Hit = Parameters<Parameters<typeof checkObjectDefaultValues>[3]>;
function run(body: string) {
 const source = 'Sub P()\n' + body + '\nEnd Sub';
 const mod = parseModule(source);
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const hits: Hit[] = [];
 const factory = checkObjectDefaultValues(source, symbols, {}, (...hit) => { hits.push(hit); });
 for (const proc of mod.members) { if (proc.kind === 'Procedure') { const visit = factory(proc); if (visit) forEachStatementWithHeaders(source, proc.body, visit); } }
 return { source, hits };
}
describe('object read token scans', () => {
 it('reports a whole Let value once', () => {
  const { source, hits } = run('Dim c As New Collection\nDim v As Variant\nv = c');
  expect(hits).toHaveLength(1);
  expect(hits[0][1]).toContain("'450'");
  expect(source.slice(hits[0][2].start, hits[0][2].end)).toBe('c');
 });
 it('preserves whole Print items before operator reads in diagnostic order', () => {
  const { source, hits } = run('Dim ws As Worksheet\nDebug.Print ws; ws + 1; ws');
  const at = source.indexOf('Debug.Print');
  const first = source.indexOf('ws', at), middle = source.indexOf('ws', first + 2), last = source.indexOf('ws', middle + 2);
  expect(hits.map(hit => hit[2].start)).toEqual([first, last, middle]);
  expect(hits.every(hit => hit[0] === 'objectDefaultValue')).toBe(true);
 });
 it('keeps repeated operator operands as separate reads', () => {
  const { source, hits } = run('Dim ws As Worksheet\nDebug.Print ws + ws + ws');
  expect(hits).toHaveLength(3);
  expect(new Set(hits.map(hit => hit[2].start)).size).toBe(3);
  expect(hits.every(hit => source.slice(hit[2].start, hit[2].end) === 'ws')).toBe(true);
 });
 it('excludes indexed late-bound values while retaining whole builtin arguments and typed indexing', () => {
  expect(run('Dim x As Object\nDim v As Variant\nSet x = New Collection\nv = x(1)').hits).toEqual([]);
  const whole = run('Dim x As Object\nSet x = New Collection\nDebug.Print CStr(x)').hits;
  expect(whole).toHaveLength(1);
  expect(whole[0][1]).toContain("'450'");
  const typed = run('Dim ws As Worksheet\nDim v As Variant\nv = ws(1)').hits;
  expect(typed).toHaveLength(1);
  expect(typed[0][1]).toContain("'438'");
 });
});
