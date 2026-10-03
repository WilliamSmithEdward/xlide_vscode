import { afterEach, describe, expect, it, vi } from 'vitest';
import { redimShapesAt } from '../src/analyzer/diagnostics/rules/arrays';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

function fixture(body: string, count = 3) {
 const source = 'Sub P()\n' + Array.from({ length: count }, (_, i) => `Dim a${i}() As Long\nReDim a${i}(3)`).join('\n') + '\n' + body + '\nEnd Sub';
 const mod = parseModule(source);
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const proc = mod.members.find(member => member.kind === 'Procedure');
 if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
 const run = () => redimShapesAt(source, symbols, proc, undefined, 0);
 return { source, run };
}
afterEach(() => { vi.restoreAllMocks(); });
describe('array-shape retained snapshots', () => {
 it('copies only retained maps while establishing and clearing many shapes', () => {
  const count = 100;
  const { run } = fixture('x = 1\nGoTo checkpoint\ncheckpoint:\nx = 2', count);
  const original = Map.prototype[Symbol.iterator];
  let entries = 0;
  vi.spyOn(Map.prototype, Symbol.iterator).mockImplementation(function* (this: Map<string, { origin?: string }>) {
   for (const entry of original.call(this)) {
    if (typeof entry[0] === 'string' && entry[0].startsWith('a') && entry[1]?.origin === 'ReDim') entries++;
    yield entry;
   }
  });
  const result = run();
  expect([...result.values()].map(map => map.size)).toEqual([count, count]);
  expect(entries).toBeLessThanOrEqual(count);
 });
 it('keeps every observed version immutable across later ReDims', () => {
  const { source, run } = fixture('x = 1\nReDim a0(7)\nx = 2\nReDim a0(n)\nx = 3');
  const states = [...run()].map(([node, shapes]) => [source.slice(node.span.start, node.span.end).trim(), shapes] as const);
  const first = states.find(([text]) => text === 'x = 1')?.[1];
  const second = states.find(([text]) => text === 'x = 2')?.[1];
  const third = states.find(([text]) => text === 'x = 3')?.[1];
  expect(first?.get('a0')?.dims[0].upper).toBe(3);
  expect(second?.get('a0')?.dims[0].upper).toBe(7);
  expect(third?.has('a0')).toBe(false);
  expect(second).not.toBe(first);
 });
 it('retains an empty branch-entry map before the first shape is created', () => {
  const { source, run } = fixture('Dim a0() As Long\nIf True Then\nReDim a0(3)\nx = 1\nElse\nx = 2\nEnd If\nx = 3', 0);
  const states = [...run()].map(([node, shapes]) => [source.slice(node.span.start, node.span.end).trim(), shapes] as const);
  expect(states.map(([text]) => text)).toEqual(['x = 1']);
  expect(states[0][1].get('a0')?.dims[0].upper).toBe(3);
 });
 it('continues sharing an observed map across statements that do not change shapes', () => {
  const { source, run } = fixture('x = 1\nx = 2');
  const states = [...run()].filter(([node]) => source.slice(node.span.start, node.span.end).trim().startsWith('x ='));
  expect(states).toHaveLength(2);
  expect(states[0][1]).toBe(states[1][1]);
 });
});
