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
describe('array-shape invalidation batches', () => {
 it('copies one shape map when a jump target clears every tracked array', () => {
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
  expect(entries).toBeLessThanOrEqual(count * (count - 1) / 2 + count);
 });
 it.each(['GoTo checkpoint\ncheckpoint:', 'GoSub handler\nhandler:'])('preserves earlier snapshots when clearing at %s', clear => {
  const { run } = fixture(`x = 1\n${clear}\nx = 2`);
  const snapshots = [...run().values()];
  expect(snapshots.length).toBeGreaterThan(0);
  expect([...snapshots[0].keys()]).toEqual(['a0', 'a1', 'a2']);
  expect(snapshots[0].get('a0')?.dims[0].upper).toBe(3);
 });
 it('preserves untouched names and branch-entry snapshots after forgetting several arrays', () => {
  const { source, run } = fixture('x = 1\nIf True Then\nErase a0, a1\nx = 2\nElse\nx = 3\nEnd If\nx = 4');
  const states = [...run()].map(([node, shapes]) => [source.slice(node.span.start, node.span.end).trim(), [...shapes.keys()]] as const);
  expect(states.find(([text]) => text === 'x = 2')?.[1]).toEqual(['a2']);
  expect(states.find(([text]) => text === 'x = 3')?.[1]).toEqual(['a0', 'a1', 'a2']);
  expect(states.find(([text]) => text === 'x = 4')?.[1]).toEqual(['a2']);
  expect(states.find(([text]) => text === 'x = 1')?.[1]).toEqual(['a0', 'a1', 'a2']);
 });
 it('leaves a snapshot shared when the invalidation names are untracked', () => {
  const { source, run } = fixture('x = 1\nErase missing\nx = 2');
  const result = run();
  const first = [...result].find(([node]) => source.slice(node.span.start, node.span.end).trim() === 'x = 1')?.[1];
  const last = [...result].find(([node]) => source.slice(node.span.start, node.span.end).trim() === 'x = 2')?.[1];
  expect(last).toBe(first);
 });
});
