import { afterEach, describe, expect, it, vi } from 'vitest';
import { heldObjectsAt, HELD_VALUE } from '../src/analyzer/diagnostics/heldObjects';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { forEachStatement } from '../src/analyzer/diagnostics/walker';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
function fixture(body: string) {
 const source = 'Sub P()\n' + body + '\nEnd Sub';
 const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const proc = mod.members.find(member => member.kind === 'Procedure');
 if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
 const run = (vba7?: boolean) => {
  const activity = vba7 === undefined ? undefined : createConditionalActivityTracker(mod, { compilerConstants: { VBA7: vba7 } });
  const at = heldObjectsAt(source, proc, symbols, activity);
  const nodes = new Set(proc.body);
  forEachStatement(proc.body, node => { nodes.add(node); });
  return [...nodes].map(node => ({ text: source.slice(node.span.start, node.span.end).trim(), value: at(node) }));
 };
 return { run };
}
afterEach(() => { vi.restoreAllMocks(); });
describe('held-object unchanged snapshots', () => {
 it('copies a wide unchanged object map only once', () => {
  const { run } = fixture(Array.from({ length: 100 }, (_, i) => `Dim o${i} As New C`).join('\n') + '\n' + 'x = 1\n'.repeat(200));
  const original = Map.prototype[Symbol.iterator];
  let copied = 0;
  vi.spyOn(Map.prototype, Symbol.iterator).mockImplementation(function* (this: Map<string, string>) {
   for (const entry of original.call(this)) {
    if (typeof entry[0] === 'string' && /^o\d+$/.test(entry[0]) && entry[1] === 'C') copied++;
    yield entry;
   }
  });
  const reads = run().filter(row => row.text === 'x = 1');
  expect(reads).toHaveLength(200);
  expect(reads[0].value.classes.size).toBe(100);
  expect(reads[199].value).toBe(reads[0].value);
  expect(copied).toBeLessThan(500);
 });
 it('preserves collection item snapshots across object and literal additions', () => {
  const { run } = fixture('Dim c As New Collection\nc.Add New A\nx = c.Count\nc.Add 1\nx = c.Count\nc.Add New B, , Before:=1\nx = c.Count');
  const reads = run().filter(row => row.text === 'x = c.Count');
  expect(reads.map(row => row.value.items.get('c'))).toEqual([['A'], ['A', HELD_VALUE], ['B', 'A', HELD_VALUE]]);
  expect(reads[0].value.items.get('c')).toEqual(['A']);
 });
 it('preserves branch-entry items and historical snapshots after restoring', () => {
  const { run } = fixture('Dim c As New Collection\nc.Add New A\nx = 1\nIf flag Then\nc.Add New B\nx = 2\nElse\nx = 3\nEnd If\nx = 4');
  const rows = run();
  expect(rows.find(row => row.text === 'x = 1')?.value.items.get('c')).toEqual(['A']);
  expect(rows.find(row => row.text === 'x = 2')?.value.items.get('c')).toEqual(['A', 'B']);
  expect(rows.find(row => row.text === 'x = 3')?.value.items.get('c')).toEqual(['A']);
  // A block that names c conservatively ends its item facts on exit.
  expect(rows.find(row => row.text === 'x = 4')?.value.items.get('c')).toBeUndefined();
  const block = rows.find(row => row.text.startsWith('If flag'));
  expect(block?.value.items.get('c')).toEqual(['A']);
 });
 it('invalidates cached facts when aliasing a collection or calling a changing member', () => {
  const { run } = fixture('Dim c As New Collection\nDim d As Object\nc.Add New A\nx = 1\nSet d = c\nx = 2\nc.Remove 1\nx = 3');
  const rows = run();
  expect(rows.find(row => row.text === 'x = 1')?.value.items.get('c')).toEqual(['A']);
  expect(rows.find(row => row.text === 'x = 2')?.value.items.get('c')).toBeUndefined();
  expect(rows.find(row => row.text === 'x = 2')?.value.classes.get('d')).toBe('Collection');
  expect(rows.find(row => row.text === 'x = 3')?.value.items.get('c')).toBeUndefined();
 });
 it('keeps activity environments independent on a reused parsed procedure', () => {
  const { run } = fixture('Dim o As New A\n#If VBA7 Then\nSet o = New B\n#End If\nx = 1');
  expect(run(false).find(row => row.text === 'x = 1')?.value.classes.get('o')).toBe('A');
  expect(run(true).find(row => row.text === 'x = 1')?.value.classes.get('o')).toBe('B');
  expect(run(false).find(row => row.text === 'x = 1')?.value.classes.get('o')).toBe('A');
 });
 it('invalidates item facts after a changing member and replaces them on a new collection', () => {
  const { run } = fixture('Dim c As New Collection\nc.Add New A\nx = 1\nc.Remove 1\nx = 2\nSet c = New Collection\nx = 3');
  const rows = run();
  expect(rows.find(row => row.text === 'x = 1')?.value.items.get('c')).toEqual(['A']);
  expect(rows.find(row => row.text === 'x = 2')?.value.items.get('c')).toBeUndefined();
  expect(rows.find(row => row.text === 'x = 2')?.value.classes.get('c')).toBe('Collection');
  expect(rows.find(row => row.text === 'x = 3')?.value.items.get('c')).toEqual([]);
 });
});
