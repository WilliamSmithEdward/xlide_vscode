import { describe, expect, it } from 'vitest';
import { classifyReferenceKinds } from '../src/analyzer/references/referenceKinds';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';

function counted(source: string) {
 let reads = 0;
 const tokens = tokenizeCached(source);
 for (const token of tokens) {
  const raw = token.rawText;
  Object.defineProperty(token, 'rawText', { get() { reads++; return raw; } });
  Object.freeze(token);
 }
 Object.freeze(tokens);
 return { tokens, reads: () => reads };
}
describe.each(['', 'Set ', 'Let ', 'LSet ', 'RSet '])('reference target suffix work (%s)', prefix => {
 it.each(['terminal', 'all'])('bounds reads for %s queried names in a deep target', mode => {
  const source = prefix + 'a(' + 'f('.repeat(1000) + '1' + ')'.repeat(1000) + ').Value = 1' + " ' " + mode;
  const c = counted(source), value = source.indexOf('Value');
  const offsets = mode === 'terminal' ? [value] : c.tokens.filter(t => t.kind === 'identifier').map(t => t.start);
  const result = classifyReferenceKinds(source, offsets);
  expect(result.get(value)).toBe('write');
  expect([...result].every(([offset, kind]) => kind === (offset === value ? 'write' : 'read'))).toBe(true);
  expect(result.size).toBe(offsets.length);
  expect(c.reads()).toBeLessThan(20 * c.tokens.length);
 });
});
it('retains the single-suffix cold scan without a parenthesis index', () => {
 const source = 'a(i) = 1', c = counted(source);
 expect([...classifyReferenceKinds(source, [0])]).toEqual([[0, 'write']]);
 expect(c.reads()).toBeLessThan(25);
});
it.each([
 ['a) f(=1)', ['read', 'read']],
 ['a) f(=1', ['read', 'write']],
 ['a(i)(j) = 1', ['write', 'read', 'read']],
 ['a(f(1)).Value = 1', ['read', 'read', 'write']],
 ['a(f(1)) = 1', ['write', 'read']],
] as const)('preserves suffix fallback and target kinds for %s', (source, kinds) => {
 const tokens = tokenizeCached(source), names = tokens.filter(t => t.kind === 'identifier');
 const result = classifyReferenceKinds(source, names.map(t => t.start));
 expect(names.map(t => result.get(t.start))).toEqual(kinds);
});
it('preserves complete Map ordering and duplicate/default offsets', () => {
 const source = 'a(f(1)).Value = 1', value = source.indexOf('Value'), f = source.indexOf('f');
 expect([...classifyReferenceKinds(source, [f, value, f, -1, 999])]).toEqual([[value, 'write'], [f, 'read'], [-1, 'read'], [999, 'read']]);
});
