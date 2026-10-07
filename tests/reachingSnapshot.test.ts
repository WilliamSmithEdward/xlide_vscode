import { describe, expect, it } from 'vitest';
import { shareImmutableReachingStart, withReachingValue } from '../src/analyzer/diagnostics/reachingSnapshot';
import { rawExpressionTokens } from '../src/analyzer/diagnostics/walker';
import type { ReachingAssignments } from '../src/analyzer/diagnostics/straightLineValues';
const value = (n: number) => rawExpressionTokens(String(n));
function initial(size: number) { return new Map(Array.from({ length: size }, (_, i) => ['k' + i, value(i)] as const)); }
function compare(actual: ReachingAssignments, expected: ReachingAssignments) {
 expect(actual.size).toBe(expected.size); expect([...actual]).toEqual([...expected]);
 expect([...actual.keys()]).toEqual([...expected.keys()]); expect([...actual.values()]).toEqual([...expected.values()]);
 for (const [key, tokens] of expected) { expect(actual.has(key)).toBe(true); expect(actual.get(key)).toBe(tokens); }
 expect(actual.has('absent')).toBe(false); expect(actual.get('absent')).toBeUndefined();
}
describe('bounded reaching snapshots', () => {
 it('shares analyzer-owned immutable starts without enumerating their constants on writes', () => {
  let entries = 0;
  class CountedMap extends Map<string, readonly ReturnType<typeof value>[number][]> {
   override *[Symbol.iterator](): MapIterator<[string, readonly ReturnType<typeof value>[number][]]> { for (const entry of super[Symbol.iterator]()) { entries++; yield entry; } }
  }
  const base = new CountedMap(initial(1000)); entries = 0;
  let state = shareImmutableReachingStart(base);
  const first = withReachingValue(state, 'n', value(1)); state = first;
  for (let i = 2; i < 100; i++) state = withReachingValue(state, 'n', value(i));
  expect(entries).toBe(0);
  expect(base.has('n')).toBe(false);
  expect(first.get('n')?.[0].rawText).toBe('1');
  expect(state.get('n')?.[0].rawText).toBe('99');
  expect(state.get('k999')?.[0].rawText).toBe('999');
 });
 it.each([0, 63, 64, 1000])('preserves native lookup and ordered iteration through updates/compaction (base=%s)', (size) => {
  let actual: ReachingAssignments = initial(size), expected = new Map(actual);
  const snapshots: [ReachingAssignments, Map<string, readonly ReturnType<typeof value>[number][]>][] = [];
  for (let i = 0; i < 150; i++) {
   if (i % 20 === 0) snapshots.push([actual, new Map(expected)]);
   const key = i % 3 === 0 ? 'k0' : 'new' + i, tokens = value(i + 5000);
   actual = withReachingValue(actual, key, tokens); expected.set(key, tokens);
   if (i % 20 === 0 || i === 149) compare(actual, expected);
  }
  for (const [kept, expectedThen] of snapshots) compare(kept, expectedThen);
 });
 it('detaches a shared base from later mutation of the initial map', () => {
  const base = initial(100), old = withReachingValue(base, 'n', value(1));
  base.clear(); expect(old.size).toBe(101); expect(old.get('k0')?.[0].rawText).toBe('0');
 });
 it('preserves deletion/reinsertion order when invalidation materializes a state', () => {
  const base = initial(100), old = withReachingValue(base, 'k3', value(300));
  const invalidated = new Map(old); invalidated.delete('k3'); invalidated.delete('k4');
  const next = withReachingValue(invalidated, 'k3', value(400));
  expect([...next.keys()].at(-1)).toBe('k3'); expect(next.has('k4')).toBe(false);
  expect(old.get('k3')?.[0].rawText).toBe('300'); expect(old.has('k4')).toBe(true);
 });
 it('supports Map forEach callback arguments and receiver binding', () => {
  const state = withReachingValue(initial(100), 'n', value(1)), receiver = {};
  const entries: [string, readonly ReturnType<typeof value>[number][]][] = [];
  state.forEach(function (this: object, tokens, key, map) { expect(this).toBe(receiver); expect(map).toBe(state); entries.push([key, tokens]); }, receiver);
  expect(entries).toEqual([...state]);
 });
 it('reads the large original base only once for repeated updates of one local', () => {
  let entries = 0;
  class CountedMap extends Map<string, readonly ReturnType<typeof value>[number][]> {
   override *[Symbol.iterator](): MapIterator<[string, readonly ReturnType<typeof value>[number][]]> { for (const entry of super[Symbol.iterator]()) { entries++; yield entry; } }
  }
  let state: ReachingAssignments = new CountedMap(initial(1000)); entries = 0;
  for (let i = 0; i < 1000; i++) state = withReachingValue(state, 'n', value(i));
  expect(entries).toBe(1000); expect(state.get('n')?.[0].rawText).toBe('999');
 });
});
