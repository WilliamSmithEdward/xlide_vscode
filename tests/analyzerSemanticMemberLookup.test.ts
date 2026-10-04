import { describe, expect, it } from 'vitest';
import { collectHostMemberMethodTokens } from '../src/analyzer/semantic/typeSemanticTokens';
import { resolveHostMemberKindAt } from '../src/analyzer/completion/memberAccess';
import type { HostMember, HostObjectModel } from '../src/analyzer/host/excelObjectModel';

function fixture(count: number) {
 let reads = 0;
 const members: HostMember[] = Array.from({ length: count }, (_, i) => Object.freeze({
  get name() { reads++; return 'Member' + i; }, kind: 'method' as const,
 }));
 const model: HostObjectModel = { source: 'Frozen work-counter model', types: { 'Excel.Range': { displayName: 'Range', members } }, aliases: { range: 'Excel.Range' }, globals: {} };
 Object.freeze(members); Object.freeze(model.types['Excel.Range']); Object.freeze(model.types); Object.freeze(model.aliases); Object.freeze(model.globals); Object.freeze(model);
 const source = (references: number, name = 'Member' + (count - 1)) => 'Sub Main()\nDim r As Range\n' + Array.from({ length: references }, () => 'r.' + name).join('\n') + '\nEnd Sub\n';
 return { model, source, reads: () => reads, reset: () => { reads = 0; } };
}
describe('semantic host member lookup work', () => {
 it.each([5, 1000])('indexes %i frozen members once for 1000 last-member references', count => {
  const f = fixture(count); collectHostMemberMethodTokens(f.source(1), { model: f.model }); f.reset();
  const source = f.source(1000), tokens = collectHostMemberMethodTokens(source, { model: f.model });
  expect(tokens).toHaveLength(1000);
  expect(tokens.every(t => t.tokenType === 'function' && source.slice(t.span.start, t.span.end) === 'Member' + (count - 1))).toBe(true);
  expect(f.reads()).toBe(count);
 });
 it('does not build a member index when a pass has no dotted references', () => {
  const f = fixture(1000); collectHostMemberMethodTokens(f.source(1), { model: f.model }); f.reset();
  expect(collectHostMemberMethodTokens('Sub Main()\nDim r As Range\nEnd Sub', { model: f.model })).toEqual([]);
  expect(f.reads()).toBe(0);
 });
 it('retains cheap uncached first-member API calls', () => {
  const f = fixture(1000); collectHostMemberMethodTokens(f.source(1), { model: f.model }); f.reset();
  const source = f.source(1, 'Member0'), offset = source.indexOf('r.Member0') + 'r.Member0'.length;
  expect(resolveHostMemberKindAt(source, offset, 'mEmBeR0', { model: f.model })).toBe('method');
  expect(f.reads()).toBe(1);
 });
 it('does not reuse surfaces across collector passes with changed code names', () => {
  const model: HostObjectModel = { source: 'Two host types', types: {
   'Excel.Range': { displayName: 'Range', members: [{ name: 'Member', kind: 'method' }] },
   'Excel.Other': { displayName: 'Other', members: [{ name: 'Member', kind: 'property' }] },
  }, aliases: {}, globals: {} };
  const source = 'Sub Main()\nSheet1.Member\nEnd Sub';
  expect(collectHostMemberMethodTokens(source, { model, codeNames: { sheet1: 'Excel.Range' } }).map(t => t.tokenType)).toEqual(['function']);
  expect(collectHostMemberMethodTokens(source, { model, codeNames: { sheet1: 'Excel.Other' } }).map(t => t.tokenType)).toEqual(['property']);
 });
 it('preserves first case-insensitive member matches', () => {
  const model: HostObjectModel = { source: 'Duplicate member names', types: { 'Excel.Range': { displayName: 'Range', members: [
   { name: 'Member', kind: 'method' }, { name: 'MEMBER', kind: 'property' },
  ] } }, aliases: { range: 'Excel.Range' }, globals: {} };
  expect(collectHostMemberMethodTokens('Sub Main()\nDim r As Range\nr.mEmBeR\nEnd Sub', { model }).map(t => t.tokenType)).toEqual(['function']);
 });
 it('keeps project type shadowing fresh across passes', () => {
  const f = fixture(5), source = f.source(1);
  expect(collectHostMemberMethodTokens(source, { model: f.model })).toHaveLength(1);
  expect(collectHostMemberMethodTokens(source, { model: f.model, projectTypes: [{ name: 'Range', kind: 'class' }] })).toEqual([]);
  expect(collectHostMemberMethodTokens(source, { model: f.model })).toHaveLength(1);
 });
});
