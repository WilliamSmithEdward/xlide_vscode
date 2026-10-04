import { describe, expect, it } from 'vitest';
import { checkSetAssignments, checkMidStatementLiteralTarget } from '../src/analyzer/diagnostics/rules/assignments';
import { forEachStatementWithHeaders } from '../src/analyzer/diagnostics/walker';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { statementTokensCached } from '../src/analyzer/lexer/tokenHelpers';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
type Hit = Parameters<Parameters<typeof checkSetAssignments>[4]>;
function prepare(source: string) {
 const mod = parseModule(source), sourceTokens = tokenizeCached(source).filter(token => token.kind !== 'comment');
 return (surfaces: VbaProjectClassMembers[]): Hit[] => {
  const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), hits: Hit[] = [];
  const ctx = { projectClassMembers: surfaces, parsedModule: mod, sourceTokens, withScanCache: new Map(), receiverTypeCache: new Map(), receiverChainCache: new Map(), memberSurfaceCache: new Map(), allowSetAssignmentRefinement: false };
  for (const proc of mod.members) if (proc.kind === 'Procedure') forEachStatementWithHeaders(source, proc.body, stmt => { const tokens = statementTokensCached(source, stmt.span); for (const token of tokens) Object.freeze(token); Object.freeze(tokens); });
  checkMidStatementLiteralTarget(source, mod, symbols, undefined, (...hit) => { hits.push(hit); });
  const visitor = checkSetAssignments(source, symbols, undefined, ctx, (...hit) => { hits.push(hit); });
  for (const proc of mod.members) if (proc.kind === 'Procedure') { const visit = visitor(proc); if (visit) forEachStatementWithHeaders(source, proc.body, visit); }
  return hits;
 };
}
describe('assignment significant token reuse', () => {
 it.each([
  'Set o = Nothing \' trailing comment',
  'L: Set o = Nothing',
  '10 Set o = Nothing',
  'If True Then Set o = Nothing Else Set o = Nothing',
  'Set o = _\n Nothing \' comment after continuation',
  'Set o(0) = Nothing \' array target',
  'Set o = Unknown(1, 2, 3) \' unknown actual',
 ])('keeps significant cached tokens immutable: %s', (write) => {
  const declaration = write.includes('o(0)') ? 'Dim o(0) As Object' : 'Dim o As Object';
  expect(prepare('Sub P()\n' + declaration + '\n' + write + '\nEnd Sub')([])).toEqual([]);
 });
 it.each(['Mid', 'Mid$', 'MidB'])('preserves %s literal spans, quoted apostrophes and trailing comments', (head) => {
  const source = 'Sub P()\nL: ' + head + '("a\'b", 1) = "x" \' tail\nEnd Sub';
  const hits = prepare(source)([]); expect(hits).toHaveLength(1); expect(hits[0][0]).toBe('midStatementLiteralTarget');
  expect(source.slice(hits[0][2].start, hits[0][2].end)).toBe('"a\'b"');
 });
 it.each(['Dim Mid As Variant', 'ReDim Mid(2)'])('retains intrinsic shadowing: %s', (declaration) => {
  expect(prepare('Sub P()\n' + declaration + '\nMid("abc", 1) = "x"\nEnd Sub')([])).toEqual([]);
 });
 it('keeps scalar and Variant literal diagnostics', () => {
  const hits = prepare('Sub P()\nDim n As Long\nDim v As Variant\nSet n = Nothing \' scalar\nSet v = 5 \' literal\nEnd Sub')([]);
  expect(hits.map(hit => hit[0])).toEqual(['setRequiresObject', 'setRequiresObject']);
 });
});
