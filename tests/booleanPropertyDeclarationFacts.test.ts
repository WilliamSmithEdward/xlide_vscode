import { describe, expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
type Hit = Parameters<Parameters<typeof checkAssignmentTypes>[6]>;
function prepare(source: string, vba7?: boolean) {
 const mod = parseModule(source), environment = vba7 === undefined ? undefined : { compilerConstants: { VBA7: vba7 } };
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod, conditionalCompilation: environment }), sourceTokens = tokenizeCached(source).filter(token => token.kind !== 'comment');
 const activity = environment ? createConditionalActivityTracker(mod, environment) : undefined;
 return { symbols, run(): Hit[] { const hits: Hit[] = []; const ctx = { parsedModule: mod, sourceTokens, withScanCache: new Map(), receiverTypeCache: new Map(), receiverChainCache: new Map(), memberSurfaceCache: new Map(), allowSetAssignmentRefinement: false }; checkAssignmentTypes(source, mod, symbols, undefined, ctx, activity, (...hit) => { hits.push(hit); }); return hits; } };
}
const property = (value: string) => 'Range("A1").Font.Size = ' + value + '\n';
describe('numeric host-property Boolean declarations', () => {
 it.each(['long', 'true', 'false', 'integer', 'literal', 'distinct', 'true-first'])('bounds actual symbol name reads (%s)', (mode) => {
  const declarations = Array.from({ length: 100 }, (_, i) => 'Dim k' + i + ' As Long\n').join('');
  const declaration = 'Dim v As ' + (mode === 'true' || mode === 'false' || mode === 'true-first' ? 'Boolean' : mode === 'integer' ? 'Integer' : 'Long') + '\n';
  const source = 'Sub P()\n' + (mode === 'true-first' ? declaration : '') + declarations + (mode === 'true-first' ? '' : declaration) + 'v = ' + (mode === 'long' ? '500' : mode === 'false' ? 'False' : 'True') + '\n' + Array.from({ length: 100 }, (_, i) => property(mode === 'literal' ? '12' : mode === 'distinct' ? 'k' + i : 'v')).join('') + 'End Sub';
  const { symbols, run } = prepare(source); let reads = 0;
  for (const symbol of symbols.all) { const name = symbol.name; Object.defineProperty(symbol, 'name', { enumerable: true, get() { reads++; return name; } }); }
  expect(run()).toHaveLength(['long', 'false', 'integer', 'distinct'].includes(mode) ? 100 : 0);
  expect(reads).toBeLessThan(3000);
 });
 it('preserves statement-specific True, False and numeric values', () => {
  const source = 'Sub P()\nDim v As Boolean\nv = True\n' + property('v') + 'v = False\n' + property('v') + 'v = True\n' + property('v') + 'End Sub';
  const hits = prepare(source).run(); expect(hits).toHaveLength(1); expect(hits[0][2].start).toBe(source.indexOf('= v', source.indexOf('v = False')) + 2);
 });
 it('considers any matching direct child rather than only the first declaration/kind', () => {
  const { symbols, run } = prepare('Sub P()\nDim v As Long\nv = -1\n' + property('v') + 'End Sub');
  const proc = symbols.root.children!.find(symbol => symbol.kind === 'sub')!, original = proc.children!.find(symbol => symbol.name === 'v')!;
  proc.children!.push({ ...original, kind: 'event', asType: 'BoOlEaN' });
  expect(run()).toEqual([]);
 });
 it('refreshes reused bound roots and uses exact type lowercasing without trimming', () => {
  const { symbols, run } = prepare('Sub P()\nDim v As Boolean\nv = True\n' + property('v') + 'End Sub');
  const child = symbols.all.find(symbol => symbol.name === 'v')!;
  expect(run()).toEqual([]);
  child.asType = ' Boolean '; expect(run()).toHaveLength(1);
  child.asType = 'BoOlEaN'; expect(run()).toEqual([]);
  child.asType = 'Long'; expect(run()).toHaveLength(1);
 });
 it('keeps declarations specific to their procedure', () => {
  const source = 'Sub P()\nDim v As Boolean\nv = True\n' + property('v') + 'End Sub\nSub Q()\nDim v As Long\nv = -1\n' + property('v') + 'End Sub';
  const hits = prepare(source).run(); expect(hits).toHaveLength(1); expect(hits[0][2].start).toBe(source.lastIndexOf('= v') + 2);
 });
 it('refreshes conditional Boolean/numeric declarations for reused parsed source', () => {
  const source = 'Sub P()\n#If VBA7 Then\nDim v As Boolean\n#Else\nDim v As Integer\n#End If\nv = True\n' + property('v') + 'End Sub';
  expect(prepare(source, true).run()).toEqual([]);
  expect(prepare(source, false).run()).toHaveLength(1);
  expect(prepare(source, true).run()).toEqual([]);
 });
});
