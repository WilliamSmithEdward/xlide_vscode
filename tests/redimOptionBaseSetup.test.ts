import { describe, expect, it } from 'vitest';
import { checkRedimPreserveDimensions } from '../src/analyzer/diagnostics/rules/arrays';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';

const procedure = (name: string) => `Sub ${name}()\nDim a() As Long\nReDim a(1 To 3)\nReDim Preserve a(4)\nEnd Sub`;
type Hit = Parameters<Parameters<typeof checkRedimPreserveDimensions>[3]>;
function fixture(source: string) {
 const mod = parseModule(source);
 const run = (vba7?: boolean) => {
  const hits: Hit[] = [];
  const activity = vba7 === undefined ? undefined : createConditionalActivityTracker(mod, { compilerConstants: { VBA7: vba7 } });
  checkRedimPreserveDimensions(source, mod, activity, (...hit) => { hits.push(hit); });
  return hits;
 };
 return { mod, run };
}

describe('ReDim Preserve module option setup', () => {
 it('reads module members only linearly when Option Base is absent', () => {
  const count = 200;
  const { mod, run } = fixture(Array.from({ length: count }, (_, i) => procedure('P'+i)).join('\n'));
  let reads = 0;
  for (const member of mod.members) {
   const kind = member.kind;
   Object.defineProperty(member, 'kind', { get: () => { reads++; return kind; } });
  }
  expect(run()).toHaveLength(count);
  expect(reads).toBeLessThan(count * 5);
 });
 it.each([0, 1])('applies Option Base %s to every procedure', base => {
  const { run } = fixture(`Option Base ${base}\n${procedure('P')}\n${procedure('Q')}`);
  expect(run()).toHaveLength(base === 0 ? 2 : 0);
 });
 it('recomputes the option for each conditional activity environment', () => {
  const { run } = fixture(`#If VBA7 Then\nOption Base 1\n#Else\nOption Base 0\n#End If\n${procedure('P')}\n${procedure('Q')}`);
  expect(run(true)).toHaveLength(0);
  expect(run(false)).toHaveLength(2);
  expect(run(true)).toHaveLength(0);
 });
 it('does not inspect Option text when the module has no active procedures', () => {
  const { mod, run } = fixture(`Option Base 1\n#If VBA7 Then\n${procedure('P')}\n#End If`);
  const option = mod.members.find(member => member.kind === 'Option');
  if (!option || option.kind !== 'Option') throw new Error('Missing option');
  Object.defineProperty(option, 'optionText', { get: () => { throw new Error('Unnecessary option read'); } });
  expect(run(false)).toEqual([]);
 });
});
