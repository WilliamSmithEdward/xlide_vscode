import { expect, it } from 'vitest';
import { checkObjectVariableNotSet, objectLetStateAt } from '../src/analyzer/diagnostics/rules/objectState';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
function surface(name: string, signature?: string): VbaProjectClassMembers {
 return { name, moduleName: name, kind: 'class', exhaustive: true, members: signature ? [{ name: 'Item', kind: 'property', moduleName: name, returns: 'Long', defaultMember: true, signature }] : [] };
}
function run(source: string, classes: VbaProjectClassMembers[], mod = parseModule(source)) {
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), out: unknown[][] = [];
 checkObjectVariableNotSet(source, mod, symbols, { projectClassMembers: classes }, undefined, (...v) => out.push(v));
 return out;
}
const lines = { let: 'actor = 1', condition: 'If actor Then Debug.Print 1', indexed: 'Debug.Print actor(1)' };
for (const path of ['let', 'condition', 'indexed'] as const) for (const layout of ['one-procedure', 'repeated-types', 'distinct-types'] as const) {
 it.each([10, 100, 1000])('bounds ' + path + ' default queries at %i reads, layout=' + layout, count => {
  let names = 0;
  const classes = Array.from({ length: count + 1 }, (_, i) => ({ ...surface('Class' + i), get name() { names++; return 'Class' + i; } }));
  const source = layout === 'one-procedure'
   ? ['Option Explicit', 'Sub Go()', 'Dim actor As Class0', ...Array(count).fill(lines[path]), 'End Sub', ''].join('\n')
   : ['Option Explicit', ...Array.from({ length: count }, (_, i) => 'Sub P' + i + '()\nDim actor As Class' + (layout === 'distinct-types' ? i : 0) + '\n' + lines[path] + '\nEnd Sub'), ''].join('\n');
  expect(run(source, classes)).toEqual([]);
  expect(names).toBeLessThanOrEqual(count * 8 + 30);
 });
}
const cases = [
 ['let', 'actor = 1', 'before the default-member assignment.', 68, 73],
 ['condition', 'If actor Then Debug.Print 1', 'when the condition reads its value.', 71, 76],
 ['value', 'value = actor', 'when its default member is read.', 76, 81],
 ['operand', 'Debug.Print actor + 1', 'when its default member is read as an operand.', 80, 85],
 ['indexed', 'Debug.Print actor(1)', 'when its default member is indexed.', 80, 85],
 ['setIndexed', 'Set target = actor(1)', 'when its default member is indexed.', 81, 86],
 ['ifBlock', 'If actor Then\nDebug.Print 1\nEnd If', 'when the condition reads its value.', 71, 76],
 ['while', 'While actor\nDebug.Print 1\nWend', 'when the condition reads its value.', 74, 79],
 ['do', 'Do While actor\nDebug.Print 1\nLoop', 'when the condition reads its value.', 77, 82],
 ['loop', 'Do\nDebug.Print 1\nLoop While actor', 'when the condition reads its value.', 96, 101],
 ['select', 'Select Case actor\nCase 1\nDebug.Print 1\nEnd Select', 'when Select Case reads its value.', 80, 85],
] as const;
const sourceFor = (body: string) => 'Sub Go()\nDim actor As Class1\nDim value As Long\nDim target As Object\n' + body + '\nEnd Sub\n';
const message = (suffix: string) => "Object variable 'actor' is Nothing " + suffix + " This will raise Run-time error '91': Object variable or With block variable not set.";
for (const [path, line, suffix, start, end] of cases) {
 it('preserves exact ' + path + ' finding and span', () => {
  expect(run(sourceFor(line), [surface('Class1', path === 'indexed' || path === 'setIndexed' ? 'Item(index As Long)' : 'Item()')])).toEqual([['objectVariableNotSet', message(suffix), { start, end }]]);
 });
 it('preserves absent-default exclusion for ' + path, () => { expect(run(sourceFor(line), [surface('Class1')])).toEqual([]); });
}
it('refreshes default metadata in a new context on a retained AST', () => {
 const source = sourceFor('actor = 1'), mod = parseModule(source), classes = [surface('Class1')];
 expect(run(source, classes, mod)).toEqual([]);
 classes[0].members = surface('Class1', 'Item()').members;
 expect(run(source, classes, mod)).toEqual([['objectVariableNotSet', message(cases[0][2]), { start: 68, end: 73 }]]);
 classes.push(surface('CLASS1', 'Item()'));
 expect(run(source, classes, mod)).toEqual([]);
});
it('preserves the incomplete-surface Let verdict', () => {
 const first = surface('Class1'); first.exhaustive = false;
 expect(run(sourceFor('actor = 1'), [first])).toEqual([['objectVariableNotSet', message(cases[0][2]), { start: 68, end: 73 }]]);
});
it('does not consult unrelated project member arrays', () => {
 const unused = surface('Unused'); Object.defineProperty(unused, 'members', { get() { throw Error('unused members'); } });
 expect(run(sourceFor('actor = 1'), [surface('Class1'), unused])).toEqual([]);
});
it('preserves direct Let-state queries before the public rule builds a walk', () => {
 const source = sourceFor('actor = 1'), mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
 const proc = mod.members.find(p => p.kind === 'Procedure'); if (!proc || proc.kind !== 'Procedure') { throw Error('missing procedure'); }
 expect(objectLetStateAt(source, mod, proc, symbols, { projectClassMembers: [surface('Class1', 'Item()')] }, undefined, 68)).toBe('unset');
});
