import { expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { checkParentheses } from '../src/analyzer/diagnostics/rules/parentheses';
import { checkArgumentTypes } from '../src/analyzer/diagnostics/rules/argumentTypes';
import { checkObjectVariableNotSet } from '../src/analyzer/diagnostics/rules/objectState';
import { createObjectDefaultQueries } from '../src/analyzer/diagnostics/typeInference';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import { forEachStatement } from '../src/analyzer/parser/statementWalk';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
function surface(name: string, mode = 'no-default'): VbaProjectClassMembers {
 return { name, moduleName: name, kind: 'class', exhaustive: true, members: mode.includes('no-default') ? [] : [{ name: 'Item', moduleName: name, kind: 'property', returns: 'Long', defaultMember: true, signature: 'Item()', writable: mode === 'writable-default' }] };
}
function run(source: string, classes: VbaProjectClassMembers[], mode: string, mod = parseModule(source)) {
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), ctx = { projectClassMembers: classes }, out: unknown[][] = [];
 if (mode === 'primed-no-default') { checkObjectVariableNotSet(source, mod, symbols, ctx, undefined, () => {}); }
 if (mode === 'parentheses') { checkParentheses(source, mod, symbols, ctx, undefined, (...v) => out.push(v)); }
 else if (mode === 'scalar-variable' || mode === 'scalar-new') {
  const factory = checkArgumentTypes(source, symbols, undefined, undefined, ctx, (...v) => out.push(v));
  for (const p of mod.members) { if (p.kind !== 'Procedure') { continue; } const visit = factory(p); if (visit) { forEachStatement(p.body, visit); } }
 } else { checkAssignmentTypes(source, mod, symbols, undefined, ctx, undefined, (...v) => out.push(v)); }
 return out;
}
for (const mode of ['cold-no-default', 'primed-no-default', 'writable-default', 'readonly-default']) {
 it.each([10, 100, 1000])('bounds ' + mode + ' at %i distinct assignments', count => {
  let names = 0;
  const classes = Array.from({ length: count + 1 }, (_, i) => ({ ...surface('Class' + i, mode), get name() { names++; return 'Class' + i; } }));
  const source = ['Option Explicit', ...Array.from({ length: count }, (_, i) => 'Sub P' + i + '()\nDim actor As Class' + i + '\nactor = 1\nEnd Sub'), ''].join('\n');
  let from = 0;
  const expected = mode === 'writable-default' ? [] : Array.from({ length: count }, (_, i) => {
   const start = source.indexOf('actor =', from); from = start + 5;
   return [mode === 'readonly-default' ? 'readonlyMemberAssignment' : 'setRequired', mode === 'readonly-default'
    ? "Assignment to 'actor' reaches the default member Item of Class" + i + ', a Property Get with no Property Let. This is a VBE compile error: Invalid use of property.'
    : "Assignment to 'actor' requires Set: Class" + i + " has no default member for a Let to reach. It is still Nothing here, so this will raise Run-time error '91': Object variable or With block variable not set.", { start, end: start + 5 }];
  });
  expect(run(source, classes, mode)).toEqual(expected);
  expect(names).toBeLessThanOrEqual(count * 12 + 50);
 });
}
for (const mode of ['parentheses', 'scalar-variable', 'scalar-new']) for (const distinct of [false, true]) {
 it.each([10, 100, 1000])('bounds ' + mode + ' at %i reads, distinct=' + distinct, count => {
  let names = 0;
  const classes = Array.from({ length: count + 1 }, (_, i) => ({ ...surface('Class' + i), get name() { names++; return 'Class' + i; } }));
  const source = ['Option Explicit', ...(mode.startsWith('scalar') ? ['Sub Take(ByVal value As Long)', 'Debug.Print value', 'End Sub'] : []), ...Array.from({ length: count }, (_, i) => 'Sub P' + i + '()\nDim actor As Class' + (distinct ? i : 0) + '\n' + (mode === 'parentheses' ? 'Debug.Print (actor)' : mode === 'scalar-new' ? 'Take New Class' + (distinct ? i : 0) : 'Take actor') + '\nEnd Sub'), ''].join('\n');
  expect(run(source, classes, mode)).toEqual([]);
  expect(names).toBeLessThanOrEqual(count * 12 + 50);
 });
}
it.each(['document', 'userform', 'userType', 'enum', 'standardModule'] as const)('factory first-class lookup skips %s surfaces', kind => {
 const prefix = { ...surface('CLASS1'), kind }, first = surface('Class1'), later = surface('CLASS1');
 const query = createObjectDefaultQueries({ projectClassMembers: [prefix, first, later] });
 expect(query.projectClassNamed('class1')).toBe(first);
 expect(query.projectTypeNamed('class1')).toBe(prefix);
});
it('factory first-class lookup retains an incomplete first class', () => {
 const first = surface('Class1'); first.exhaustive = false;
 expect(createObjectDefaultQueries({ projectClassMembers: [first, surface('CLASS1')] }).projectClassNamed('class1')).toBe(first);
});
it('factory first-class lookup refreshes in a new query', () => {
 const classes = [surface('Other')], ctx = { projectClassMembers: classes };
 expect(createObjectDefaultQueries(ctx).projectClassNamed('class1')).toBeUndefined();
 classes.unshift(surface('Class1'));
 expect(createObjectDefaultQueries(ctx).projectClassNamed('class1')).toBe(classes[0]);
});
it('factory first-class lookup does not read unrelated member arrays', () => {
 const extra = surface('Extra'); Object.defineProperty(extra, 'members', { get() { throw Error('unused members'); } });
 expect(createObjectDefaultQueries({ projectClassMembers: [surface('Class1'), extra] }).projectClassNamed('class1')?.name).toBe('Class1');
});
it('preserves assignment metadata freshness on a retained AST', () => {
 const source = 'Sub Go()\nDim actor As Class1\nactor = 1\nEnd Sub\n', mod = parseModule(source), classes = [surface('Class1', 'writable-default')];
 expect(run(source, classes, 'writable-default', mod)).toEqual([]);
 classes[0].members[0].writable = false;
 expect(run(source, classes, 'readonly-default', mod)).toEqual([['readonlyMemberAssignment', "Assignment to 'actor' reaches the default member Item of Class1, a Property Get with no Property Let. This is a VBE compile error: Invalid use of property.", { start: 29, end: 34 }]]);
});
