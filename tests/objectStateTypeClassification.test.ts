import { expect, it } from 'vitest';
import { checkObjectVariableNotSet, objectLetStateAt } from '../src/analyzer/diagnostics/rules/objectState';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

function surface(name: string, kind: VbaProjectClassMembers['kind'] = 'class'): VbaProjectClassMembers {
 return { name, moduleName: name, kind, exhaustive: true, members: [{ name: 'Value', kind: 'property', moduleName: name, returns: 'Long' }] };
}
function run(source: string, classes: VbaProjectClassMembers[], mod = parseModule(source)) {
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), out: unknown[][] = [];
 checkObjectVariableNotSet(source, mod, symbols, { projectClassMembers: classes }, undefined, (...v) => out.push(v));
 return out;
}
const families = ['functions', 'module-variable', 'locals', 'arrays'] as const;
for (const family of families) for (const distinct of family === 'module-variable' ? [false] : [false, true]) {
 it.each([10, 100, 1000])('bounds ' + family + ' classification at %i procedures, distinct=' + distinct, count => {
  let names = 0;
  const classes = Array.from({ length: count + 1 }, (_, i) => ({ ...surface('Class' + i), get name() { names++; return 'Class' + i; } }));
  const lines = ['Option Explicit'];
  if (family === 'module-variable') { lines.push('Private actor As Class0'); }
  for (let i = 0; i < count; i++) {
   const type = 'Class' + (distinct ? i : 0);
   if (family === 'functions') { lines.push('Function P' + i + '() As ' + type, 'End Function'); }
   else {
    lines.push('Sub P' + i + '()');
    if (family === 'locals') { lines.push('Dim actor As ' + type); }
    if (family === 'arrays') { lines.push('Dim actor(0) As ' + type); }
    lines.push('End Sub');
   }
  }
  expect(run(lines.join('\n') + '\n', classes)).toEqual([]);
  expect(names).toBeLessThanOrEqual(count * 8 + 40);
 });
}
const controls = [
 ['locals', 'Sub Go()\nDim actor As Class1\nDebug.Print actor.Value\nEnd Sub\n', "Object variable 'actor' is Nothing before member access. This will raise Run-time error '91': Object variable or With block variable not set.", 41, 46],
 ['module-variable', 'Private actor As Class1\nSub Go()\nDebug.Print actor.Value\nEnd Sub\n', "Object variable 'actor' is never set anywhere in this module, so it is Nothing here. This will raise Run-time error '91': Object variable or With block variable not set.", 45, 50],
 ['arrays', 'Sub Go()\nDim actor(0) As Class1\nDebug.Print actor(0).Value\nEnd Sub\n', "Element actor(0) of 'actor' is Nothing before member access. This will raise Run-time error '91': Object variable or With block variable not set.", 44, 52],
 ['functions', 'Function Factory() As Class1\nEnd Function\nSub Go()\nDebug.Print Factory().Value\nEnd Sub\n', "Function 'Factory' never sets its result, so it returns Nothing, and '.Value' has no object to reach. This will raise Run-time error '91': Object variable or With block variable not set.", 63, 72],
] as const;
for (const [family, source, message, start, end] of controls) {
 it.each(['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const)('preserves ' + family + ' %s eligibility and exact finding', kind => {
  expect(run(source, [surface('Class1', kind)])).toEqual(['class', 'document', 'userform'].includes(kind) ? [['objectVariableNotSet', message, { start, end }]] : []);
 });
}
it.each(['Dim actor As New Class1', 'Static actor As Class1', 'Dim actor As Long'])('preserves %s exclusion', declaration => {
 expect(run('Sub Go()\n' + declaration + '\nDebug.Print actor.Value\nEnd Sub\n', [surface('Class1')])).toEqual([]);
});
it('refreshes classification under a new context with a retained AST', () => {
 const [, source, message, start, end] = controls[0], mod = parseModule(source), classes: VbaProjectClassMembers[] = [];
 expect(run(source, classes, mod)).toEqual([]);
 classes.push(surface('Class1'));
 expect(run(source, classes, mod)).toEqual([['objectVariableNotSet', message, { start, end }]]);
 classes.push(surface('CLASS1'));
 expect(run(source, classes, mod)).toEqual([]);
});
it('keeps direct Let-state queries working before the public rule fills its walk cache', () => {
 const source = 'Sub Go()\nDim actor As Class1\nactor = 1\nEnd Sub\n', mod = parseModule(source);
 const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), proc = mod.members.find(p => p.kind === 'Procedure');
 if (!proc || proc.kind !== 'Procedure') { throw Error('missing procedure'); }
 expect(objectLetStateAt(source, mod, proc, symbols, { projectClassMembers: [surface('Class1')] }, undefined, source.indexOf('actor ='))).toBe('unset');
});
