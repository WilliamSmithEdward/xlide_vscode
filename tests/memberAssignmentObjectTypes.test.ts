import { expect, it } from 'vitest';
import { checkAssignmentTypes } from '../src/analyzer/diagnostics/rules/assignments';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
const source = 'Sub Go()\nDim holder As New Holder\nholder.Item = 1\nEnd Sub\n';
function surface(name: string, kind: VbaProjectClassMembers['kind'] = 'class', members: VbaProjectClassMember[] = []): VbaProjectClassMembers { return { name, moduleName: name, kind, exhaustive: true, members }; }
function item(type: string, name = 'Item'): VbaProjectClassMember { return { name, moduleName: 'Holder', kind: 'property', returns: type, writable: true }; }
function run(text: string, classes: VbaProjectClassMembers[], mod = parseModule(text)) {
 const symbols = buildModuleSymbols('M', 'standard', text, { parsedModule: mod }), out: unknown[][] = [];
 checkAssignmentTypes(text, mod, symbols, undefined, { projectClassMembers: classes }, undefined, (...v) => out.push(v)); return out;
}
function expected(text: string, label = 'holder.Item', type = 'Class1') { const start = text.indexOf(label) + 7; return ['setRequired', "Object assignment to '" + label + "' requires Set because it expects " + type + '.', { start, end: start + label.length - 7 }]; }
for (const distinct of [false, true]) it.each([10, 100, 1000])('bounds member object predicates at %i assignments, distinct=' + distinct, count => {
 let names = 0;
 const classes = [surface('Holder', 'class', distinct ? Array.from({ length: count }, (_, i) => item('Class' + i, 'Item' + i)) : [item('Class0')]), ...Array.from({ length: count + 1 }, (_, i) => ({ ...surface('Class' + i), get name() { names++; return 'Class' + i; } }))];
 const text = ['Option Explicit', 'Sub Go()', 'Dim holder As New Holder', ...Array.from({ length: count }, (_, i) => 'holder.Item' + (distinct ? i : '') + ' = 1'), 'End Sub', ''].join('\n');
 let from = 0;
 const want = Array.from({ length: count }, (_, i) => { const label = 'holder.Item' + (distinct ? i : ''), start = text.indexOf(label, from) + 7; from = start + label.length - 7; return ['setRequired', "Object assignment to '" + label + "' requires Set because it expects Class" + (distinct ? i : 0) + '.', { start, end: start + label.length - 7 }]; });
 expect(run(text, classes)).toEqual(want); expect(names).toBeLessThanOrEqual(count * 10 + 100);
});
it.each(['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const)('preserves %s expected-object eligibility', kind => { expect(run(source, [surface('Holder', 'class', [item('Class1')]), surface('Class1', kind)])).toEqual(['class', 'document', 'userform'].includes(kind) ? [expected(source)] : []); });
it.each([['Object', true], ['Collection', true], ['Worksheet', true], ['Long', false], ['Variant', false], ['Missing', false], ['cLaSs1', true]] as const)('preserves %s expected-member type', (type, known) => { expect(run(source, [surface('Holder', 'class', [item(type)]), surface('Class1')])).toEqual(known ? [expected(source, 'holder.Item', type)] : []); });
it.each([{ letAccessor: true }, { isArray: true }, { writable: undefined }])('preserves member guard %j', guard => { expect(run(source, [surface('Holder', 'class', [{ ...item('Class1'), ...guard }]), surface('Class1')])).toEqual([]); });
it('preserves ambiguous eligible object names', () => { expect(run(source, [surface('Holder', 'class', [item('Class1')]), surface('Class1'), surface('CLASS1')])).toEqual([]); });
it('refreshes expected-member metadata with a retained AST', () => {
 const mod = parseModule(source), classes = [surface('Holder', 'class', [item('Class1')])]; expect(run(source, classes, mod)).toEqual([]);
 classes.push(surface('Class1')); expect(run(source, classes, mod)).toEqual([expected(source)]);
 classes[0].members[0].letAccessor = true; expect(run(source, classes, mod)).toEqual([]);
});
