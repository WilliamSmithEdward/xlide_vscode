import { expect, it } from 'vitest';
import { checkParameterDefaultValues, checkNonConstantParameterDefaults } from '../src/analyzer/diagnostics/rules/declarations';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

const rules = { checkParameterDefaultValues, checkNonConstantParameterDefaults };
function surface(name: string, kind: VbaProjectClassMembers['kind'] = 'class'): VbaProjectClassMembers {
	return { name, moduleName: name, kind, members: [] };
}
function run(rule: keyof typeof rules, source: string, classes: VbaProjectClassMembers[]) {
	const mod = parseModule(source), output: unknown[][] = [];
	rules[rule](source, mod, createConditionalActivityTracker(mod), { projectClassMembers: classes }, (...args) => output.push(args));
	return output;
}
function sourceFor(types: string[]) {
	return types.map((type, i) => `Sub P${i}(Optional actor As ${type} = 0, Optional n As Long = Factory())\nEnd Sub\n`).join('');
}
function expected(rule: keyof typeof rules, source: string, types: string[]) {
	let from = 0;
	return types.map(type => {
		const needle = rule === 'checkParameterDefaultValues' ? '= 0' : 'Factory()', index = source.indexOf(needle, from);
		const start = index + (needle === '= 0' ? 2 : 0), end = index + needle.length;
		from = end;
		return rule === 'checkParameterDefaultValues'
			? ['parameterDefaultTypeMismatch', `Optional parameter 'actor' expects ${type}, but its default value is numeric literal 0. Optional object parameter defaults must be Nothing.`, { start, end }]
			: ['parameterDefaultNotConstant', "Optional parameter 'n' default must be a constant expression; the call 'Factory(...)' is not constant.", { start, end }];
	});
}
for (const rule of Object.keys(rules) as (keyof typeof rules)[]) {
	for (const distinct of [false, true]) {
		it.each([10, 100, 1000])(`${rule} bounds lookup work across %i procedures, distinct=${distinct}`, count => {
			let names = 0;
			const classes: VbaProjectClassMembers[] = Array.from({ length: count + 1 }, (_, i) => ({
				...surface('Class' + (i + 1)), get name() { names++; return 'Class' + (i + 1); },
			}));
			const types = Array.from({ length: count }, (_, i) => 'Class' + (distinct ? i + 1 : 1)), source = sourceFor(types);
			expect(run(rule, source, classes)).toEqual(expected(rule, source, types));
			expect(names).toBeLessThanOrEqual(count * 3 + 5);
		});
	}
	it(`${rule} skips project metadata for scalar, generic, host, library and absent defaults`, () => {
		const classes = [{ ...surface('unused'), get name(): string { throw new Error('unused project metadata'); } }];
		const source = 'Sub Go(Optional n As Long = 1, Optional v As Variant = 1, Optional o As Object = Nothing, Optional c As Collection = Nothing, Optional w As Worksheet = Nothing, Optional d As Scripting.Dictionary = Nothing, Optional p As Missing)\nEnd Sub\n';
		expect(run(rule, source, classes)).toEqual([]);
	});
	it(`${rule} skips inactive procedure defaults without consulting metadata`, () => {
		const classes = [{ ...surface('unused'), get name(): string { throw new Error('inactive project metadata'); } }];
		expect(run(rule, '#If False Then\n' + sourceFor(['Class1']) + '#End If\n', classes)).toEqual([]);
	});
	it(`${rule} observes changed metadata on the next public query`, () => {
		const source = 'Sub Go(Optional actor As Class1 = 0)\nEnd Sub\n', classes = [surface('Class1')];
		if (rule === 'checkParameterDefaultValues') expect(run(rule, source, classes)).toEqual(expected(rule, source, ['Class1']));
		else expect(run(rule, source, classes)).toEqual([]);
		classes.push(surface('CLASS1'));
		expect(run(rule, source, classes)).toEqual([]);
		classes.pop();
		classes[0].kind = 'enum';
		expect(run(rule, source, classes)).toEqual([]);
	});
}
it.each(['class', 'document', 'userform'] as const)('preserves object defaults for %s metadata', kind => {
	const source = sourceFor(['Class1']), classes = [surface('Class1', kind)];
	expect(run('checkParameterDefaultValues', source, classes)).toEqual(expected('checkParameterDefaultValues', source, ['Class1']));
	expect(run('checkNonConstantParameterDefaults', source, classes)).toEqual(expected('checkNonConstantParameterDefaults', source, ['Project.Class1']));
});
it.each(['userType', 'enum', 'standardModule'] as const)('does not treat %s as an object or ambiguous duplicate', kind => {
	const source = sourceFor(['Class1']);
	expect(run('checkParameterDefaultValues', source, [surface('Class1', kind)])).toEqual([]);
	expect(run('checkParameterDefaultValues', source, [surface('Class1'), surface('CLASS1', kind)])).toEqual(expected('checkParameterDefaultValues', source, ['Class1']));
	const callSource = 'Sub Go(Optional actor As Class1 = Factory())\nEnd Sub\n';
	const start = callSource.indexOf('Factory()');
	expect(run('checkNonConstantParameterDefaults', callSource, [surface('Class1', kind)])).toEqual([
		['parameterDefaultNotConstant', "Optional parameter 'actor' default must be a constant expression; the call 'Factory(...)' is not constant.", { start, end: start + 'Factory()'.length }],
	]);
});
it('preserves ambiguous object-call defaults and rebuilds after ambiguity disappears', () => {
	const source = 'Sub Go(Optional actor As Class1 = Factory())\nEnd Sub\n', classes = [surface('Class1'), surface('CLASS1')];
	const start = source.indexOf('Factory()');
	expect(run('checkNonConstantParameterDefaults', source, classes)).toEqual([
		['parameterDefaultNotConstant', "Optional parameter 'actor' default must be a constant expression; the call 'Factory(...)' is not constant.", { start, end: start + 'Factory()'.length }],
	]);
	classes.pop();
	expect(run('checkNonConstantParameterDefaults', source, classes)).toEqual([]);
});

it('preserves host priority over ambiguous project names', () => {
 const source = sourceFor(['Worksheet']), classes = [surface('Worksheet'), surface('WORKSHEET')];
 expect(run('checkParameterDefaultValues', source, classes)).toEqual(expected('checkParameterDefaultValues', source, ['Worksheet']));
 expect(run('checkNonConstantParameterDefaults', 'Sub Go(Optional actor As Worksheet = Factory())\nEnd Sub\n', classes)).toEqual([]);
});
it('preserves unknown qualified project types', () => {
 const source = sourceFor(['Project.Class1']), classes = [surface('Class1')];
 expect(run('checkParameterDefaultValues', source, classes)).toEqual([]);
 expect(run('checkNonConstantParameterDefaults', source, classes)).toEqual(expected('checkNonConstantParameterDefaults', source, ['Project.Class1']));
});
