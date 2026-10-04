import { expect, it, vi } from 'vitest';

const work = vi.hoisted(() => ({ words: 0 }));
vi.mock('../src/analyzer/diagnostics/walker', async original => {
	const actual = await original<typeof import('../src/analyzer/diagnostics/walker')>();
	return {
		...actual,
		tokenText: (...args: Parameters<typeof actual.tokenText>) => {
			work.words++;
			return actual.tokenText(...args);
		},
		tokenName: (...args: Parameters<typeof actual.tokenName>) => {
			work.words++;
			return actual.tokenName(...args);
		},
	};
});

import { checkClassInstanceValues } from '../src/analyzer/diagnostics/rules/classInstanceValues';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

const classes: VbaProjectClassMembers[] = [{
	name: 'Class1', moduleName: 'Class1', kind: 'class',
	members: [{ name: 'M', moduleName: 'Class1', kind: 'property', returns: 'Object', knownValue: 'nothing' }],
}];

function prepare(source: string) {
	const mod = parseModule(source);
	return { mod, symbols: buildModuleSymbols('M', 'standard', source, { parsedModule: mod }) };
}

function run(source: string) {
	const { mod, symbols } = prepare(source);
	const output: unknown[][] = [];
	checkClassInstanceValues(source, mod, symbols, { projectClassMembers: classes }, undefined, (...args) => output.push(args));
	return output;
}

for (const kind of ['scalar', 'class'] as const) {
	it.each([10, 100, 1000])('bounds total token-name/word queries at %i locals and statements, ' + kind, count => {
		const source = [
			'Sub Go()',
			...Array.from({ length: count }, (_, i) => 'Dim actor' + i + ' As ' + (kind === 'scalar' ? 'Long' : 'New Class1')),
			...Array(count).fill('Debug.Print 1'),
			'End Sub', '',
		].join('\n');
		const { mod, symbols } = prepare(source);
		const output: unknown[][] = [];
		work.words = 0;
		checkClassInstanceValues(source, mod, symbols, { projectClassMembers: classes }, undefined, (...args) => output.push(args));
		const words = work.words;
		expect(output).toEqual([]);
		expect(words).toBeLessThanOrEqual(count * 8 + 10);
	});
}

for (const declaration of ['Dim actor As New Class1', 'Dim actor As Class1', 'Dim actor As Object', 'Dim actor As Variant', 'Dim actor']) {
	it('retains a tracked instance for ' + declaration, () => {
		const setup = declaration.includes('New') ? [] : ['Set actor = New Class1'];
		const source = ['Sub Go()', declaration, ...setup, 'Debug.Print actor.M.Count', 'End Sub', ''].join('\n');
		const start = source.indexOf('actor.M');
		expect(run(source)).toEqual([[
			'objectVariableNotSet',
			"'actor.M' is Nothing here: nothing in Class1 sets M. This will raise Run-time error '91': Object variable or With block variable not set.",
			{ start, end: start + 'actor.M'.length },
		]]);
	});
}

for (const body of [
	'Set actor = New Class1\nSet actor = New Class1\nDebug.Print actor.M.Count',
	'Debug.Print actor\nDebug.Print actor.M.Count',
	'Take actor\nDebug.Print actor.M.Count',
	'Set actor.M = New Collection\nDebug.Print actor.M.Count',
	'Set [actor].[M] = New Collection\nDebug.Print actor.M.Count',
]) {
	it('preserves Set multiplicity, escape and assigned-field suppression: ' + body, () => {
		expect(run('Sub Go()\nDim actor As New Class1\n' + body + '\nEnd Sub\n')).toEqual([]);
	});
}

for (const declaration of ['Dim actor As Long', 'Dim actor As Other', 'Dim actor() As Class1', 'Static actor As Class1']) {
	it('excludes unsupported declaration ' + declaration, () => {
		expect(run('Sub Go()\n' + declaration + '\nSet actor = New Class1\nDebug.Print actor.M.Count\nEnd Sub\n')).toEqual([]);
	});
}
