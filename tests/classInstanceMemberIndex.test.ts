import { expect, it } from 'vitest';
import { checkClassInstanceValues } from '../src/analyzer/diagnostics/rules/classInstanceValues';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

function run(source: string, type: VbaProjectClassMembers) {
	const mod = parseModule(source);
	const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
	const output: unknown[][] = [];
	checkClassInstanceValues(source, mod, symbols, { projectClassMembers: [type] }, undefined, (...args) => output.push(args));
	return output;
}

function member(name: string, knownValue: VbaProjectClassMember['knownValue'] = 'nothing'): VbaProjectClassMember {
	return { name, moduleName: 'Class1', kind: 'property', returns: 'Object', knownValue };
}
function type(members: VbaProjectClassMember[]): VbaProjectClassMembers {
	return { name: 'Class1', moduleName: 'Class1', kind: 'class', members };
}

for (const distinct of [false, true]) {
	it.each([10, 100, 1000])('bounds member-name reads for %i members, distinct=' + distinct, count => {
		let reads = 0;
		const metadata = type(Array.from({ length: count }, (_, i) => ({
			...member('M' + i),
			get name() { reads++; return 'M' + i; },
		})));
		const source = ['Sub Go()', 'Dim actor As New Class1',
			...Array.from({ length: count }, (_, i) => 'Debug.Print actor.M' + (distinct ? i : count - 1) + '.Count'),
			'End Sub', ''].join('\n');
		const output = run(source, metadata);
		expect(output).toHaveLength(count);
		for (let i = 0; i < count; i++) {
			const name = 'actor.M' + (distinct ? i : count - 1);
			const start = source.indexOf(name + '.Count', i === 0 ? 0 : (output[i - 1][2] as { end: number }).end);
			expect(output[i]).toEqual(['objectVariableNotSet',
				`'${name}' is Nothing here: nothing in Class1 sets M${distinct ? i : count - 1}. This will raise Run-time error '91': Object variable or With block variable not set.`,
				{ start, end: start + name.length }]);
		}
		expect(reads).toBeLessThanOrEqual(count * 2);
	});
}

it.each([2, 10])('keeps the first case-insensitive duplicate at %i members', count => {
	const source = 'Sub Go()\nDim actor As New Class1\nDebug.Print actor.[M].Count\nEnd Sub\n';
	const metadata = type([member('m', undefined), member('M'), ...Array.from({ length: count - 2 }, (_, i) => member('Other' + i))]);
	// Explicitly clear the first fact: the helper defaults undefined to Nothing.
	metadata.members[0].knownValue = undefined;
	expect(run(source, metadata)).toEqual([]);
});

it('rebuilds after metadata changes between queries', () => {
	const source = 'Sub Go()\nDim actor As New Class1\nDebug.Print actor.M.Count\nEnd Sub\n';
	const metadata = type([member('M'), ...Array.from({ length: 9 }, (_, i) => member('Other' + i))]);
	expect(run(source, metadata)).toHaveLength(1);
	metadata.members[0] = member('M');
	metadata.members[0].knownValue = undefined;
	expect(run(source, metadata)).toEqual([]);
});

it('does not inspect members when no tracked member is read', () => {
	const metadata = type(Array.from({ length: 100 }, () => ({ ...member('M'), get name(): string { throw new Error('unused member accessed'); } })));
	expect(run('Sub Go()\nDim actor As New Class1\nDebug.Print 1\nEnd Sub\n', metadata)).toEqual([]);
});

it('shares a consulted index across procedures within one query', () => {
	let reads = 0;
	const metadata = type(Array.from({ length: 100 }, (_, i) => ({ ...member('M' + i), get name() { reads++; return 'M' + i; } })));
	const source = ['Sub A()', 'Dim actor As New Class1', 'Debug.Print actor.M99.Count', 'End Sub',
		'Sub B()', 'Dim actor As New Class1', 'Debug.Print actor.M99.Count', 'End Sub', ''].join('\n');
	expect(run(source, metadata)).toHaveLength(2);
	expect(reads).toBe(102);
});
