import { expect, it } from 'vitest';
import { checkTypeOfIsCompatibility } from '../src/analyzer/diagnostics/rules/typeOfIs';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

function surface(name: string, implementsNames: string[] = [], kind: VbaProjectClassMembers['kind'] = 'class'): VbaProjectClassMembers {
	return { name, moduleName: name, kind, members: [], implements: implementsNames };
}
function sourceFor(targets: string[]): string {
	return ['Sub Go(ByVal actor As Class1)', ...targets.map(target => `If TypeOf actor Is ${target} Then\nDebug.Print 1\nEnd If`), 'End Sub', ''].join('\n');
}
function run(source: string, classes: VbaProjectClassMembers[]) {
	const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
	const output: unknown[][] = [];
	const visitor = checkTypeOfIsCompatibility(symbols, { projectClassMembers: classes }, (...args) => output.push(args));
	for (const procedure of mod.members) {
		if (procedure.kind !== 'Procedure') { continue; }
		const visit = visitor(procedure);
		if (!visit) { continue; }
		for (const node of procedure.body) {
			if (node.kind === 'IfBlock') { for (const branch of node.branches) { if (branch.condition) { visit(branch.condition); } } }
		}
	}
	return output;
}
function expected(source: string, targets: string[]) {
	let from = 0;
	return targets.map(target => {
		const text = 'TypeOf actor Is ' + target, start = source.indexOf(text, from);
		from = start + text.length;
		return ['typeOfIsAlwaysFalse', `'TypeOf ... Is ${target}' is always False: 'actor' is declared As Class1, which is never ${target}.`, { start, end: from }];
	});
}

for (const distinct of [false, true]) {
	it.each([10, 100, 1000])('bounds all type and interface metadata work at %i expressions, distinct=' + distinct, count => {
		let names = 0, interfaces = 0;
		const classes: VbaProjectClassMembers[] = Array.from({ length: count + 1 }, (_, i) => ({
			moduleName: 'Class' + (i + 1), kind: 'class', members: [],
			get name() { names++; return 'Class' + (i + 1); },
			get implements() { interfaces++; return []; },
		}));
		const targets = Array.from({ length: count }, (_, i) => 'Class' + (distinct ? i + 2 : 2));
		const source = sourceFor(targets);
		expect(run(source, classes)).toEqual(expected(source, targets));
		expect(names).toBeLessThanOrEqual(count * 4 + 12);
		expect(interfaces).toBeLessThanOrEqual(count * 4 + 12);
	});
}

it.each(['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const)('retains interface exclusion from %s metadata', kind => {
	const classes = [surface('Class1'), surface('Class2'), surface('Bridge', ['cLaSs1'], kind)];
	expect(run(sourceFor(['Class2']), classes)).toEqual([]);
});
it('preserves direct implementation compatibility', () => {
	expect(run(sourceFor(['Class2']), [surface('Class1', ['Class2']), surface('Class2')])).toEqual([]);
});
it('preserves ambiguity from eligible duplicate names', () => {
	expect(run(sourceFor(['Class2']), [surface('Class1'), surface('CLASS1'), surface('Class2')])).toEqual([]);
});
it.each(['userType', 'enum', 'standardModule'] as const)('keeps excluded %s duplicates out of object-type ambiguity', kind => {
	const source = sourceFor(['Class2']);
	expect(run(source, [surface('Class1'), surface('CLASS1', [], kind), surface('Class2')])).toEqual(expected(source, ['Class2']));
});
it('observes changed Implements facts in the next query', () => {
	const source = sourceFor(['Class2']), classes = [surface('Class1'), surface('Class2')];
	expect(run(source, classes)).toEqual(expected(source, ['Class2']));
	classes[1].implements = ['Class1'];
	expect(run(source, classes)).toEqual([]);
});
it('observes type ambiguity introduced in the next query', () => {
	const source = sourceFor(['Class2']), classes = [surface('Class1'), surface('Class2')];
	expect(run(source, classes)).toEqual(expected(source, ['Class2']));
	classes.push(surface('class1'));
	expect(run(source, classes)).toEqual([]);
});
it('does not scan unused project metadata for a generic operand', () => {
	const classes: VbaProjectClassMembers[] = [{ ...surface('Class1'), get name(): string { throw new Error('unused project lookup'); }, get implements(): string[] { throw new Error('unused interface lookup'); } }];
	expect(run('Sub Go(ByVal actor As Object)\nIf TypeOf actor Is Collection Then\nDebug.Print 1\nEnd If\nEnd Sub\n', classes)).toEqual([]);
});
