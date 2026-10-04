import { expect, it } from 'vitest';
import { checkRuntimeMemberNotFound } from '../src/analyzer/diagnostics/rules/lateBoundMembers';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMember, VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';

function surface(name: string, members: VbaProjectClassMember[] = []): VbaProjectClassMembers {
	return { name, moduleName: name, kind: 'class', exhaustive: true, members };
}
function field(name: string): VbaProjectClassMember { return { name, moduleName: 'Class1', kind: 'property', returns: 'Long' }; }
function run(source: string, classes: VbaProjectClassMembers[]) {
	const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), out: unknown[][] = [];
	checkRuntimeMemberNotFound(source, mod, symbols, { projectClassMembers: classes }, createConditionalActivityTracker(mod), (...args) => out.push(args));
	return out;
}
function sourceFor(count: number, mode: 'new' | 'typed' | 'collection' | 'distinct') {
	return Array.from({ length: count }, (_, i) => {
		const type = mode === 'distinct' ? 'Class' + (i + 1) : 'Class1';
		const body = mode === 'collection' ? `Dim c As New Collection\nc.Add New ${type}\nDebug.Print c(1).Nope`
			: mode === 'typed' ? `Dim actor As Object, known As ${type}\nSet actor = known\nDebug.Print actor.Nope`
				: `Dim actor As Object\nSet actor = New ${type}\nDebug.Print actor.Nope`;
		return `Sub P${i}()\n${body}\nEnd Sub\n`;
	}).join('');
}
function expected(source: string, count: number, mode: 'new' | 'typed' | 'collection' | 'distinct') {
	let from = 0;
	return Array.from({ length: count }, (_, i) => {
		const type = mode === 'distinct' ? 'Class' + (i + 1) : 'Class1';
		const receiver = mode === 'collection' ? 'c(1)' : 'actor', needle = mode === 'collection' ? 'c(1).Nope' : 'Nope';
		const start = source.indexOf(needle, from), end = start + 4;
		from = start + needle.length;
		const suffix = mode === 'typed' ? ", or '91' while 'actor' is Nothing" : '';
		return ['runtimeMemberNotFound', `'${receiver}' holds a ${type} here, which has no member 'Nope'. This will raise Run-time error '438': Object doesn't support this property or method${suffix}.`, { start, end }];
	});
}
for (const mode of ['new', 'typed', 'collection', 'distinct'] as const) {
	it.each([10, 100, 1000])(`bounds class and member metadata across %i procedures via ${mode}`, count => {
		let names = 0, memberNames = 0;
		const shared = Array.from({ length: count }, (_, i) => ({ ...field('M' + i), get name() { memberNames++; return 'M' + i; } }));
		const classes = Array.from({ length: count + 1 }, (_, i) => ({
			...surface('Class' + (i + 1), mode === 'distinct' ? [{ ...field('M'), get name() { memberNames++; return 'M'; } }] : i === 0 ? shared : []),
			get name() { names++; return 'Class' + (i + 1); },
		}));
		const source = sourceFor(count, mode);
		expect(run(source, classes)).toEqual(expected(source, count, mode));
		expect(names).toBeLessThanOrEqual(count * 4 + 10);
		expect(memberNames).toBeLessThanOrEqual(count * 4 + 10);
	});
}
it('shares one projection between collection and direct consumers', () => {
	let names = 0;
	const members = Array.from({ length: 100 }, (_, i) => ({ ...field('M' + i), get name() { names++; return 'M' + i; } }));
	const source = sourceFor(100, 'new') + sourceFor(100, 'collection').replace(/Sub P/g, 'Sub Q');
	expect(run(source, [surface('Class1', members)])).toEqual([...expected(source, 100, 'new'), ...expected(source, 100, 'collection')]);
	expect(names).toBeLessThanOrEqual(410);
});
it.each(['document', 'userform', 'userType', 'enum', 'standardModule'] as const)('does not project %s metadata as a known class', kind => {
	const type = { ...surface('Class1'), kind, get members(): VbaProjectClassMember[] { throw new Error('ineligible members'); } };
	expect(run(sourceFor(1, 'new'), [type])).toEqual([]);
});
it('skips an incomplete class before choosing the first exhaustive duplicate', () => {
	const source = sourceFor(1, 'new'), incomplete = { ...surface('CLASS1'), exhaustive: false, get members(): VbaProjectClassMember[] { throw new Error('incomplete members'); } };
	expect(run(source, [incomplete, surface('Class1')])).toEqual(expected(source, 1, 'new'));
});
it('preserves first eligible duplicate selection', () => {
	const source = sourceFor(1, 'new'), first = surface('Class1'), duplicate = surface('CLASS1', [field('Nope')]);
	expect(run(source, [first, duplicate])).toEqual(expected(source, 1, 'new'));
	expect(run(source, [duplicate, first])).toEqual([]);
});
it('does not project unused class members', () => {
	const unused = { ...surface('Other'), get members(): VbaProjectClassMember[] { throw new Error('unused members'); } };
	const source = sourceFor(1, 'new');
	expect(run(source, [unused, surface('Class1')])).toEqual(expected(source, 1, 'new'));
});
it('skips project lookup for Collection and inactive defaults', () => {
	const unused = { ...surface('Class1'), get name(): string { throw new Error('unused project lookup'); } };
	const source = 'Sub Go()\nDim actor As Object\nSet actor = New Collection\nDebug.Print actor.Count\nEnd Sub\n';
	expect(run(source, [unused])).toEqual([]);
	expect(run('#If False Then\n' + sourceFor(1, 'new') + '#End If\n', [unused])).toEqual([]);
});
it('observes member replacement on the next public query', () => {
	const source = sourceFor(1, 'new'), classes = [surface('Class1')];
	expect(run(source, classes)).toEqual(expected(source, 1, 'new'));
	classes[0].members = [field('Nope')];
	expect(run(source, classes)).toEqual([]);
	classes[0].exhaustive = false;
	expect(run(source, classes)).toEqual([]);
});
it('keeps Nothing state on typed receiver overlays rather than shared class metadata', () => {
	const source = 'Sub Go()\nDim actor As Object, fresh As Object, known As Class1\nSet actor = known\nDebug.Print actor.Nope\nSet fresh = New Class1\nDebug.Print fresh.Nope\nEnd Sub\n';
	const first = source.indexOf('Nope'), second = source.indexOf('Nope', first + 4);
	expect(run(source, [surface('Class1')])).toEqual([
		['runtimeMemberNotFound', "'actor' holds a Class1 here, which has no member 'Nope'. This will raise Run-time error '438': Object doesn't support this property or method, or '91' while 'actor' is Nothing.", { start: first, end: first + 4 }],
		['runtimeMemberNotFound', "'fresh' holds a Class1 here, which has no member 'Nope'. This will raise Run-time error '438': Object doesn't support this property or method.", { start: second, end: second + 4 }],
	]);
});
