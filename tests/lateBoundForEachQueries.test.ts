import { expect, it } from 'vitest';
import { checkRuntimeMemberNotFound } from '../src/analyzer/diagnostics/rules/lateBoundMembers';
import { createConditionalActivityTracker } from '../src/analyzer/conditional/conditionalCompilation';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';
import type { VbaProjectClassMembers } from '../src/analyzer/symbols/symbolModel';
import type { HostObjectModel } from '../src/analyzer/host/excelObjectModel';
import { getExcelObjectModel } from '../src/analyzer/host/excelObjectModel';
import { createObjectTypeImplementationLookup, implementsObjectType, type KnownObjectAssignmentType } from '../src/analyzer/diagnostics/typeInference';

function surface(name: string, implementsNames: string[] = []): VbaProjectClassMembers {
	return { name, moduleName: name, kind: 'class', exhaustive: true, members: [], implements: implementsNames };
}
function run(source: string, classes: VbaProjectClassMembers[], model?: HostObjectModel) {
	const mod = parseModule(source), symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod }), out: unknown[][] = [];
	checkRuntimeMemberNotFound(source, mod, symbols, { projectClassMembers: classes, model }, createConditionalActivityTracker(mod), (...args) => out.push(args));
	return out;
}
function heldSource(items: string[], expected = 'Class1', body = 'Debug.Print 1') {
	return `Sub Go()\nDim c As New Collection\nDim actor As ${expected}\n${items.map(item => 'c.Add ' + item).join('\n')}\nFor Each actor In c\n${body}\nNext\nEnd Sub\n`;
}
function hostSource(count: number) {
	return Array.from({ length: count }, (_, i) => `Sub P${i}()\nDim actor As Class1\nFor Each actor In Worksheets\nDebug.Print 1\nNext\nEnd Sub\n`).join('');
}
function expectedHost(source: string, count: number) {
	let from = 0;
	return Array.from({ length: count }, () => {
		const start = source.indexOf('Worksheets', from); from = start + 10;
		return ['assignmentObjectTypeMismatch', "For Each Sets the items of 'Worksheets', each a Worksheet, into 'actor', a Class1. This will raise Run-time error '13': Type mismatch.", { start, end: from }];
	});
}
function expectedItem(source: string, position: number, type: string | undefined) {
	const start = source.indexOf('In c') + 3;
	return [['assignmentObjectTypeMismatch', type
		? `For Each Sets each item of 'c' into 'actor', a Class1, and item ${position} is a ${type}. This will raise Run-time error '13': Type mismatch.`
		: `For Each Sets each item of 'c' into 'actor', a Class1, and item ${position} is a number or string, no object. This will raise Run-time error '424': Object required.`, { start, end: start + 1 }]];
}
for (const mode of ['held', 'host', 'shared', 'distinct-shared'] as const) {
	it.each([10, 100, 1000])(`bounds all For Each type/interface metadata at %i via ${mode}`, count => {
		let names = 0, interfaces = 0;
		const classes = Array.from({ length: count + 2 }, (_, i) => ({
			...surface('Class' + (i + 1)),
			get name() { names++; return 'Class' + (i + 1); },
			get implements() { interfaces++; return i === count + 1 ? mode === 'shared' ? ['Class1', 'Class2'] : mode === 'distinct-shared' ? Array.from({ length: count + 1 }, (_, n) => 'Class' + (n + 1)) : [] : []; },
		}));
		const source = mode === 'host' ? hostSource(count) : heldSource(Array.from({ length: count }, (_, i) => 'New ' + (mode === 'shared' ? 'Class2' : mode === 'distinct-shared' ? 'Class' + (i + 2) : 'Class1')));
		expect(run(source, classes)).toEqual(mode === 'host' ? expectedHost(source, count) : []);
		expect(names).toBeLessThanOrEqual(count * 4 + 16);
		expect(interfaces).toBeLessThanOrEqual(count * 4 + 16);
	});
}
it('shares a single resolver between held and host queries across procedures', () => {
	let names = 0;
	const classes = Array.from({ length: 101 }, (_, i) => ({ ...surface('Class' + (i + 1)), get name() { names++; return 'Class' + (i + 1); } }));
	const source = hostSource(100) + heldSource(Array.from({ length: 100 }, () => 'New Class1'));
	expect(run(source, classes)).toEqual(expectedHost(source, 100));
	expect(names).toBeLessThanOrEqual(410);
});
it.each(['class', 'document', 'userform', 'userType', 'enum', 'standardModule'] as const)('preserves shared implementers from %s metadata', kind => {
	const bridge = { ...surface('Bridge', ['cLaSs1', 'CLASS2']), kind };
	expect(run(heldSource(['New Class2']), [surface('Class1'), surface('Class2'), bridge])).toEqual([]);
});
it('preserves direct interface compatibility in both directions', () => {
	const source = heldSource(['New Class2']);
	expect(run(source, [surface('Class1'), surface('Class2', ['Class1'])])).toEqual([]);
	expect(run(source, [surface('Class1', ['Class2']), surface('Class2')])).toEqual([]);
});
it('preserves the first incompatible item and early loop exit', () => {
	const source = heldSource(['New Class1', 'New Class2', 'New Class3']);
	expect(run(source, [surface('Class1'), surface('Class2'), surface('Class3')])).toEqual(expectedItem(source, 2, 'Class2'));
	expect(run(heldSource(['New Class1', 'New Class2'], 'Class1', 'Exit For'), [surface('Class1'), surface('Class2')])).toEqual([]);
});
it('preserves scalar item refusal and generic Variant controls', () => {
	const source = heldSource(['New Class1', '1']);
	expect(run(source, [surface('Class1')])).toEqual(expectedItem(source, 2, undefined));
	expect(run(heldSource(['1'], 'Variant'), [surface('Class1')])).toEqual([]);
});
it.each(['userType', 'enum', 'standardModule'] as const)('preserves excluded %s duplicates', kind => {
	const source = hostSource(1), classes = [surface('Class1'), { ...surface('CLASS1'), kind }];
	expect(run(source, classes)).toEqual(expectedHost(source, 1));
});
it('observes new eligible ambiguity on the next public query', () => {
	const source = hostSource(1), classes = [surface('Class1')];
	expect(run(source, classes)).toEqual(expectedHost(source, 1));
	classes.push(surface('CLASS1'));
	expect(run(source, classes)).toEqual([]);
});
it('observes changed Implements on the next public query', () => {
	const source = heldSource(['New Class2']), classes = [surface('Class1'), surface('Class2')];
	expect(run(source, classes)).toEqual(expectedItem(source, 1, 'Class2'));
	classes[1].implements = ['Class1'];
	expect(run(source, classes)).toEqual([]);
});
it('keeps default and explicit Excel host resolution equivalent', () => {
	const source = hostSource(1), classes = [surface('Class1')];
	expect(run(source, classes)).toEqual(expectedHost(source, 1));
	expect(run(source, classes, getExcelObjectModel())).toEqual(expectedHost(source, 1));
});
it('honors an explicit model for host collection element and control types', () => {
	const model: HostObjectModel = { source: 'test', hostName: 'Test', types: { 'Test.Items': { displayName: 'Items', members: [{ name: 'Item', kind: 'property', returns: 'Test.Element' }] }, 'Test.Element': { displayName: 'Element', members: [] } }, aliases: { class1: 'Test.Element' }, globals: { Worksheets: 'Test.Items' } };
	expect(run(hostSource(1), [surface('Class1')], model)).toEqual([]);
});
it('does not resolve inactive loops or generic controls against project metadata', () => {
	const unused = { ...surface('unused'), get name(): string { throw new Error('unused metadata'); } };
	expect(run('#If False Then\n' + hostSource(1) + '#End If\n', [unused])).toEqual([]);
	expect(run(heldSource(['New Collection'], 'Object'), [unused])).toEqual([]);
});

for (const mode of ['repeated-direct', 'distinct-reverse', 'generic-collection'] as const) {
 it.each([10, 100, 1000])('bounds implemented-name iteration at %i for ' + mode, count => {
  let reads = 0;
  const names = mode === 'distinct-reverse' ? Array.from({ length: count }, (_, i) => 'Class' + (i + 2)) : Array.from({ length: count }, (_, i) => i === count - 1 ? mode === 'generic-collection' ? 'Collection' : 'Class1' : 'Extra' + i);
  const implemented = new Proxy(names, { get(target, key, receiver) { if (typeof key === 'string' && /^(0|[1-9]\d*)$/.test(key)) reads++; return Reflect.get(target, key, receiver); } });
  const classes = Array.from({ length: count + 2 }, (_, i) => surface('Class' + (i + 1), (mode === 'distinct-reverse' ? i === 0 : i === 1) ? implemented : []));
  const source = heldSource(Array.from({ length: count }, (_, i) => 'New ' + (mode === 'distinct-reverse' ? 'Class' + (i + 2) : 'Class2')), mode === 'generic-collection' ? 'Collection' : 'Class1');
  expect(run(source, classes)).toEqual([]);
  expect(reads).toBeLessThanOrEqual(count * 3 + 10);
 });
}
it.each([
 ['Worksheet', 'Excel.Worksheet', 'host', ['worksheet'], true],
 ['Excel.Worksheet', 'excel.worksheet', 'host', ['Excel.Worksheet'], true],
 ['Worksheet', 'worksheet', 'project', ['Excel.Worksheet'], false],
 ['FoO', 'foo', 'project', ['fOo', 'FOO'], true],
 ['Foo', 'foo', 'project', ['Other'], false],
 ['Collection', 'collection', 'generic', ['collection'], true],
 ['Collection', 'collection', 'generic', ['Excel.Collection'], false],
] as const)('preserves direct implementation aliases for %s / %s', (display, key, kind, implemented, expectedResult) => {
 const actual: Extract<KnownObjectAssignmentType, { kind: 'project' }> = { kind: 'project', display: 'Actual', key: 'actual', implements: implemented };
 const target = { kind, display, key, ...(kind === 'project' ? { implements: [] } : {}) } as KnownObjectAssignmentType;
 expect(implementsObjectType(actual, target)).toBe(expectedResult);
 expect(createObjectTypeImplementationLookup()(actual, target)).toBe(expectedResult);
});
it('preserves sparse implementation arrays and refreshes a changed list in a new query', () => {
 const implemented: string[] = []; implemented[2] = 'Other';
 const actual: Extract<KnownObjectAssignmentType, { kind: 'project' }> = { kind: 'project', display: 'Actual', key: 'actual', implements: implemented };
 const expected: KnownObjectAssignmentType = { kind: 'host', display: 'Worksheet', key: 'excel.worksheet' };
 expect(createObjectTypeImplementationLookup()(actual, expected)).toBe(false);
 implemented[1] = 'WORKSHEET';
 expect(createObjectTypeImplementationLookup()(actual, expected)).toBe(true);
});
