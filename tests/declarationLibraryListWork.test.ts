import {beforeEach, describe, expect, it, vi} from 'vitest';
import {parseModule} from '../src/analyzer/parser/parseModule';
import {checkInvalidAsTypeNames} from '../src/analyzer/diagnostics/rules/declarations';
import {libraryTypeNames} from '../src/analyzer/host/libraryTypeNames';
import {createConditionalActivityTracker} from '../src/analyzer/conditional/conditionalCompilation';

vi.mock('../src/analyzer/host/libraryTypeNames', async importOriginal => {
	const actual = await importOriginal<typeof import('../src/analyzer/host/libraryTypeNames')>();
	return {...actual, libraryTypeNames: vi.fn(actual.libraryTypeNames)};
});
beforeEach(() => {vi.mocked(libraryTypeNames).mockClear();});
const standard = ['VBA', 'Excel', 'stdole', 'Office'];

function run(source: string, references: readonly string[] | undefined) {
	const rows: unknown[] = [];
	checkInvalidAsTypeNames(source, parseModule(source), undefined, {referencedLibraries: references}, (rule, message, span) => {rows.push({rule, message, span});});
	return rows;
}
function expected(source: string, names: readonly string[]) {
	return names.map(name => ({rule: 'invalidAsTypeName', message: `No type of this project and none of the libraries it references is named '${name}'. This is a VBE compile error: User-defined type not defined.`, span: {start: source.indexOf(name), end: source.indexOf(name) + name.length}}));
}

describe('declaration referenced library list work', () => {
	it.each([10, 1000].flatMap(count => ['\n', '\r\n', '\r'].map(eol => ({count, eol}))))('resolves libraries once for $count type references with EOL $eol', ({count, eol}) => {
		const names = Array.from({length: count}, (_, i) => 'NotKnownType_' + i);
		const source = names.map((name, i) => `Dim value${i} As ${name}`).join(eol);
		expect(run(source, Object.freeze([...standard]))).toEqual(expected(source, names));
		expect(vi.mocked(libraryTypeNames).mock.calls).toEqual(standard.map(name => [name]));
	});

	it('stays lazy when no declaration needs reference-library knowledge', () => {
		const source = 'Dim number As Long\nDim sheet As Worksheet\nDim text As String\n';
		expect(run(source, standard)).toEqual([]);
		expect(libraryTypeNames).not.toHaveBeenCalled();
	});

	it('keeps absent, empty and unmodelled reference lists quiet', () => {
		const source = 'Dim value As NotKnownType\n';
		for (const references of [undefined, [], [...standard, 'ADODB'], [...standard, 'constructor']]) {
			expect(run(source, references)).toEqual([]);
		}
	});

	it('reads a caller reference list again after it changes between analyses', () => {
		const source = 'Dim value As DataObject\n';
		const references = [...standard];
		expect(run(source, references)).toEqual(expected(source, ['DataObject']));
		references.push('MSForms');
		expect(run(source, references)).toEqual([]);
		references.pop();
		expect(run(source, references)).toEqual(expected(source, ['DataObject']));
	});
});

it.each(['\n', '\r\n', '\r'])('does not resolve libraries for inactive unknown types with EOL %j', eol => {
	const source = ['#If False Then', 'Dim value As NotKnownType', '#End If', 'Dim active As Long'].join(eol);
	const mod = parseModule(source), rows: unknown[] = [];
	checkInvalidAsTypeNames(source, mod, createConditionalActivityTracker(mod), {referencedLibraries: standard}, (...args) => {rows.push(args);});
	expect(rows).toEqual([]);
	expect(libraryTypeNames).not.toHaveBeenCalled();
});
