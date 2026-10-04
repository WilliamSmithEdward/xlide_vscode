import { describe, expect, it } from 'vitest';
import { checkRuntimeMemberNotFound } from '../src/analyzer/diagnostics/rules/lateBoundMembers';
import { getExcelObjectModel, type HostMember } from '../src/analyzer/host/excelObjectModel';
import { getHostMembers } from '../src/analyzer/host/hostModel';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { buildModuleSymbols } from '../src/analyzer/symbols/buildModuleSymbols';

function fixture(members: HostMember[], n = 2, collection = 'Worksheets') {
	const base = getExcelObjectModel();
	const model = { ...base, types: { ...base.types, 'Excel.Worksheets': { ...base.types['Excel.Worksheets'], members } } };
	const source = ['Option Explicit', 'Sub Go()', 'Dim cell As Range', ...Array<string>(n).fill(`For Each cell In ${collection}\nDebug.Print cell.Address\nNext cell`), 'End Sub', ''].join('\n');
	const mod = parseModule(source);
	const symbols = buildModuleSymbols('M', 'standard', source, { parsedModule: mod });
	const run = () => {
		const out: unknown[][] = [];
		checkRuntimeMemberNotFound(source, mod, symbols, { model, parsedModule: mod, memberSurfaceCache: new Map() }, undefined, (...v) => out.push(v));
		return out;
	};
	let from = 0;
	const warnings = Array.from({ length: n }, () => {
		const start = source.indexOf(collection, from);
		from = start + collection.length;
		return ['assignmentObjectTypeMismatch', `For Each Sets the items of '${collection}', each a Worksheet, into 'cell', a Range. This will raise Run-time error '13': Type mismatch.`, { start, end: from }];
	});
	return { model, run, warnings };
}
const item = (returns: string, name = 'Item'): HostMember => ({ name, kind: 'property', returns });
describe('host collection Item queries', () => {
	for (const present of [true, false]) for (const n of [10, 100, 1000]) {
		it(`scans ${n} members once for ${n} loops with Item ${present ? 'last' : 'absent'}`, () => {
			const members = Array.from({ length: n }, (_, i) => item(i === n - 1 ? 'Excel.Worksheet' : 'Long', present && i === n - 1 ? 'Item' : `Member${i}`));
			const f = fixture(members, n);
			getHostMembers('Excel.Worksheets', f.model);
			let reads = 0, scanning = false;
			const originalFind = Array.prototype.find;
			const descriptors = members.map(m => Object.getOwnPropertyDescriptor(m, 'name')!);
			try {
				Array.prototype.find = function (callback: any, thisArg?: any): any {
					const prior = scanning;
					scanning = (new Error().stack ?? '').includes('hostElementType');
					try { return originalFind.call(this, callback, thisArg); } finally { scanning = prior; }
				};
				members.forEach(m => {
					const name = m.name;
					Object.defineProperty(m, 'name', { configurable: true, get() {
						if (scanning) reads++;
						return name;
					} });
				});
				expect(f.run()).toEqual(present ? f.warnings : []);
				expect(reads).toBe(n);
			} finally {
				Array.prototype.find = originalFind;
				members.forEach((m, i) => Object.defineProperty(m, 'name', descriptors[i]));
			}
		});
	}
	it.each([
		['wrong case before exact case', [item('Excel.Range', 'item'), item('Excel.Worksheet')], true],
		['only wrong case', [item('Excel.Worksheet', 'item')], false],
		['first exact match', [item('Excel.Range'), item('Excel.Worksheet')], false],
		['event excluded', [{ ...item('Excel.Worksheet'), kind: 'event' }], false],
		['unqualified Object', [item('Object')], false],
		['unqualified String', [item('String')], false],
		['compatible Range', [item('Excel.Range')], false],
	] as const)('preserves %s', (_label, members, warns) => {
		const f = fixture([...members]);
		expect(f.run()).toEqual(warns ? f.warnings : []);
	});
	it('sees renamed members between invocations, including previously absent Item', () => {
		const member = item('Excel.Worksheet', 'Other');
		const f = fixture([member]);
		expect(f.run()).toEqual([]);
		member.name = 'Item';
		expect(f.run()).toEqual(f.warnings);
		member.name = 'Other';
		expect(f.run()).toEqual([]);
	});
	it('sees changed returns between invocations', () => {
		const member = item('Excel.Range');
		const f = fixture([member]);
		expect(f.run()).toEqual([]);
		member.returns = 'Excel.Worksheet';
		expect(f.run()).toEqual(f.warnings);
	});
	it('isolates models and cached absence', () => {
		const absent = fixture([]), present = fixture([item('Excel.Worksheet')]);
		expect(absent.run()).toEqual([]);
		expect(present.run()).toEqual(present.warnings);
		expect(absent.run()).toEqual([]);
	});
	it('preserves Range collection handling', () => {
		const f = fixture([item('Excel.Worksheet')], 2, 'Range("A1:A2")');
		expect(f.run()).toEqual([]);
	});
});
