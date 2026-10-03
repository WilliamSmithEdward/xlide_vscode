import { describe, expect, it } from 'vitest';
import { resolveTypeCompletions, resolveTypeName } from '../src/analyzer/completion/typeCompletion';
import type { HostObjectModel, HostType } from '../src/analyzer/host/excelObjectModel';

function modelWith(types: Record<string, HostType>): HostObjectModel {
	return { source: 'test', types, aliases: {}, globals: {} };
}
const type = (name: string): HostType => ({ name, members: [] });

describe('qualified type metadata index', () => {
	it('enumerates metadata once across lookups, misses, completions and qualifiers', () => {
		let scans = 0;
		const types = new Proxy({ 'Excel.Range': type('Range'), 'Word.Range': type('Range') }, {
			ownKeys(target) { scans++; return Reflect.ownKeys(target); },
		});
		const model = modelWith(types);
		for (let i = 0; i < 100; i++) {
			expect(resolveTypeName('excel.RANGE', { model })?.detail).toBe('Excel type');
			expect(resolveTypeName('Word.Range', { model })?.detail).toBe('Word type');
			expect(resolveTypeName('Missing.Range', { model })).toBeUndefined();
			expect(resolveTypeName('Excel.Missing', { model })).toBeUndefined();
		}
		expect(resolveTypeCompletions('Dim x As Excel.', 15, { model }).map(c => c.name)).toEqual(['Range']);
		expect(resolveTypeCompletions('Dim x As Exc', 12, { model }).map(c => c.name)).toContain('Excel');
		// Unqualified completion also enumerates types for its own host ordering.
		expect(scans).toBe(2);
	});

	it('keeps first casing, duplicates, order and library qualification rules', () => {
		const model = modelWith({
			'ExCeL.Nested.Type': type('Nested'), 'EXCEL.Range': type('First'),
			'excel.range': type('Second'), 'Excel.': type('Empty'),
			'Excel.Worksheet': type('Worksheet'), Bare: type('Bare'),
		});
		expect(resolveTypeName('excel.range', { model })).toMatchObject({ name: 'Range', detail: 'ExCeL type' });
		expect(resolveTypeName('Excel.Nested.Type', { model })).toBeUndefined();
		expect(resolveTypeCompletions('Dim x As Excel.', 15, { model }).map(c => c.name)).toEqual(['Range', 'Worksheet']);
	});

	it('preserves project and stdole precedence over the host library', () => {
		const model = modelWith({ 'Excel.Range': type('Range'), 'stdole.IUnknown': type('IUnknown') });
		const projectTypes = [{ name: 'Range', kind: 'class' as const, moduleName: 'Excel' }];
		expect(resolveTypeName('eXcel.Range', { model, projectTypes })).toMatchObject({ kind: 'class', moduleName: 'eXcel' });
		expect(resolveTypeName('Excel.Range', { model, projectTypes: [...projectTypes, ...projectTypes] })?.kind).toBe('ambiguous');
		expect(resolveTypeName('stdole.IUnknown', { model })?.kind).toBe('external');
		expect(resolveTypeName('Excel.Range', { model })?.kind).toBe('host');
	});

	it('isolates replacement models and caller mutations', () => {
		const first = modelWith({ 'Excel.Range': type('Range') });
		const second = modelWith({ 'Excel.Worksheet': type('Worksheet') });
		const candidate = resolveTypeName('Excel.Range', { model: first })!;
		candidate.name = 'Changed';
		const completion = resolveTypeCompletions('Dim x As Excel.', 15, { model: first });
		expect(completion[0].name).toBe('Range');
		completion[0].detail = 'Changed';
		expect(resolveTypeName('Excel.Range', { model: first })?.detail).toBe('Excel type');
		expect(resolveTypeName('Excel.Range', { model: second })).toBeUndefined();
		expect(resolveTypeName('Excel.Worksheet', { model: second })?.name).toBe('Worksheet');
		expect(resolveTypeName('Excel.Range', { model: first })?.name).toBe('Range');
	});
});
