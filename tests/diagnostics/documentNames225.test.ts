// Diagnostics tests: a document module's own name as an assignment target
// (issue #225). Each case was measured in 64-bit Excel 16.0 (build 20326,
// 2026-09-30) from a standard module.

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

function errors(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
	return analyzeProjectModule(src, [
		{ moduleName: 'Module1', source: src },
		{ moduleName: 'Sheet1', source: 'Option Explicit\n', type: 'document' },
		{ moduleName: 'ThisWorkbook', source: 'Option Explicit\n', type: 'document' },
		{ moduleName: 'Class1', source: 'Option Explicit\n', type: 'class' },
	], 'Module1').filter((d) => d.severity === 'error').map((d) => `${d.code}: ${d.message}`);
}

describe('assigning to a document module name (issue #225)', () => {
	it.each(['Set Sheet1 = Nothing', 'Set ThisWorkbook = Nothing', 'Set Sheet1 = Sheet1'])('flags %s at compile time', (line) => {
		const hits = errors(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^set-requires-object: .*Invalid use of property/);
	});

	it.each(['Sheet1 = 5', 'ThisWorkbook = 5', 'Let Sheet1 = 5', 'Sheet1 = Sheet1'])('flags %s as error 438', (line) => {
		const hits = errors(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^set-required: .*'438'/);
	});

	it.each([
		'Sheet1.Name = Sheet1.Name',
		'Dim Sheet1 As Long: Sheet1 = 5',
		'Dim Sheet1 As Object: Set Sheet1 = Nothing',
	])('stays quiet on %s', (line) => {
		expect(errors(line)).toHaveLength(0);
	});

	it('leaves a class module name to the other rules', () => {
		expect(errors('Class1 = 5').filter((hit) => hit.startsWith('set-required'))).toHaveLength(0);
	});

	it('keeps New on a document a compile error, as a full compile refuses it', () => {
		// Run mode compiles only what runs, so `Dim s As New Sheet1` appears to
		// work there; Debug > Compile refuses it, "Invalid use of New keyword".
		const hits = errors('Dim s As New Sheet1', 'Main = s.Name');
		expect(hits.some((hit) => hit.startsWith('invalid-new-type-name: '))).toBe(true);
	});
});
