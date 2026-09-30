// Diagnostics tests: Excel objects read as values (issue #221). Each case was
// measured in 64-bit Excel 16.0 (build 20326, 2026-09-30) with the variable
// declared as the type and Set first.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getExcelObjectModel } from '../../src/analyzer/host/excelObjectModel';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const hostModel = getExcelObjectModel();

function wrap(type: string, ...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim o As ${type}\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function diags(type: string, ...lines: string[]) {
	return analyzeModule(wrap(type, ...lines), { hostModel });
}

describe('Excel objects with no default member raise 438 (issue #221)', () => {
	it.each(['Workbook', 'Font', 'Interior', 'Validation', 'Window', 'PageSetup', 'Border', 'Shape', 'Hyperlink', 'Chart', 'Worksheet'])('flags %s read five ways', (type) => {
		for (const read of ['Dim v As Variant', 'Dim s As String']) {
			const src = wrap(type, read, `${read.includes('String') ? 's' : 'v'} = o`);
			expectDiagnostic(src, analyzeModule(src, { hostModel }), 'object-default-value', { span: 'o', message: "'438'" });
		}
		for (const line of ['Main = o & "x"', 'Debug.Print o', 'If o = 7 Then Main = 2']) {
			const src = wrap(type, line);
			expect(byCode(analyzeModule(src, { hostModel }), 'object-default-value')).toHaveLength(1);
		}
	});

	it.each(['Range', 'Application', 'Style', 'Name', 'ListObject'])('stays quiet on %s, whose default member gives a value', (type) => {
		const hits = diags(type, 'Dim v As Variant', 'v = o', 'Debug.Print o', 'Main = o & "x"');
		expect([...byCode(hits, 'object-default-value'), ...byCode(hits, 'collection-operand')]).toHaveLength(0);
	});
});

describe('Excel collections whose default member is their Item (issue #221)', () => {
	it.each(['Hyperlinks', 'Areas', 'Borders', 'Windows', 'Workbooks', 'Shapes', 'Comments', 'ListObjects'])('%s read as a value raises 450, and does not compile into a String or with an operator', (type) => {
		const read = wrap(type, 'Dim v As Variant', 'v = o', 'Debug.Print o');
		const readHits = byCode(analyzeModule(read, { hostModel }), 'object-default-value');
		expect(readHits).toHaveLength(2);
		expect(readHits[0].message).toContain("'450'");
		const typed = wrap(type, 'Dim s As String', 's = o', 'Main = o & "x"', 'If o = 7 Then Main = 2');
		const typedHits = analyzeModule(typed, { hostModel });
		expect(byCode(typedHits, 'collection-operand')).toHaveLength(3);
		expect(byCode(typedHits, 'object-default-value')).toHaveLength(0);
	});

	it('does not judge Worksheets, Sheets or Names', () => {
		for (const type of ['Worksheets', 'Sheets', 'Names']) {
			const hits = diags(type, 'Dim v As Variant', 'v = o');
			expect(byCode(hits, 'object-default-value')).toHaveLength(0);
		}
	});

	it('refuses a Let into the variable with "Invalid use of property"', () => {
		const src = wrap('Hyperlinks', 'o = 5');
		expectDiagnostic(src, analyzeModule(src, { hostModel }), 'set-required', { message: 'Invalid use of property' });
		const collection = `Option Explicit\nSub P()\n    Dim c As New Collection\n    c = 5\nEnd Sub\n`;
		expectDiagnostic(collection, analyzeModule(collection, { hostModel }), 'set-required', { message: 'Argument not optional' });
	});
});

describe('a Collection read into a typed value (issue #221)', () => {
	it('is a compile error, not the run-time 450', () => {
		const src = `Option Explicit\nFunction Main() As Variant\n    Dim c As New Collection, s As String, v As Variant\n    s = c\n    v = c\nEnd Function\n`;
		const hits = analyzeModule(src);
		expectDiagnostic(src, hits, 'collection-operand', { message: "for 's'" });
		expectDiagnostic(src, hits, 'object-default-value', { message: "'450'" });
	});
});
