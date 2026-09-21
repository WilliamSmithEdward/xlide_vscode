// Which Excel types can prove a member absent.
//
// A COM interface marked NONEXTENSIBLE gains no members at run time, so VBA
// refuses a name that is not on it while compiling. Without that flag the
// object is extensible and the name is deferred to IDispatch, which is how
// `Application.Match` - a worksheet function on no interface in the library
// at all - compiles and runs. Reporting it was the bug behind 10.4.2.
//
// Every expectation here was measured against the VBE and kept as an oracle
// case (syntax_corpus/oracle/vbe_oracle_cases.json), and the flag agreed with
// the VBE on all seven receivers tried.

import { describe, expect, it } from 'vitest';
import { analyzeModule } from '../src/analyzer/diagnostics/analyzeModule';
import {
	EXCEL_CLOSED_TYPE_NAMES,
	hostTypeResolvesWhenCompiling,
} from '../src/analyzer/host/typeExtensibility';

const found = (body: string): string[] => analyzeModule(
	['Option Explicit', '', 'Public Sub P(ws As Worksheet, rng As Range, wb As Workbook)',
		...body.split('\n').map((line) => `    ${line}`), 'End Sub'].join('\r\n'),
	{ moduleName: 'M', moduleKind: 'standard' },
).filter((one) => one.code === 'member-not-found').map((one) => one.message);

describe('a worksheet function called off Application', () => {
	it('is not reported, because Excel resolves it when the code runs', () => {
		// The report: a red squiggle on code that compiles and runs.
		expect(found('Dim v As Variant\nv = Application.Match("a", rng, 0)')).toEqual([]);
	});

	it('is not reported for any of the family, which all live on WorksheetFunction', () => {
		for (const call of ['Application.VLookup("a", rng, 2, False)',
			'Application.Sum(rng)', 'Application.CountA(rng)',
			'Application.Index(rng, 1)', 'Application.Transpose(rng)']) {
			expect(found(`Dim v As Variant\nv = ${call}`), call).toEqual([]);
		}
	});

	it('is not reported through WorksheetFunction either, which really has it', () => {
		expect(found('Dim v As Variant\nv = Application.WorksheetFunction.Match("a", rng, 0)'))
			.toEqual([]);
	});
});

describe('an unknown member on an extensible type', () => {
	// The VBE accepts every one of these. Excel's Application, Range, Workbook
	// and Font interfaces carry no NONEXTENSIBLE flag.
	it('says nothing on Application', () => {
		expect(found('Application.NoSuchMemberXyz')).toEqual([]);
		expect(found('Application.NoSuchPropertyXyz = 1')).toEqual([]);
	});

	it('says nothing on Range, declared or called for', () => {
		expect(found('rng.NoSuchMemberXyz')).toEqual([]);
		expect(found('Range("A1").NoSuchMemberXyz')).toEqual([]);
		expect(found('ActiveCell.NoSuchMemberXyz')).toEqual([]);
	});

	it('says nothing on Workbook or Font', () => {
		expect(found('wb.NoSuchMemberXyz')).toEqual([]);
		expect(found('rng.Font.NoSuchMemberXyz')).toEqual([]);
	});
});

describe('an unknown member on a closed type', () => {
	// These the VBE refuses, so XLIDE still does.
	it('is reported on Worksheet', () => {
		expect(found('ws.NoSuchMemberXyz'))
			.toEqual(["Method or data member not found: 'Excel.Worksheet.NoSuchMemberXyz'."]);
	});

	it('is reported on Chart', () => {
		expect(found('Dim c As Chart\nc.NoSuchMemberXyz'))
			.toEqual(["Method or data member not found: 'Excel.Chart.NoSuchMemberXyz'."]);
	});

	it('is reported on the Sheets and Workbooks collections', () => {
		expect(found('Sheets.NoSuchMemberXyz'))
			.toEqual(["Method or data member not found: 'Excel.Sheets.NoSuchMemberXyz'."]);
		expect(found('Workbooks.NoSuchMemberXyz'))
			.toEqual(["Method or data member not found: 'Excel.Workbooks.NoSuchMemberXyz'."]);
	});

	it('leaves the real members of those types alone', () => {
		expect(found('ws.Calculate\nSheets.Add\nWorkbooks.Open "Book.xlsx"')).toEqual([]);
	});
});

describe('the flag itself', () => {
	it('answers for a qualified key and a bare name, Excel being the default', () => {
		expect(hostTypeResolvesWhenCompiling('Excel.Worksheet')).toBe(true);
		expect(hostTypeResolvesWhenCompiling('Worksheet')).toBe(true);
		expect(hostTypeResolvesWhenCompiling('Excel.Range')).toBe(false);
		expect(hostTypeResolvesWhenCompiling('Range')).toBe(false);
	});

	it('leaves another host to the answer it had, none having been measured', () => {
		// Word and PowerPoint are closed almost throughout, and neither has a
		// type marked exhaustive, so this changes nothing for them today.
		expect(hostTypeResolvesWhenCompiling('Word.Range')).toBe(true);
		expect(hostTypeResolvesWhenCompiling('PowerPoint.Shape')).toBe(true);
	});

	it('names the types the type library marks NONEXTENSIBLE, and only those', () => {
		// Measured with LoadRegTypeLib over Excel 16: 27 of its 747 interfaces
		// carry the flag. The two a user meets are Worksheet and Chart.
		expect(EXCEL_CLOSED_TYPE_NAMES).toContain('Worksheet');
		expect(EXCEL_CLOSED_TYPE_NAMES).toContain('Chart');
		expect(EXCEL_CLOSED_TYPE_NAMES).toContain('Sheets');
		expect(EXCEL_CLOSED_TYPE_NAMES).toContain('Workbooks');
		expect(EXCEL_CLOSED_TYPE_NAMES).not.toContain('Application');
		expect(EXCEL_CLOSED_TYPE_NAMES).not.toContain('Range');
		expect(EXCEL_CLOSED_TYPE_NAMES).not.toContain('Workbook');
		expect(EXCEL_CLOSED_TYPE_NAMES).toHaveLength(27);
	});
});
