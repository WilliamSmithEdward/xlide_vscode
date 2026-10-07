import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const wrap = (body: string, extra = '') => `Option Explicit\nSub Main()\n${body}\nEnd Sub\n${extra}`;
const rangeErrors = (source: string) => analyzeModule(source).filter(d => d.code === 'host-argument-out-of-range');
const deletedErrors = (source: string) => analyzeModule(source).filter(d => d.code === 'object-used-after-delete');

describe('uncertain runtime object state', () => {
	it('does not assume protected cells are locked', () => {
		expect(rangeErrors(wrap('Dim ws As Worksheet\nSet ws = ActiveSheet\nws.Protect\nws.Range("A1").Value = 1'))).toEqual([]);
	});

	it.each(['ActiveSheet.Unprotect "pw"', 'Dim alias As Worksheet\nSet alias = ws\nalias.Unprotect "pw"'])('forgets protection after unprotecting through another reference: %s', change => {
		expect(rangeErrors(wrap(`Dim ws As Worksheet\nSet ws = ActiveSheet\nws.Protect "pw"\n${change}\nws.Unprotect "nope"\nws.Range("A1").Value = 1`))).toEqual([]);
	});

	it('forgets protection after an assignment invoking a helper', () => {
		const source = wrap('Dim ws As Worksheet, ok As Boolean\nSet ws = ActiveSheet\nws.Protect "pw"\nok = Restore(ws)\nws.Unprotect "nope"', 'Function Restore(ws As Worksheet) As Boolean\nws.Unprotect "pw"\nRestore = True\nEnd Function\n');
		expect(rangeErrors(source)).toEqual([]);
	});

	it('forgets protection after a Set expression invoking a helper', () => {
		const source = wrap('Dim ws As Worksheet, r As Range\nSet ws = ActiveSheet\nws.Protect "pw"\nSet r = Restore(ws)\nws.Unprotect "nope"', 'Function Restore(ws As Worksheet) As Range\nws.Unprotect "pw"\nSet Restore = ws.Range("A1")\nEnd Function\n');
		expect(rangeErrors(source)).toEqual([]);
	});

	it('forgets protection before evaluating a condition or operation argument', () => {
		const helper = 'Function Restore(wb As Workbook) As String\nwb.Unprotect "pw"\nRestore = "Renamed"\nEnd Function\n';
		for (const operation of ['If Restore(wb) = "Renamed" Then\nws.Name = "Renamed"\nEnd If', 'ws.Name = Restore(wb)']) {
			const source = wrap(`Dim wb As Workbook, ws As Worksheet\nSet wb = ThisWorkbook\nSet ws = wb.Worksheets(1)\nwb.Protect "pw", True\n${operation}`, helper);
			expect(rangeErrors(source)).toEqual([]);
		}
	});

	it('does not assume Delete succeeded when execution resumes after an error', () => {
		const source = wrap('Dim ws As Worksheet\nSet ws = ThisWorkbook.Worksheets(1)\nThisWorkbook.Protect Structure:=True\nOn Error Resume Next\nws.Delete\nOn Error GoTo 0\nDebug.Print ws.Name');
		expect(deletedErrors(source)).toEqual([]);
	});

	it('does not assume Delete or Close was accepted', () => {
		expect(deletedErrors(wrap('Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Delete\nDebug.Print ws.Name'))).toEqual([]);
		expect(deletedErrors(wrap('Dim wb As Workbook\nSet wb = Workbooks.Add\nwb.Close False\nDebug.Print wb.Name'))).toEqual([]);
	});

	it('does not retain active-sheet identity across a function call', () => {
		const source = wrap('Dim w1 As Worksheet, w2 As Worksheet, ok As Boolean\nSet w1 = ActiveSheet\nSet w2 = Worksheets.Add\nok = Restore(w1)\nDebug.Print w1.Range(Cells(1,1), Cells(2,2)).Count', 'Function Restore(ws As Worksheet) As Boolean\nws.Activate\nRestore = True\nEnd Function\n');
		expect(rangeErrors(source)).toEqual([]);
	});

	it('does not assume Add leaves the new sheet active after activation events', () => {
		expect(rangeErrors(wrap('Dim w1 As Worksheet, w2 As Worksheet\nSet w1 = ActiveSheet\nSet w2 = Worksheets.Add\nDebug.Print w1.Range(Cells(1,1), Cells(2,2)).Count'))).toEqual([]);
	});

	it('retains immediate wrong-password and noncancellable deletion diagnostics', () => {
		expect(rangeErrors(wrap('Dim ws As Worksheet\nSet ws = Worksheets.Add\nws.Protect "pw"\nws.Unprotect "nope"'))).toHaveLength(1);
		expect(deletedErrors(wrap('Dim n As Name\nSet n = ThisWorkbook.Names.Add("xlideName", "=1")\nn.Delete\nDebug.Print n.Name'))).toHaveLength(1);
	});
});
