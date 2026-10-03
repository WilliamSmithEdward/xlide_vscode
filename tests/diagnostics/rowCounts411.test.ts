// Diagnostics tests: a sheet's row count read inside With, and passed to a
// ByVal Integer parameter (issue #411). Each case was run through
// pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const TAIL = 'Private Function TakeI(ByVal i As Integer) As Integer\n    TakeI = i\nEnd Function\n'
	+ 'Private Function TakeL(ByVal i As Long) As Long\n    TakeL = i\nEnd Function\n';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n${TAIL}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe("a sheet's row count (issue #411)", () => {
	it('overflows an Integer inside With and through a ByVal Integer parameter', () => {
		expect(errors('Dim i As Integer\n    With ActiveSheet\n        i = .Rows.Count\n    End With')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim i As Integer\n    With ActiveSheet: i = .Rows.Count: End With')).toEqual(['arithmetic-overflow']);
		expect(errors('Dim i As Integer, ws As Worksheet\n    Set ws = ActiveSheet\n    With ws\n        i = .Rows.Count\n    End With')).toEqual(['arithmetic-overflow']);
		expect(errors('Main = TakeI(Rows.Count)')).toEqual(['arithmetic-overflow']);
		expect(errors('Main = TakeI(ActiveSheet.Rows.Count)')).toEqual(['arithmetic-overflow']);
	});

	it('stays quiet where the count fits', () => {
		expect(errors('Dim i As Long\n    With ActiveSheet\n        i = .Rows.Count\n    End With')).toEqual([]);
		expect(errors('Dim i As Integer\n    With ActiveSheet\n        i = .Columns.Count\n    End With')).toEqual([]);
		expect(errors('Main = TakeL(Rows.Count)')).toEqual([]);
		expect(errors('Main = TakeI(Columns.Count)')).toEqual([]);
		expect(errors('Dim i As Integer\n    i = TakeI(Rows.Count \\ 64)')).toEqual([]);
	});
});
