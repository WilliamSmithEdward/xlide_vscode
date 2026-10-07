// Diagnostics tests: #584's regressions, an error handler after a Delete,
// ReDim after Erase, and Adds inside With c. Measured on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function raised(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

const COLLECTION = 'Dim c As Collection\n    Set c = New Collection\n    ';

describe('#584: the shapes that run', () => {
	it('reads an error handler as reached from before the Delete', () => {
		const body = 'Dim w2 As Worksheet\n    Set w2 = Worksheets.Add\n    On Error GoTo H\n    w2.Protect\n    w2.Unprotect\n    Application.DisplayAlerts = False\n    w2.Delete\n    Application.DisplayAlerts = True\n    Exit Function\nH:\n    On Error Resume Next\n    w2.Unprotect "pw"\n    w2.Delete';
		expect(raised(body)).toEqual([]);
	});

	it('takes ReDim as allocating again what Erase emptied', () => {
		for (const body of ['Dim a As Variant\n    a = Split("x,y")\n    Erase a\n    ReDim a(1)\n    Main = a(1)', 'Dim a As Variant\n    a = Array(10, 20, 30)\n    Erase a\n    ReDim Preserve a(5)\n    Main = UBound(a)']) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('counts the Adds inside With c', () => {
		expect(raised(`${COLLECTION}With c\n        .Add 10, "k"\n        .Add 20\n    End With\n    Main = c("k") + c(2)`)).toEqual([]);
		expect(raised(`${COLLECTION}With c\n        .Add 10\n        If Main = 0 Then .Add 20\n    End With\n    Main = c(2)`)).toEqual([]);
		expect(raised(`${COLLECTION}With c\n        .Add 10\n    End With\n    c.Add 20\n    Main = c(2)`)).toEqual([]);
	});
});

describe('#584: what still raises', () => {
	it('reports Erase without ReDim, and what With c leaves', () => {
		const cases: Array<[string, string]> = [
			['Dim a As Variant\n    a = Array(10, 20, 30)\n    Erase a\n    Main = a(1)', '9'],
			[`${COLLECTION}With c\n        .Add 10\n        .Add 20\n    End With\n    Main = c(3)`, '9'],
			[`${COLLECTION}With c\n        .Add 10, "k"\n    End With\n    Main = c("zz")`, '5'],
			[`${COLLECTION}c.Add 10\n    c.Add 20\n    With c\n        .Remove 1\n    End With\n    Main = c(2)`, '9'],
		];
		for (const [body, error] of cases) {
			expect(raised(body), body).toEqual([error]);
		}
	});
});
