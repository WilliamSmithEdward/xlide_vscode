// Diagnostics tests: #624's regressions, a Boolean local and a Date into a
// Byte, CStr of a number into a Byte, and Erase twice. Measured on
// 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const B = 'Dim o As Boolean, r As Byte, t As Date, i As Integer, n As Long\n    o = True\n    ';

function raised(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

describe('#624', () => {
	it('stores a Boolean True as 255 in a Byte, as the literal is', () => {
		for (const body of ['r = o', 'Main = CByte(o)', 'r = True', 'o = Not False\n    r = o', 'r = (n = 1)', 'n = o', 'i = o']) {
			expect(raised(B + body), body).toEqual([]);
		}
		expect(raised(`${B}r = o + 0`)).toEqual(['6']);
	});

	it('keeps the low byte of a Date in the Integer range', () => {
		for (const body of ['r = o + t', 'r = True + t', 'r = t - 1', 'r = CDate(-1)', 'r = t + 256', 'r = t + 1000', 'r = CDate(300.7)', 'r = o - t']) {
			expect(raised(B + body), body).toEqual([]);
		}
		for (const body of ['r = t + 65536', 'r = #1/1/2000#', 'i = t + 40000']) {
			expect(raised(B + body), body).toEqual(['6']);
		}
	});

	it('reads CStr of a number as the text it makes', () => {
		for (const body of ['r = CStr(40000)', 'r = CStr(300)', 'i = CStr(40000)', 'r = "300"']) {
			expect(raised(B + body), body).toEqual(['6']);
		}
		expect(raised(`${B}r = CStr(200)`)).toEqual([]);
	});

	it('keeps an array Erase emptied empty through a second Erase', () => {
		for (const body of ['Dim a As Variant\n    a = Split("x,y")\n    Erase a\n    Erase a\n    a(0) = 7', 'Dim a As Variant\n    a = Split("x,y")\n    Erase a\n    Erase a\n    Main = a(0)', 'Dim a() As Long\n    ReDim a(2)\n    Erase a\n    Erase a\n    Main = a(0)']) {
			expect(raised(body), body).toEqual(['9']);
		}
	});
});
