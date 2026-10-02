// Diagnostics tests: a value held in a typed local, passed to a built-in
// (issue #332). Each sample was run through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430), as `Dim a As <type>: a = <value>: Main = <call>`.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function raised(body: string): string[] {
	const source = `Option Explicit\nPrivate Function TakeVL(ByVal n As Long) As Long\n    TakeVL = n\nEnd Function\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

const held = (type: string, value: string, call: string): string => `Dim a As ${type}\n    a = ${value}\n    Main = ${call}`;

describe('a value a local holds, passed to a built-in (issue #332)', () => {
	it('is judged as the literal would be', () => {
		const cases: Array<[string, string, string, string]> = [
			['Double', '-1.5', 'Chr(a)', '5'],
			['Double', '-1.5', 'Sqr(a)', '5'],
			['Double', '-1.5', 'Left("abc", a)', '5'],
			['Currency', '922337203685477.5807@', 'Space(a)', '6'],
			['Currency', '922337203685477.5807@', 'DateAdd("d", a, #1/1/2000#)', '6'],
			['Double', '1E+300', 'Day(a)', '13'],
			['Long', '2147483647', 'Weekday(a)', '13'],
			['String', '"x"', 'Chr(a)', '13'],
			['String', '""', 'Left("abc", a)', '13'],
			['String', '"x"', 'DateAdd("d", a, #1/1/2000#)', '13'],
			['Variant', 'Null', 'Sqr(a)', '94'],
			['Variant', 'Null', 'Space(a)', '94'],
			['Variant', 'Null', 'Replace(a, "a", "b")', '94'],
			['Variant', 'Empty', 'Log(a)', '5'],
			['Variant', 'Empty', 'Asc(a)', '5'],
		];
		for (const [type, value, call, error] of cases) {
			expect(raised(held(type, value, call)), `${type} ${value} ${call}`).toEqual([error]);
		}
	});

	it('follows #239\'s Null into conversions, a loop bound, an index and a procedure', () => {
		const V = 'Dim v\n    v = Null\n    ';
		for (const body of ['Main = Mid$(v, 1)', 'Main = CLng(v)', 'Main = InStr(v, "abc", "a")', 'Main = UBound(Split(v))', 'Dim a(3) As Long\n    Main = a(v)', 'Dim i As Long\n    For i = 1 To v\n    Next', 'Main = TakeVL(v)']) {
			expect(raised(`${V}${body}`), body).toEqual(['94']);
		}
	});

	it('stays quiet where the value is taken', () => {
		const quiet = [
			held('Currency', '922337203685477.5807@', 'Hex(a)'),
			held('Currency', '922337203685477.5807@', 'Oct(a)'),
			'Main = Hex(3000000000.5)',
			held('Variant', 'Null', 'Oct(a)'),
			held('Variant', 'Null', 'IsNull(Left(a, 1))'),
			'Main = DateAdd("yyyy", 2147483647, #1/1/2000#)',
			'Main = DateAdd("d", -0.6, #1/1/100#)',
			held('Double', '2.5', 'Chr(a)'),
			held('String', '"65"', 'Chr(a)'),
		];
		for (const body of quiet) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('raises 6 for a DateAdd count outside a Long', () => {
		expect(raised('Main = DateAdd("d", 2147483648#, #1/1/2000#)')).toEqual(['6']);
		expect(raised('Main = DateAdd("m", -2147483649#, #1/1/2000#)')).toEqual(['6']);
	});
});
