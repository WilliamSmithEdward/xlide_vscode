// Diagnostics tests: values kept through a self-increment, an empty
// Collection's Count, and a label forward GoTos reach (issue #614).
// Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const L = 'Dim a As Long, b As Long, s As String, arr(0 To 3) As Long, c As Collection\n    ';

function raised(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${L}${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

describe('values the straight line keeps (issue #614)', () => {
	it('follows a self-increment', () => {
		const cases: Array<[string, string]> = [
			['a = 1\n    b = b + 1\n    Main = 10 / (a - b)', '11'],
			['b = b + 1\n    b = b - 1\n    Main = 10 \\ b', '11'],
			['b = 5\n    b = b + 1\n    Main = arr(b)', '9'],
			['b = b + 4\n    Main = arr(b)', '9'],
			['a = a + 1\n    a = a + -1\n    Main = Mid(s, a)', '5'],
		];
		for (const [body, error] of cases) {
			expect(raised(body), body).toEqual([error]);
		}
		for (const body of ['b = 2\n    b = b + 1\n    Main = arr(b)', 'a = 1\n    b = b + 2\n    Main = 10 / (a - b)']) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('reads the Count of an empty Collection as 0', () => {
		for (const body of ['Set c = New Collection\n    Main = 10 \\ c.Count', 'Set c = New Collection\n    Main = 10 / c.Count', 'Dim d As New Collection\n    Main = 10 \\ d.Count']) {
			expect(raised(body), body).toEqual(['11']);
		}
		expect(raised('Set c = New Collection\n    c.Add 1\n    Main = 10 \\ c.Count')).toEqual([]);
	});

	it('enters a label forward GoTos reach with what every way in agrees on', () => {
		const cases: Array<[string, string]> = [
			['a = 1\n    If a > -1 Then GoTo Skip1\n    a = a + 1\nSkip1:\n    Main = c.Count', '91'],
			['Set c = Nothing\n    If b <> 1 Then GoTo Skip1\n    b = 1\nSkip1:\n    Main = c(b)', '91'],
			['b = 9\n    If a = 1 Then GoTo Skip1\n    a = 2\nSkip1:\n    Main = arr(b)', '9'],
			['b = 9\n    If Second(Now) > 70 Then GoTo Skip1\n    a = 2\nSkip1:\n    Main = arr(b)', '9'],
		];
		for (const [body, error] of cases) {
			expect(raised(body), body).toEqual([error]);
		}
		for (const body of [
			'b = 9\n    If a = 1 Then GoTo Skip1\n    b = 2\nSkip1:\n    Main = arr(b)',
			'b = 2\n    If a = 0 Then GoTo Skip1\n    b = 9\nSkip1:\n    Main = arr(b)',
			'Set c = New Collection\n    If a = 0 Then GoTo Skip1\n    Set c = Nothing\nSkip1:\n    Main = c.Count',
			'b = 9\n    If a = 1 Then GoTo Skip1\n    a = 2\nSkip1:\n    b = 1\n    Main = arr(b)',
		]) {
			expect(raised(body), body).toEqual([]);
		}
	});
});
