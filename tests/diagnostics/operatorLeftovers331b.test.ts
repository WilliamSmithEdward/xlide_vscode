// Diagnostics tests: #331's last operator shapes. Measured on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function raised(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

describe('#331\'s leftovers', () => {
	it('reports each as Excel raises it', () => {
		const cases: Array<[string, string]> = [
			['Dim s As String\n    s = "12"\n    Main = s ^ 32767', '6'],
			['Main = "12" ^ 32767', '6'],
			['Dim n As Long\n    n = Empty\n    Main = 10 / n', '11'],
			['Dim n As Long, b As Boolean\n    n = 2147483647\n    b = True\n    Main = n - b', '6'],
			['Dim n As Integer, b As Boolean\n    n = 32767\n    b = True\n    Main = n - b', '6'],
			['Dim s As String, b As Boolean\n    s = "True"\n    b = True\n    Main = s + b', '13'],
			['Dim s As String\n    s = "True"\n    Main = s + True', '13'],
		];
		for (const [body, error] of cases) {
			expect(raised(body), body).toEqual([error]);
		}
	});

	it('leaves alone the neighbours that run', () => {
		for (const body of [
			'Dim s As String\n    s = "12"\n    Main = s ^ 2',
			'Dim n As Long, b As Boolean\n    n = 2147483647\n    b = False\n    Main = n - b',
			'Dim n As Long, b As Boolean\n    n = 2147483647\n    b = True\n    Main = n + b',
			'Dim s As String, b As Boolean\n    s = "1"\n    b = True\n    Main = s + b',
			'Dim s As String, b As Boolean\n    s = "True"\n    b = True\n    Main = s & b',
		]) {
			expect(raised(body), body).toEqual([]);
		}
	});
});
