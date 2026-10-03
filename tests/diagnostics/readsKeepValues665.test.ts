// Diagnostics tests: statements that only read a local leave its value
// known (issue #665, #655's leftovers). Measured on 2026-10-03 in Excel 16.0
// (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HEAD = 'Private Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n'
	+ 'Private Sub SetR(ByRef x As Long)\n    x = 5\nEnd Sub\n'
	+ 'Private Function SetA(ByRef x As Long) As Long\n    x = 5\nEnd Function\n';
const D = 'Dim a As Long, b As Long, c As Collection, k As New Collection\n    ';
const READS = ['k.Add a', 'With k\n        .Add a\n    End With', 'ActiveSheet.Cells(a + 1, 1).Value = 1', 'b = Twice(a)', 'L1:', 'b = a'];

function errors(body: string): string[] {
	const src = `Option Explicit\n${HEAD}Function Main() As Variant\n    ${D}${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('statements that only read a local (issue #665)', () => {
	it('keep the value for a later division', () => {
		for (const read of READS) {
			expect(errors(`a = 0\n    ${read}\n    Main = 10 \\ a`), read).toEqual(['division-by-zero']);
		}
	});

	it('keep a guard on it decided', () => {
		for (const read of READS) {
			expect(errors(`a = 1\n    ${read}\n    If a = 2 Then c.Add 1`), read).toEqual([]);
		}
	});

	it('takes a Function that passes its ByRef argument on as writing it', () => {
		const src = 'Option Explicit\n'
			+ 'Private Function TryGetObject(ByVal k As String, ByRef found As Collection) As Boolean\n    Set found = New Collection\n    TryGetObject = True\nEnd Function\n'
			+ 'Private Function TryGet(ByVal k As String, ByRef found As Collection) As Boolean\n    TryGet = TryGetObject(k, found)\nEnd Function\n'
			+ 'Function Main() As Variant\n    Dim c As Collection\n    If TryGet("a", c) Then\n        Main = c.Count\n    End If\nEnd Function\n';
		expect(analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code)).toEqual([]);
	});

	it('still forget a value a callee may write', () => {
		expect(errors('a = 0\n    SetR a\n    Main = 10 \\ a')).toEqual([]);
		expect(errors('a = 0\n    b = Twice(a) + SetA(a)\n    Main = 10 \\ a')).toEqual([]);
	});
});
