// Diagnostics tests: a module variable keeps the value the procedure gave it
// across a call to a procedure of the module that leaves it alone (issue
// #618). Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeProjectModule } from './helpers';

const HEAD = 'Private mn As Long\nPrivate mc As Collection\nPublic pn As Long\n'
	+ 'Private Sub Nop()\nEnd Sub\n'
	+ 'Private Sub Log(ByVal t As String)\n    Debug.Print t\nEnd Sub\n'
	+ 'Private Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n'
	+ 'Private Sub SetMn()\n    mn = 1\nEnd Sub\n'
	+ 'Private Sub Indirect()\n    SetMn\nEnd Sub\n'
	+ 'Private Sub Outer()\n    Nop\n    Log "y"\nEnd Sub\n'
	+ 'Private Sub Fill(ByRef x As Long)\n    x = 2\nEnd Sub\n'
	+ 'Private Sub Calls()\n    Application.Calculate\nEnd Sub\n'
	+ 'Private Sub SetMc()\n    Set mc = New Collection\nEnd Sub\n';
const CODES = new Set(['array-subscript-out-of-bounds', 'division-by-zero', 'object-variable-not-set']);

function codes(body: string): string[] {
	const src = `Option Explicit\n${HEAD}Function Main() As Variant\n    Dim arr(0 To 3) As Long, r As Long\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeProjectModule(src, [{ moduleName: 'Module1', source: src }], 'Module1')
		.filter((d) => CODES.has(d.code)).map((d) => d.code);
}

describe('a module variable across a call that leaves it alone (issue #618)', () => {
	it('keeps the value across a call that never writes it', () => {
		expect(codes('mn = 4\n    Nop\n    Main = arr(mn)')).toEqual(['array-subscript-out-of-bounds']);
		for (const call of ['Nop', 'Log "x"', 'r = Twice(3)', 'Outer']) {
			expect(codes(`mn = 0\n    ${call}\n    Main = 10 \\ mn`), call).toEqual(['division-by-zero']);
		}
		expect(codes('pn = 0\n    Nop\n    Main = 10 \\ pn')).toEqual(['division-by-zero']);
	});

	it('follows a module object variable set to Nothing', () => {
		expect(codes('Set mc = Nothing\n    Main = mc.Count')).toEqual(['object-variable-not-set']);
		expect(codes('Set mc = Nothing\n    Nop\n    Main = mc.Count')).toEqual(['object-variable-not-set']);
	});

	it('stays quiet when a call may write it', () => {
		for (const body of ['mn = 0\n    SetMn\n    Main = 10 \\ mn', 'mn = 0\n    Indirect\n    Main = 10 \\ mn',
			'mn = 0\n    Fill mn\n    Main = 10 \\ mn', 'mn = 4\n    Calls\n    Main = arr(mn)',
			'Set mc = Nothing\n    SetMc\n    Main = mc.Count', 'Set mc = Nothing\n    If r = 0 Then SetMc\n    Main = mc.Count',
			'Set mc = Nothing\n    If r = 0 Then\n        SetMc\n    End If\n    Main = mc.Count',
			'Set mc = Nothing\n    Set mc = New Collection\n    Main = mc.Count', 'Set mc = Nothing\n    Main = mc Is Nothing']) {
			expect(codes(body), body).toEqual([]);
		}
	});
});
