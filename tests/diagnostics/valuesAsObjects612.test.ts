// Diagnostics tests: #325's follow-ups (issue #612), a Variant holding no
// object used as one. Measured on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const HEAD = 'Private Sub Init(x As Variant)\n    Set x = New Collection\nEnd Sub\nPrivate Function Make() As Object\n    Set Make = New Collection\nEnd Function\n';

function raised(body: string): string[] {
	const source = `Option Explicit\n${HEAD}Function Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source)
		.filter((diag) => diag.severity === 'error')
		.map((diag) => /Run-time error '(-?\d+)'/.exec(diag.message)?.[1] ?? diag.code);
}

describe('a Variant holding no object, used as one (issue #612)', () => {
	it('leaves alone TypeOf and an IsObject guard', () => {
		for (const body of [
			'Dim v\n    v = 5\n    Main = TypeOf v Is Collection',
			'Dim v\n    v = "a"\n    Main = TypeOf v Is Range',
			'Dim v\n    v = 5\n    If IsObject(v) Then\n        Main = (v Is Nothing)\n    Else\n        Main = 0\n    End If',
		]) {
			expect(raised(body), body).toEqual([]);
		}
	});

	it('reports Null, an error value, an array, a Static never set and a For Each\'s variable', () => {
		for (const body of [
			'Dim v\n    v = Null\n    Main = (v Is Nothing)',
			'Dim v\n    v = CVErr(5)\n    Main = (v Is Nothing)',
			'Static v\n    Main = (v Is Nothing)',
			'Dim v\n    v = Array(1)\n    Main = (v Is Nothing)',
			'Dim v\n    v = ActiveSheet.Range("A1")\n    With v\n        .Value = 1\n    End With',
			'Dim v, c As New Collection\n    c.Add 5\n    For Each v In c: Next\n    Main = (v Is Nothing)',
			'Dim v, c As New Collection\n    c.Add 5\n    For Each v In c\n        Main = (v Is Nothing)\n    Next',
		]) {
			expect(raised(body), body).toEqual(['424']);
		}
	});

	it('reports With on what a VBA function returns', () => {
		expect(raised('With CStr(1)\n        Main = 2\n    End With')).toEqual(['with-scalar-target']);
	});

	it('stays quiet where v holds an object, or With reaches no member', () => {
		for (const body of [
			'Dim v\n    v = 5\n    Init v\n    Main = (v Is Nothing)',
			'Dim v\n    v = 5\n    Set v = Make()\n    Main = (v Is Nothing)',
			'Dim v\n    v = 5\n    If True Then Set v = Nothing\n    Main = (v Is Nothing)',
			'Dim v\n    v = Null\n    Set v = Nothing\n    Main = (v Is Nothing)',
			'Dim v, c As New Collection\n    c.Add New Collection\n    For Each v In c\n        Main = (v Is Nothing)\n    Next',
			'Dim v, c As New Collection\n    c.Add 5\n    For Each v In c\n        Set v = New Collection\n        Main = (v Is Nothing)\n    Next',
			'Dim v, c As New Collection\n    c.Add 5\n    For Each v In c: Next\n    Set v = New Collection\n    Main = (v Is Nothing)',
			'Dim v, c As New Collection\n    c.Add 5\n    For Each v In c\n        Main = v + 1\n    Next',
			'Dim v\n    v = Array(1)\n    With v\n        Main = 2\n    End With',
			'Dim v\n    v = 5\n    With v\n        Debug.Print 1\n    End With',
		]) {
			expect(raised(body), body).toEqual([]);
		}
	});
});
