// Diagnostics tests: a value used where an object is expected (issue #325).
// Each sample was compiled or run through pyVBAharness on 2026-10-02 in
// Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string): string[] {
	const source = `Option Explicit\nPrivate Function TakeL(ByVal n As Long) As Long\n    TakeL = n\nEnd Function\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(source).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a value where an object is expected (issue #325)', () => {
	it('raises 424 on With or Is over a Variant holding a value', () => {
		for (const body of [
			'Dim v\n    v = 5\n    With v\n        .x = 1\n    End With',
			'Dim v\n    v = "abc"\n    With v\n        .x = 1\n    End With',
			'Dim v\n    With v\n        .x = 1\n    End With',
			'Dim v\n    v = 5\n    Main = (v Is Nothing)',
			'Dim v\n    v = "a"\n    Main = (v Is Nothing)',
			'Dim v\n    Main = (v Is Nothing)',
		]) {
			expect(found(body), body).toEqual(['variant-value-misuse']);
		}
	});

	it('is With object must be ... on a scalar element or Function result', () => {
		expect(found('Dim a(2) As Long\n    With a(1)\n    End With')).toEqual(['with-scalar-target']);
		expect(found('Dim a(2) As String\n    With a(1)\n    End With')).toEqual(['with-scalar-target']);
		expect(found('With TakeL(1)\n    End With')).toEqual(['with-scalar-target']);
	});

	it('is Invalid use of object on Nothing beside any operator but Is', () => {
		for (const expr of ['Nothing + 1', '"a" & Nothing', '(1 = Nothing)', '(Nothing And True)', '-Nothing', 'Not Nothing', '(Nothing Like "a")']) {
			expect(found(`Main = ${expr}`), expr).toEqual(['non-scalar-binary-operand']);
		}
	});

	it('is Type mismatch on Is over scalars in a one-line If or a loop condition', () => {
		for (const body of [
			'Dim n As Long\n    If n Is Nothing Then Main = 1',
			'Dim a As Long, b As Long\n    If a Is b Then Main = 1',
			'Dim s As String\n    If s Is Nothing Then Main = 1',
			'Dim n As Long\n    Do While n Is Nothing\n    Loop',
			'Dim n As Long\n    Do\n    Loop Until n Is Nothing',
		]) {
			expect(found(body), body).toEqual(['is-operator-non-object']);
		}
	});

	it('stays quiet where an object or a Variant goes', () => {
		for (const body of [
			'Dim v\n    v = 5\n    With v\n    End With',
			'Dim v\n    Set v = New Collection\n    With v\n        .Add 1\n    End With\n    Main = v.Count',
			'Dim v\n    Set v = New Collection\n    Main = (v Is Nothing)',
			'Dim a(2) As Variant\n    With a(1)\n    End With',
			'Dim a(2) As Collection\n    Set a(1) = New Collection\n    With a(1)\n    End With',
			'Main = (Nothing Is Nothing)',
			'Dim o As Object\n    If o Is Nothing Then Main = 1',
		]) {
			expect(found(body), body).toEqual([]);
		}
	});
});
