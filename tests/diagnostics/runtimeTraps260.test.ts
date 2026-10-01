// Diagnostics tests: Split's limit, Filter's result, an empty Split through a
// String local, string arithmetic on an element, and omitted Optional and
// ParamArray arguments read by the procedure called (issue #260). Every case
// was measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function source(lines: string[], procedures = '', optionBase = ''): string {
	return `Option Explicit\n${optionBase}Function Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n${procedures}`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('Split and Filter shapes', () => {
	it.each([
		[['Dim v As Variant', 'v = Split("a,b,c", ",", 1)', 'Main = v(1)'], '1', 'above the upper bound 0'],
		[['Dim v As Variant', 'v = Split("a,b,c", ",", 2)', 'Main = v(2)'], '2', 'above the upper bound 1'],
		[['Main = Split("a,b,c", ",", 2)(2)'], '2', 'the array Split returns here is above the upper bound 1'],
		[['Dim v As Variant', 'v = Split("a,b,c", ",", 0)', 'Main = v(0)'], '0', 'the array is empty'],
		[['Dim v As Variant', 'v = Split("aXbxc", "x", -1, vbTextCompare)', 'Main = v(3)'], '3', 'above the upper bound 2'],
		[['Dim v As Variant', 'v = Filter(Array("a"), "z")', 'Main = v(0)'], '0', "array 'v' (Filter(...)) has no element to reach"],
		[['Main = Filter(Array("a"), "z")(0)'], '0', 'the array Filter returns here has no element to reach'],
		[['Dim v As Variant', 'v = Filter(Array("ab", "b", "c"), "b", False)', 'Main = v(1)'], '1', 'above the upper bound 0'],
		[['Dim v As Variant', 'v = Filter(Array("B", "b"), "b")', 'Main = v(1)'], '1', 'above the upper bound 0'],
		[['Dim v As Variant', 'v = Filter(Split("a b c"), "z")', 'Main = v(0)'], '0', 'the array is empty'],
	])('knows the bounds of %j', (lines, span, message) => {
		const src = source(lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'array-subscript-out-of-bounds'), 'array-subscript-out-of-bounds', { span, message });
	});

	it.each([
		[['Dim s As String, v As Variant', 'v = Split(s)', 'Main = v(0)'], 'the array is empty'],
		[['Dim s As String, v As Variant', 's = "a,b"', 'v = Split(s, ",")', 'Main = v(2)'], 'above the upper bound 1'],
		[['Dim d As String, v As Variant', 'd = ","', 'v = Split("a,b", d)', 'Main = v(2)'], 'above the upper bound 1'],
		[['Dim s As Variant, v As Variant', 's = "a,b"', 'v = Split(s, ",")', 'Main = v(2)'], 'above the upper bound 1'],
		[['Const S As String = "a,b"', 'Dim v As Variant', 'v = Split(S, ",")', 'Main = v(2)'], 'above the upper bound 1'],
		// After a label only the procedure-wide shape is known.
		[['Dim s As String, v As Variant', 's = "a,b"', 'v = Split(s, ",")', 'Again:', 'Main = v(2)'], 'above the upper bound 1'],
	])('reads a String local or Const where the array is built: %j', (lines, message) => {
		const src = source(lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'array-subscript-out-of-bounds'), 'array-subscript-out-of-bounds', { message });
	});

	it('reads the local as the Split statement saw it', () => {
		const src = source(['Dim s As String, v As Variant', 's = "a,b"', 'v = Split(s, ",")', 's = "a"', 'Main = v(1)']);
		expect(errors(src)).toEqual([]);
		// v is assigned twice, so only the straight-line shape knows it.
		const twice = source(['Dim s As String, v As Variant', 's = "a,b"', 'v = Split(s, ",")', 'Main = v(2)', 'v = Array(1, 2, 3)']);
		expectDiagnostic(twice, byCode(analyzeModule(twice), 'array-subscript-out-of-bounds'), 'array-subscript-out-of-bounds', { message: 'above the upper bound 1' });
	});

	it('stays quiet where the index is in range or the shape is unknown', () => {
		for (const lines of [
			['Dim v As Variant', 'v = Split("a,b,c", ",", 2)', 'Main = v(1)'],
			['Dim v As Variant', 'v = Split("a,b,c", ",", -1)', 'Main = v(2)'],
			['Dim v As Variant', 'v = Split("a,b,c", ",", 5)', 'Main = v(2)'],
			['Dim v As Variant', 'v = Split(",,", ",", 2)', 'Main = v(1)'],
			['Dim v As Variant', 'v = Filter(Array("a", "b"), "a")', 'Main = v(0)'],
			['Dim v As Variant', 'v = Filter(Array(1, 12, 3), "1")', 'Main = v(1)'],
			['Dim v As Variant', 'v = Filter(Array("a", "b"), "")', 'Main = v(1)'],
			['Dim v As Variant', 'v = Filter(Array("B", "b"), "b", True, vbTextCompare)', 'Main = v(1)'],
			['Dim s As String, v As Variant', 's = "a" & ",b"', 'v = Split(s, ",")', 'Main = v(1)'],
		]) {
			expect(errors(source(lines)), lines.join(': ')).toEqual([]);
		}
		expect(errors(source(['Dim v As Variant', 'v = Filter(Array("a", "b"), "a")', 'Main = v(0)'], '', 'Option Base 1\n'))).toEqual([]);
	});
});

describe('string arithmetic on an element', () => {
	it.each([
		[['Dim v As Variant', 'v = Array("1", "b")', 'Main = v(1) + 1'], 'v(1)', '+'],
		[['Dim v As Variant', 'v = Split("1 b")', 'Main = v(1) + 1'], 'v(1)', '+'],
		[['Main = Split("1 b")(1) + 1'], 'Split("1 b")(1)', '+'],
		[['Dim v As Variant', 'v = Array("1", "b")', 'Main = 2 * v(1)'], 'v(1)', '*'],
		[['Dim v As Variant', 'v = Array("1", "b")', 'Main = -v(1)'], 'v(1)', '-'],
		[['Dim v As Variant', 'v = Array("1", "b")', 'Main = v(1) Mod 2'], 'v(1)', 'Mod'],
		[['Dim v As Variant', 'v = Filter(Array("1", "b"), "b")', 'Main = v(0) + 1'], 'v(0)', '+'],
	])('reports %j', (lines, span, operator) => {
		const src = source(lines);
		expectDiagnostic(src, byCode(analyzeModule(src), 'string-arithmetic-coercion'), 'string-arithmetic-coercion', { span, message: `Operator '${operator}' coerces '${span}', which holds "b"` });
	});

	it('stays quiet on a number, a concatenation, two strings, and an element written since', () => {
		for (const lines of [
			['Dim v As Variant', 'v = Array("1", "b")', 'Main = v(0) + 1'],
			['Dim v As Variant', 'v = Array("1", "b")', 'Main = v(1) & 1'],
			['Dim v As Variant', 'v = Array("1", "b")', 'Main = v(1) + "c"'],
			['Dim v As Variant', 'v = Array(" 2 ", "b")', 'Main = v(0) + 1'],
			['Dim v As Variant', 'v = Array("&H10", "b")', 'Main = v(0) + 1'],
			['Dim v As Variant', 'v = Array("1", "b")', 'v(1) = 2', 'Main = v(1) + 1'],
			['Dim v As Variant', 'v = Array("1", "b")', 'Fill v(1)', 'Main = v(1) + 1'],
		]) {
			expect(errors(source(lines, 'Private Sub Fill(ByRef x As Variant)\n    x = 1\nEnd Sub\n')), lines.join(': ')).toEqual([]);
		}
		expect(errors(source(['Dim v As Variant', 'v = Array("1", "b")', 'Main = v(1) + 1'], '', 'Option Base 1\n'))).toEqual([]);
	});
});

describe('an omitted argument read by the procedure called', () => {
	const opt = (body: string, signature = 'Optional x As Variant'): string => `Private Function Opt(${signature}) As Variant\n${body}\nEnd Function\nPrivate Function Inner(ByVal y As Long) As Long\n    Inner = y\nEnd Function\n`;

	it.each([
		['    Opt = x + 1', "uses it in 'x + 1'"],
		['    Opt = Len(x)', 'passes it to Len'],
		['    Opt = Left$(x, 1)', 'passes it to Left$'],
		['    Opt = "a" & x', `uses it in '"a" & x'`],
		['    If x Then Opt = 1', "tests it in 'If x Then Opt = 1'"],
		['    Dim n As Long\n    n = x\n    Opt = n', "assigns it to 'n', a Long"],
		['    Opt = Inner(x)', "passes it ByVal to the Long 'y' of 'Inner'"],
		['    Select Case x\n    Case 1\n    End Select', "tests it in 'Select Case x'"],
		['    Dim i As Long\n    For i = 1 To 2\n    Next i\n    Opt = x + 1', "uses it in 'x + 1'"],
		['Start:\n    Opt = Not x', "uses it in 'Not x'"],
	])('reports a Missing read: %j', (body, does) => {
		const src = source(['Main = Opt()'], opt(body));
		expectDiagnostic(src, byCode(analyzeModule(src), 'variant-value-misuse'), 'variant-value-misuse', {
			span: 'Opt',
			message: `This call omits 'x', so it is Missing in 'Opt', which ${does}`,
		});
	});

	it('names the skipped slot and the line', () => {
		const src = source(['Main = Opt(1, , 3)'], opt('    Opt = a + x + y', 'a As Long, Optional x As Variant, Optional y As Variant'));
		const diag = byCode(analyzeModule(src), 'variant-value-misuse');
		expect(diag).toHaveLength(1);
		// The empty slot, between its commas.
		expect(diag[0].span.start).toBeGreaterThan(src.indexOf('Opt(1,'));
		expect(diag[0].span.end).toBeLessThanOrEqual(src.indexOf(', 3)') + 1);
		expect(diag[0].message).toContain(`(line ${src.split('\n').findIndex((line) => line.includes('a + x + y')) + 1})`);
	});

	it('reports the statement and named forms', () => {
		for (const [call, signature, body] of [
			['Opt', 'Optional x As Variant', '    Debug.Print x * 2'],
			['Call Opt', 'Optional x As Variant', '    Opt = x + 1'],
			['Main = Opt(y:=2)', 'Optional x As Variant, Optional y As Variant', '    Opt = x + y'],
			['Main = Opt()', 'Optional x', '    Opt = x + 1'],
		]) {
			expect(errors(source([call, 'Main = 1'], opt(body, signature))), call).toContain('variant-value-misuse');
		}
	});

	it('stays quiet where the Missing value is read cleanly or not reached', () => {
		for (const body of [
			'    Opt = IsMissing(x)',
			'    Opt = IsEmpty(x)',
			'    Opt = CStr(x)',
			'    Opt = CLng(x)',
			'    Opt = VarType(x)',
			'    Opt = x',
			'    Dim v As Variant\n    v = x\n    Opt = 1',
			'    If IsMissing(x) Then x = 5\n    Opt = x + 1',
			'    If IsMissing(x) Then Opt = 0 Else Opt = x + 1',
			'    On Error Resume Next\n    Opt = x + 1',
			'    If True Then\n        Exit Function\n    End If\n    Opt = x + 1',
			'    If False Then Opt = x + 1\n    Opt = 0',
			'    Debug.Print x\n    Opt = 1',
			'    Opt = Opt2(x)',
		]) {
			const src = source(['Main = Opt()'], `${opt(body)}Private Function Opt2(y As Variant) As Variant\n    Opt2 = 1\nEnd Function\n`);
			expect(errors(src), body).toEqual([]);
		}
		expect(errors(source(['Main = Opt(1)'], opt('    Opt = x + 1')))).toEqual([]);
	});

	it.each([
		['Optional x As Long', '    Opt = 10 / x', 'division-by-zero', "so it is 0 in 'Opt', which divides by it in '10 / x'"],
		['Optional x As Long = 0', '    Opt = 10 \\ x', 'division-by-zero', "so it is 0 in 'Opt'"],
		['Optional x As Boolean', '    Opt = 10 Mod x', 'division-by-zero', "so it is 0 in 'Opt'"],
		['Optional x&', '    Opt = 10 / x * 2', 'division-by-zero', "so it is 0 in 'Opt'"],
		['Optional x As Variant = 0', '    Opt = 10 / x', 'division-by-zero', "so it is 0 in 'Opt'"],
		['Optional x As String', '    Opt = CLng(x)', 'runtime-conversion-value', 'so it is "" in \'Opt\', which converts it with CLng'],
		['Optional x As String', '    Opt = x * 2', 'string-arithmetic-coercion', 'so it is "" in \'Opt\', which uses it as a number'],
		['Optional x As String = "abc"', '    Opt = x + 1', 'string-arithmetic-coercion', 'so it is "abc" in \'Opt\''],
	])('reports a typed default: %s', (signature, body, code, message) => {
		const src = source(['Main = Opt()'], opt(body, signature));
		expectDiagnostic(src, byCode(analyzeModule(src), code), code, { span: 'Opt', message });
	});

	it('stays quiet on a default that does not raise', () => {
		for (const [signature, body] of [
			['Optional x As Long = 2', '    Opt = 10 / x'],
			['Optional x As Long', '    Opt = 10 / x ^ 0'],
			['Optional x As String', '    Opt = Val(x)'],
			['Optional x As String = "5"', '    Opt = x * 2'],
		]) {
			expect(errors(source(['Main = Opt()'], opt(body, signature))), signature + body).toEqual([]);
		}
	});

	it.each([
		['Main = Total()', 'Total = args(0)', 'reads args(0) (line', 'passes no values to that ParamArray, so its upper bound is -1'],
		['Main = Total(1, 2)', 'Total = args(3)', 'reads args(3)', 'passes 2 values'],
		['Main = Total(1)', 'Total = args(-1)', 'reads args(-1)', 'passes 1 value to'],
		['Total 1\n    Main = 1', 'Debug.Print args(1)', 'reads args(1)', 'its upper bound is 0'],
	])('reports a ParamArray read past what is passed: %s', (call, body, reads, passes) => {
		const src = source([call], `Private Function Total(ParamArray args() As Variant) As Variant\n    ${body}\nEnd Function\n`, 'Option Base 1\n');
		const diag = byCode(analyzeModule(src), 'array-subscript-out-of-bounds');
		expect(diag).toHaveLength(1);
		expect(diag[0].message).toContain(`'Total' ${reads}`);
		expect(diag[0].message).toContain(passes);
	});

	it('reports a skipped ParamArray element and stays quiet on what is passed', () => {
		const procs = 'Private Function Total(ParamArray args() As Variant) As Variant\n    Total = args(1) + 1\nEnd Function\n';
		expectDiagnostic(source(['Main = Total(1, , 3)'], procs), byCode(analyzeModule(source(['Main = Total(1, , 3)'], procs)), 'variant-value-misuse'), 'variant-value-misuse', {
			span: 'Total',
			message: "This call skips an element of 'args', so it is Missing in 'Total', which uses it in 'args(1) + 1'",
		});
		expect(errors(source(['Main = Total(1, 2)'], procs))).toEqual([]);
		const guarded = 'Private Function Total(ParamArray args() As Variant) As Variant\n    If UBound(args) < 0 Then Exit Function\n    Total = args(0)\nEnd Function\n';
		expect(errors(source(['Main = Total()'], guarded))).toEqual([]);
	});
});
