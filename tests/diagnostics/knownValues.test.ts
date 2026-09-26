// Diagnostics tests: values the code makes plain (issues #119 and #106).
// Each raising statement was measured in Excel 16.0 (build 20326,
// 2026-09-26); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const COERCE = 'string-arithmetic-coercion';
const DIVIDE = 'division-by-zero';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('string-arithmetic-coercion - the operator raises whatever the target (issue #119)', () => {
	it('flags a nonnumeric string added to a number into a Variant and as a function result', () => {
		const src = wrap('Dim v As Variant', 'v = "abc" + 1', 'Main = "abc" + 1');
		expectDiagnostics(src, analyzeModule(src), COERCE, [
			{ span: '"abc"', message: ["Operator '+'", "error '13'"] },
			{ span: '"abc"' },
		]);
	});

	it('flags Not, unary minus and a comparison with a number', () => {
		const src = wrap('Main = Not "abc"', 'Main = -"abc"', 'If "abc" = 1 Then Main = 1');
		expectDiagnostics(src, analyzeModule(src), COERCE, [
			{ span: '"abc"', message: "Operator 'Not'" },
			{ span: '"abc"', message: "Operator '-'" },
			{ span: '"abc"', message: "Operator '='" },
		]);
	});

	it('flags a String local whose one assignment is a nonnumeric literal', () => {
		const src = wrap('Dim s As String', 's = "abc"', 'Main = s * 2');
		expectDiagnostic(src, analyzeModule(src), COERCE, { span: 's', message: ["Operator '*'", '"abc"'] });
	});

	it('reports a numeric-target assignment once, through the assignment rule', () => {
		const src = wrap('Dim l As Long', 'l = "abc" + 1', 'Main = l');
		expectDiagnostic(src, analyzeModule(src), COERCE, { span: '"abc"', message: "Assignment to 'l'" });
	});

	it('stays quiet where the strings concatenate, compare as text, or are numeric', () => {
		const src = wrap(
			'Dim s As String, t As String',
			's = "abc"',
			't = "5"',
			'Main = "abc" + "def"',
			'Main = "abc" & 1',
			'If "abc" = "abc" Then Main = 1',
			'Main = "5" + 1',
			'Main = t * 2',
			'Main = s & 2',
			's = "1"',
		);
		expect(byCode(analyzeModule(src), COERCE)).toHaveLength(0);
	});
});

describe('division-by-zero - divisors the code makes plain (issues #119 and #106)', () => {
	it('flags a divisor local the procedure never assigns, and one only ever assigned zero', () => {
		const src = wrap('Dim d As Long, e As Long', 'e = 0', 'Main = 10 / d', 'Main = 10 / e');
		expectDiagnostics(src, analyzeModule(src), DIVIDE, [{ span: 'd' }, { span: 'e' }]);
	});

	it('stays quiet for a local that is assigned anywhere else, passed whole, or counted', () => {
		const src = wrap(
			'Dim d As Long, e As Long, f As Long',
			'e = 0',
			'If Main Then e = 2',
			'Fill f',
			'For d = 1 To 3',
			'Next d',
			'Main = 10 / d + 10 / e + 10 / f',
		);
		expect(byCode(analyzeModule(`${src}Sub Fill(ByRef n As Long)\n    n = 1\nEnd Sub\n`), DIVIDE)).toHaveLength(0);
	});

	it('flags integer division and Mod by a literal that rounds to zero', () => {
		const src = wrap('Main = 5 \\ 0.4', 'Main = 5 Mod 0.5', 'Main = 5 / 0.4', 'Main = 5 \\ 0.6');
		expectDiagnostics(src, analyzeModule(src), DIVIDE, [
			{ span: '0.4', message: "'\\'" },
			{ span: '0.5', message: "'Mod'" },
		]);
	});

	it('says Overflow for zero divided by zero with /', () => {
		const src = wrap('Main = 0 / 0', 'Main = 0 \\ 0');
		expectDiagnostics(src, analyzeModule(src), DIVIDE, [
			{ span: '0', message: "error '6'" },
			{ span: '0', message: "error '11'" },
		]);
	});

	it('stays quiet under a guard that tests the constant divisor', () => {
		const src =
			'Option Explicit\nPrivate Const SCALE_BY = 0\nFunction Main() As Variant\n' +
			'    Main = 10\n' +
			'    If SCALE_BY <> 0 Then Main = 10 / SCALE_BY\n' +
			'    If SCALE_BY <> 0 Then\n        Main = 10 / SCALE_BY\n    End If\n' +
			'    If SCALE_BY = 0 Then\n        Main = 0\n    Else\n        Main = 10 / SCALE_BY\n    End If\n' +
			'    If Not SCALE_BY = 0 And Main > 0 Then Main = 10 / SCALE_BY\n' +
			'End Function\n';
		expect(byCode(analyzeModule(src), DIVIDE)).toHaveLength(0);
		const unguarded =
			'Option Explicit\nPrivate Const SCALE_BY = 0\nFunction Main() As Variant\n' +
			'    If SCALE_BY = 0 Then Main = 10 / SCALE_BY\n' +
			'    If Main > 0 Or SCALE_BY <> 0 Then Main = 10 / SCALE_BY\n' +
			'End Function\n';
		expect(byCode(analyzeModule(unguarded), DIVIDE)).toHaveLength(2);
	});
});
