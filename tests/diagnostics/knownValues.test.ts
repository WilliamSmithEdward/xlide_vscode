// Diagnostics tests: values the code makes plain (issues #119, #106 and #180).
// Each raising statement was measured in Excel 16.0 (build 20326,
// 2026-09-26); each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics, spanText } from '../helpers/diagnostics';

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

describe('string-arithmetic-coercion - conditions, logical operators, Case and For (issue #191)', () => {
	// Measured in Excel 16.0 (build 20326, 2026-09-29): each raises 13.
	const run = (...lines: string[]): string[] => {
		const src = `Option Explicit\nFunction Main() As Variant\n    Dim answer As String, flag As String, last As String, i As Long, ready As Boolean\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
		return byCode(analyzeModule(src), COERCE).map((d) => d.message);
	};

	it('flags a string condition, a string under a logical operator, a Case value and a For bound', () => {
		const cases: Array<[string[], string]> = [
			[['answer = "yes"', 'If answer Then Main = "on"'], "'If' converts 'answer', which holds \"yes\" to Boolean"],
			[['If "abc" Then Main = 1'], "'If' converts string literal \"abc\""],
			[['If "" Then Main = 1'], "'If'"],
			[['If " True " Then Main = 1'], "'If'"],
			[['If "abc" Then', '    Main = 1', 'End If'], "'If'"],
			[['If False Then', 'ElseIf "abc" Then', '    Main = 2', 'End If'], "'ElseIf'"],
			[['Do While "abc"', '    Exit Do', 'Loop'], "'While'"],
			[['Do', '    Main = 1', 'Loop Until "abc"'], "'Until'"],
			[['While "abc"', '    Main = 1', 'Wend'], "'While'"],
			[['Main = IIf("abc", 1, 0)'], "'IIf'"],
			[['flag = "Y"', 'Main = ready And flag'], "Operator 'And' coerces 'flag'"],
			[['Main = "abc" And 1'], "Operator 'And'"],
			[['Main = "True" Or 0'], "Operator 'Or' coerces string literal \"True\""],
			[['Main = "" Xor 1'], "Operator 'Xor'"],
			[['Main = "abc" Eqv True'], "Operator 'Eqv'"],
			[['Select Case 1', 'Case "abc"', '    Main = 1', 'End Select'], 'Case compares string literal "abc" with a number'],
			[['For i = 1 To "abc"', 'Next i'], 'for its end'],
			[['last = "ten"', 'For i = 1 To last', 'Next i'], "'last', which holds \"ten\""],
			[['For i = 1 To 3 Step "abc"', 'Next i'], 'for its step'],
			[['Main = #1/1/2000# + "abc"'], "Operator '+'"],
		];
		for (const [lines, message] of cases) {
			const messages = run(...lines);
			expect(messages, lines.join(' / ')).toHaveLength(1);
			expect(messages[0], lines.join(' / ')).toContain(message);
		}
	});

	it('stays quiet where the string converts, or belongs to a comparison', () => {
		const quiet = [
			['If "True" Then Main = 1'],
			['If "1" Then Main = 1'],
			['Main = "5" And 1'],
			['Select Case 1', 'Case "1"', '    Main = 1', 'End Select'],
			['For i = 1 To "3"', 'Next i'],
			['Main = #1/1/2000# + "1"'],
			['answer = "yes"', 'If answer = "yes" Then Main = 1'],
			['answer = "yes"', 'If answer = "yes" Or answer = "no" Then Main = 1'],
			['answer = "yes"', 'Select Case answer', 'Case "abc"', '    Main = 1', 'End Select'],
			['Main = "abc" = "abc" And 1'],
			['Main = 1 + 2 And "5"'],
		];
		for (const lines of quiet) {
			expect(run(...lines), lines.join(' / ')).toEqual([]);
		}
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

describe('the value a statement sees, though the variable changes later (issue #180)', () => {
	const found = (...lines: string[]): string[] => {
		const src = wrap(...lines);
		return analyzeModule(src)
			.filter((d) => [DIVIDE, COERCE, 'variant-value-misuse', 'array-subscript-out-of-bounds', 'runtime-argument-value'].includes(d.code))
			.map((d) => `${d.code} ${spanText(src, d)}`);
	};

	it('flags each statement of the issue, reassigned after the line that fails', () => {
		expect(found('Dim d As Long', 'd = 0: Main = 10 / d: d = 2')).toEqual([`${DIVIDE} d`]);
		expect(found('Dim s As String, x As Long', 's = "abc": x = s + 1: s = "1"', 'Main = x')).toEqual([`${COERCE} s`]);
		expect(found('Dim v As Variant', 'v = Array(1, 2): Main = v + 1: v = 5')).toEqual(['variant-value-misuse v']);
		expect(found('Dim v As Variant', 'v = Array(1, 2): v = v + 1', 'Main = v')).toEqual(['variant-value-misuse v']);
		expect(found('Dim v As Variant', 'v = 5: Main = UBound(v): v = 6')).toEqual(['variant-value-misuse v']);
		expect(found('Dim v As Variant', 'v = 5: v = v.Count', 'Main = v')).toEqual(['variant-value-misuse v']);
		expect(found('Dim v As Variant', 'v = Array(1, 2): Main = v(2): v = Array(1, 2, 3)')).toEqual(['array-subscript-out-of-bounds 2']);
		expect(found('Dim n As Long', 'n = -1: Main = Space(n): n = 2')).toEqual(['runtime-argument-value n']);
	});

	it('follows the value inside a loop body and past a block that leaves it alone', () => {
		expect(found('Dim d As Long, i As Long', 'For i = 1 To 1', '    d = 0: Main = 10 / d: d = 2', 'Next i')).toEqual([`${DIVIDE} d`]);
		expect(found('Dim d As Long', 'd = 0', 'Do While False', '    Main = 1', 'Loop', 'Main = 10 / d', 'd = 2')).toEqual([`${DIVIDE} d`]);
	});

	it('stays quiet where another path may have changed the value first', () => {
		// Each runs in Excel: d is 2, or 5, by the time the division runs.
		const quiet = [
			['Dim d As Long, i As Long', 'd = 0', 'For i = 1 To 2', '    If i = 2 Then Main = 10 / d', '    d = 2', 'Next i'],
			['Dim d As Long', 'd = 0', 'If Len("a") = 1 Then d = 2', 'Main = 10 / d'],
			['Dim d As Long', 'd = 0', 'If Len("a") = 1 Then d = 2: Main = 10 / d'],
			['Dim d As Long', 'd = 0', 'GoTo Skip', 'Back:', 'Main = 10 / d', 'Exit Function', 'Skip:', 'd = 2', 'GoTo Back'],
			['Dim d As Long', 'd = 0', 'GoSub SetIt', 'Main = 10 / d', 'Exit Function', 'SetIt:', 'd = 2', 'Return'],
			['Dim d As Long', 'd = 0', 'Fill d', 'Main = 10 / d'],
			['Dim d As Long', 'd = 5', 'Select Case 2', 'Case 1', '    d = 0', 'Case 2', '    Main = 10 / d', 'End Select'],
			['Dim d As Long', 'd = 5', 'If Len("a") = 2 Then', '    d = 0', 'Else', '    Main = 10 / d', 'End If'],
			['Dim v As Variant', 'v = Array(1, 2)', 'v = Array(1, 2, 3)', 'Main = UBound(v) + v(2)'],
		];
		for (const lines of quiet) {
			const src = `${wrap(...lines)}Sub Fill(n As Long)\n    n = 2\nEnd Sub\n`;
			const hits = analyzeModule(src).filter((d) => [DIVIDE, 'variant-value-misuse', 'array-subscript-out-of-bounds'].includes(d.code));
			expect(hits, lines.join(' / ')).toEqual([]);
		}
	});
});
