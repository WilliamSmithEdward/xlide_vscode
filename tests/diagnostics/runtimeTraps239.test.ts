// Diagnostics tests: the runtime traps issue #239 found unreported. Each
// raising case was measured in Excel 16.0 (build 20326, 2026-09-30) and
// raises the error named every time it runs; each quiet neighbour runs clean
// there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { isInvalidDateString } from '../../src/analyzer/diagnostics/stringConversion';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function count(code: string, ...lines: string[]): number {
	return byCode(analyzeModule(wrap(...lines)), code).length;
}

describe('Null carried through a Variant local', () => {
	it('reports the Null a Variant holds going into a String or a Long', () => {
		const src = wrap('Dim v As Variant, s As String, n As Long', 'v = Null', 's = v', 'n = v', 'If Len("x") = 1 Then s = v');
		const hits = byCode(analyzeModule(src), 'assignment-type-mismatch');
		expect(hits).toHaveLength(3);
		expectDiagnostic(src, [hits[0]], 'assignment-type-mismatch', { span: 'v', message: ["'v' holds Null here", "'94'"] });
	});

	it('stays quiet where the Variant holds something else, or the target is a Variant', () => {
		expect(count('assignment-type-mismatch', 'Dim v As Variant, w As Variant', 'v = Null', 'w = v', 'Main = IsNull(w)')).toBe(0);
		expect(count('assignment-type-mismatch', 'Dim v As Variant, s As String', 'v = Null', 'v = "x"', 's = v')).toBe(0);
		expect(count('assignment-type-mismatch', 'Dim v As Variant, s As String', 'v = Null', 'If Len("x") = 1 Then', '    v = "y"', 'End If', 's = v')).toBe(0);
		expect(count('assignment-type-mismatch', 'Dim v As Variant, s As String', 'v = Null', 'If Len("x") = 1 Then v = "y"', 's = v')).toBe(0);
	});
});

describe('\\ and Mod by a local that rounds to 0', () => {
	it('reports a Double local holding 0.5 or less', () => {
		const src = wrap('Dim d As Double', 'd = 0.4', 'Main = 10 Mod d', 'Main = 10 \\ d');
		expect(byCode(analyzeModule(src), 'division-by-zero')).toHaveLength(2);
		expect(count('division-by-zero', 'Dim d As Double', 'd = 0.5', 'Main = 10 Mod d')).toBe(1);
		expect(count('division-by-zero', 'Dim d As Double', 'd = -0.4', 'Main = 10 Mod d')).toBe(1);
	});

	it('stays quiet above 0.5 either side of 0, with /, and behind a guard', () => {
		expect(count('division-by-zero', 'Dim d As Double', 'd = 0.6', 'Main = 10 Mod d')).toBe(0);
		expect(count('division-by-zero', 'Dim d As Double', 'd = -5.5', 'Main = 10 Mod d')).toBe(0);
		expect(count('division-by-zero', 'Dim d As Double', 'd = 0.4', 'Main = 10 / d')).toBe(0);
		expect(count('division-by-zero', 'Dim d As Double', 'd = 0.4', 'If d <> 0 Then Main = 10 / d')).toBe(0);
	});
});

describe('a year-first date that names no day', () => {
	it('refuses a day the calendar does not have, in either order after the year', () => {
		for (const text of ['2020-02-30', '2021-02-29', '2020-04-31', '1900-02-29', '2020-00-10']) {
			expect(isInvalidDateString(text), text).toBe(true);
		}
		// A month past 12 still reads as a day in the other order, and a
		// two-digit year is no year first: "20-02-30" runs as 2030-02-20.
		for (const text of ['2020-02-29', '2000-02-29', '2020-13-01', '2020-31-12', '20-02-30']) {
			expect(isInvalidDateString(text), text).toBe(false);
		}
	});

	it('reports DateValue and CDate of 30 February', () => {
		const src = wrap('Main = DateValue("2020-02-30")', 'Main = CDate("2020-02-30")');
		expect(byCode(analyzeModule(src), 'runtime-conversion-value')).toHaveLength(2);
	});
});

describe('Join of a value that is no array', () => {
	it('reports a literal, Null, True, a date and a scalar local', () => {
		for (const value of ['5', '"a,b"', 'Null', 'True', '#1/1/2000#']) {
			const src = wrap(`Main = Join(${value})`);
			expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { span: value, message: ['Join takes an array', "'13'"] });
		}
		const local = wrap('Dim n As Long', 'Main = Join(n)');
		expectDiagnostic(local, analyzeModule(local), 'runtime-argument-value', { span: 'n', message: "'n' is declared As Long" });
		const held = wrap('Dim v As Variant', 'v = 5', 'Main = Join(v)');
		expectDiagnostic(held, analyzeModule(held), 'variant-value-misuse', { span: 'v', message: 'is not an array' });
	});

	it('stays quiet on an array', () => {
		expect(count('runtime-argument-value', 'Main = Join(Split("a b"), "-")')).toBe(0);
		expect(count('runtime-argument-value', 'Dim a(1) As String', 'Main = Join(a, "-")')).toBe(0);
	});
});

describe('an array a call returns, used as a scalar', () => {
	it.each(['Array(1) + 1', '1 + Array(1)', 'Array(1) & "x"', '-Array(1)', 'Not Array(1)', '(Array(1) = 1)', 'Split("a") + 1', 'VBA.Array(1) + 1', 'Array(1) Like "x"'])('reports %s', (expression) => {
		const src = wrap(`Main = ${expression}`);
		expectDiagnostic(src, analyzeModule(src), 'variant-value-misuse', { message: ['returns an array', "'13'"] });
	});

	it('stays quiet on an element, an assignment and an argument', () => {
		expect(count('variant-value-misuse', 'Main = Split("a b")(0) & "x"', 'Main = 1 + Array(5)(0)')).toBe(0);
		expect(count('variant-value-misuse', 'Dim c As New Collection', 'c.Add Array(1)', 'Main = Array(1)')).toBe(0);
		expect(count('variant-value-misuse', 'If IsArray(Array(1)) Then Main = Array(1)')).toBe(0);
		// The module's own Split is not VBA's.
		const shadowed = `${wrap('Main = Split("a") + 1')}Function Split(x As String) As Long\n    Split = 1\nEnd Function\n`;
		expect(byCode(analyzeModule(shadowed), 'variant-value-misuse')).toHaveLength(0);
	});
});

describe('For Each over what the grammar refuses', () => {
	it.each(['5', '"abc"', '5.5', '#1/1/2000#', 'Null', 'True', 'Nothing', 'Empty', '-v', '(v)', 'New Collection', 'Not v', 'v & v', '5 + 1', 'v Is v', 'Len("abc")'])('reports In %s as a Syntax error', (expression) => {
		const src = wrap('Dim x As Variant, v As Variant', `For Each x In ${expression}`, 'Next');
		expectDiagnostic(src, byCode(analyzeModule(src), 'malformed-statement'), 'malformed-statement', { span: expression, message: 'Syntax error' });
	});

	it('names a date literal as the literal it is', () => {
		const src = wrap('Dim x As Variant', 'For Each x In #1/1/2000#', 'Next');
		expectDiagnostic(src, analyzeModule(src), 'malformed-statement', { message: 'the literal #1/1/2000#' });
	});

	it('stays quiet on a variable, a member chain, a call and a bracketed name', () => {
		for (const expression of ['v', 'c.Items', 'c!Items', 'Split("a" & "b")', 'VBA.Array(1, 2)', '[A1:B2]', 'ThisWorkbook.Worksheets(1).Range("A1:B2")']) {
			expect(count('malformed-statement', 'Dim x As Variant, v As Variant, c As Object', `For Each x In ${expression}`, 'Next'), expression).toBe(0);
		}
	});
});
