// Diagnostics tests: built-in functions given the wrong kind of argument
// (issue #242). Each raising case was measured in Excel 16.0 (build 20326,
// 2026-09-30); each quiet neighbour runs clean there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('a string that is no number into a math function', () => {
	it.each(['Abs', 'Sqr', 'Int', 'Fix', 'Round', 'Hex', 'Oct', 'Exp', 'Log', 'Sin', 'Cos', 'Tan', 'Atn'])('reports %s("abc")', (fn) => {
		const src = wrap(`Main = ${fn}("abc")`);
		expectDiagnostic(src, analyzeModule(src), 'runtime-conversion-value', { span: '"abc"', message: [`${fn} cannot convert`, "'13'"] });
	});

	it('stays quiet on a numeric string', () => {
		for (const call of ['Oct("8")', 'Abs("-5")']) {
			expect(byCode(analyzeModule(wrap(`Main = ${call}`)), 'runtime-conversion-value'), call).toHaveLength(0);
		}
		expect(byCode(analyzeModule(wrap('Dim s As String', 's = "3"', 'Main = Abs(s)')), 'runtime-conversion-value')).toHaveLength(0);
	});
});

describe('Null into Sgn', () => {
	it('reports Sgn(Null) as error 94, and leaves Abs, Int and Hex, which pass Null through', () => {
		const src = wrap('Main = IsNull(Sgn(Null))');
		expectDiagnostic(src, analyzeModule(src), 'argument-type-mismatch', { span: 'Null', message: ["'Sgn'", "'94'"] });
		for (const fn of ['Abs', 'Int', 'Hex']) {
			expect(byCode(analyzeModule(wrap(`Main = IsNull(${fn}(Null))`)), 'argument-type-mismatch'), fn).toHaveLength(0);
		}
	});
});

describe('Filter over an array it cannot filter', () => {
	it('reports an array of Long, and one of two dimensions, as error 13', () => {
		const long = wrap('Dim a(1) As Long', 'Main = UBound(Filter(a, "1"))');
		expectDiagnostic(long, analyzeModule(long), 'runtime-argument-value', { span: 'a', message: ['an array of Long', "'13'"] });
		const twoD = wrap('Dim a(1, 1) As String', 'Main = UBound(Filter(a, "1"))');
		expectDiagnostic(twoD, analyzeModule(twoD), 'runtime-argument-value', { span: 'a', message: ['has 2', "'13'"] });
	});

	it('reports a scalar local, as Join does', () => {
		const src = wrap('Dim n As Long', 'Main = UBound(Filter(n, "1"))');
		expectDiagnostic(src, analyzeModule(src), 'runtime-argument-value', { span: 'n', message: ["'n' is declared As Long", "'13'"] });
	});

	it('stays quiet on Strings and Variants', () => {
		for (const lines of [['Main = UBound(Filter(Array("1", "2"), "1"))'], ['Dim a(1) As Variant', 'Main = UBound(Filter(a, "1"))']]) {
			expect(byCode(analyzeModule(wrap(...lines)), 'runtime-argument-value'), lines.join('; ')).toHaveLength(0);
		}
	});
});

describe('a Collection where a built-in reads a value', () => {
	it.each(['Len', 'CStr', 'Val', 'CLng', 'CDbl', 'CInt', 'CBool', 'CDate', 'Trim$', 'UCase$', 'LCase$', 'InStrRev', 'Asc', 'Chr', 'Abs'])('reports %s(New Collection) as a compile error', (fn) => {
		const src = wrap(`Main = ${fn}(New Collection)`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'collection-operand'), 'collection-operand', { span: 'New Collection', message: [fn, 'Argument not optional'] });
	});

	it.each(['InStr', 'Format', 'UCase', 'LCase', 'Trim', 'LTrim', 'RTrim', 'Left', 'Right', 'Mid', 'CVar', 'StrComp', 'Hex'])('reports %s(New Collection) as error 450', (fn) => {
		const args = ['InStr', 'StrComp'].includes(fn) ? ', "a"' : ['Left', 'Right', 'Mid'].includes(fn) ? ', 1' : '';
		const src = wrap(`Main = ${fn}(New Collection${args})`);
		expectDiagnostic(src, byCode(analyzeModule(src), 'object-default-value'), 'object-default-value', { span: 'New Collection', message: [fn, "'450'"] });
	});

	it('reads a Collection variable the same way', () => {
		const len = wrap('Dim c As New Collection', 'Main = Len(c)');
		expectDiagnostic(len, byCode(analyzeModule(len), 'collection-operand'), 'collection-operand', { span: 'c' });
		const instr = wrap('Dim c As New Collection', 'Main = InStr(c, "a")');
		expectDiagnostic(instr, byCode(analyzeModule(instr), 'object-default-value'), 'object-default-value', { span: 'c' });
	});

	it('stays quiet where nothing reads its value, or another rule already reports it', () => {
		for (const call of ['TypeName(New Collection)', 'IsNumeric(New Collection)', 'IsArray(5)']) {
			const diags = analyzeModule(wrap(`Main = ${call}`));
			expect(byCode(diags, 'collection-operand').length + byCode(diags, 'object-default-value').length, call).toBe(0);
		}
		// Left$ types its parameter, which argument-object-type-mismatch reports.
		const left = analyzeModule(wrap('Main = Left$(New Collection, 1)'));
		expect(byCode(left, 'collection-operand')).toHaveLength(0);
		expect(byCode(left, 'argument-object-type-mismatch')).toHaveLength(1);
	});

	it('leaves a member of the Collection, and an object\'s own method, alone', () => {
		expect(byCode(analyzeModule(wrap('Dim c As New Collection', 'Main = Format(c.Count)')), 'object-default-value')).toHaveLength(0);
		const member = `${wrap('Dim o As New Class1', 'Main = o.Format(New Collection)')}`;
		expect(byCode(analyzeModule(member), 'object-default-value')).toHaveLength(0);
	});

	it('leaves a module Function of the same name alone', () => {
		const src = `${wrap('Main = Format(New Collection)')}Private Function Format(x As Variant) As String\nEnd Function\n`;
		expect(byCode(analyzeModule(src), 'object-default-value')).toHaveLength(0);
	});
});
