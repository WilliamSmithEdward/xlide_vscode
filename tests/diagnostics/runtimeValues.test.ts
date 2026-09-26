// Diagnostics tests: runtime argument and conversion values (issue #118).
// Every raising call below was measured in Excel 16.0 (build 20326,
// 2026-09-26): it compiles and raises the error named every time it runs. The
// quiet neighbours - Mid("abc", 10), CLng("&H10"), CBool("1"), Round(1.5, 0),
// Weekday(Date, 7), Environ(1) - run clean there and stay quiet here.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const ARG = 'runtime-argument-value';
const CONV = 'runtime-conversion-value';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('runtime-argument-value - error 5 arguments (issue #118)', () => {
	it.each([
		['Main = Asc("")', '""', 'Asc'],
		['Main = String(3, "")', '""', 'String'],
		['Main = Sqr(-4)', '-4', 'Sqr'],
		['Main = Log(0)', '0', 'Log'],
		['Main = MonthName(13)', '13', 'MonthName'],
		['Main = WeekdayName(0)', '0', 'WeekdayName'],
		['Main = Round(1.5, -1)', '-1', 'Round'],
		['Main = Weekday(Date, 8)', '8', 'Weekday'],
		['Main = DateAdd("x", 1, Date)', '"x"', 'DateAdd'],
		['Main = DateSerial(10000, 1, 1)', '10000', 'DateSerial'],
		['Main = InStrRev("abc", "b", 0)', '0', 'InStrRev'],
		['Main = Split("a,b", ",", -2)', '-2', 'Split'],
		['Main = FormatNumber(1, -2)', '-2', 'FormatNumber'],
		['Main = Environ(0)', '0', 'Environ'],
		['Main = InStr(1, "abc", "b", vbDatabaseCompare)', 'vbDatabaseCompare', 'InStr'],
	])('flags %s', (statement, span, name) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), ARG, { severity: 'error', span, message: [name, "Run-time error '5'"] });
	});

	it('reads a String local the procedure never assigns as ""', () => {
		const src = wrap('Dim s As String', 'Main = Asc(s)');
		expectDiagnostic(src, analyzeModule(src), ARG, { span: 's', message: ['Asc', 'never given another value'] });
	});

	it('flags a Mid statement that starts past the end of a string it knows', () => {
		const src = wrap('Dim s As String', 's = "abc"', 'Mid(s, 5, 1) = "x"', 'Main = s');
		expectDiagnostic(src, analyzeModule(src), ARG, { span: '5', message: ['Mid statement', '3 character'] });
		const inRange = wrap('Dim s As String', 's = "abc"', 'Mid(s, 3, 1) = "x"', 'Main = s');
		expect(byCode(analyzeModule(inRange), ARG)).toHaveLength(0);
	});

	it('flags DateAdd past December 31, 9999 with a date literal', () => {
		const src = wrap('Main = DateAdd("d", 1, #12/31/9999#)');
		expectDiagnostic(src, analyzeModule(src), ARG, { span: '#12/31/9999#', message: 'DateAdd' });
		expect(byCode(analyzeModule(wrap('Main = DateAdd("d", -1, #12/31/9999#)')), ARG)).toHaveLength(0);
	});

	it('flags Err.Raise and Error with a number outside 1 to 65535', () => {
		const src = wrap('Err.Raise 0', 'Err.Raise 65536', 'Error 0', 'Err.Raise 5', 'Error 5');
		expectDiagnostics(src, analyzeModule(src), ARG, [
			{ span: '0', message: 'Err.Raise 0' },
			{ span: '65536', message: 'Err.Raise 65536' },
			{ span: '0', message: 'Error 0' },
		]);
	});

	it('flags a negative base with a fractional power and zero with a negative power', () => {
		const src = wrap('Main = (-8) ^ (1 / 3)', 'Main = 0 ^ -1', 'Main = 2 ^ -1', 'Main = (-8) ^ 3');
		expectDiagnostics(src, analyzeModule(src), ARG, [
			{ span: '^', message: 'negative number raised to the fractional power' },
			{ span: '^', message: 'Zero raised to the negative power' },
		]);
	});

	it('flags an invalid Like pattern as error 93', () => {
		const src = wrap('Main = "b" Like "[z-a]"', 'Main = "b" Like "[a-"', 'Main = "b" Like "[a-z]"', 'Main = "b" Like "[!a-c]*"');
		expectDiagnostics(src, analyzeModule(src), ARG, [
			{ span: '"[z-a]"', message: ['reversed range z-a', "error '93'"] },
			{ span: '"[a-"', message: 'never closes' },
		]);
	});

	it('stays quiet on the values that run', () => {
		const src = wrap(
			'Main = Mid("abc", 10) & CLng("&H10") & CBool("1")',
			'Main = Round(1.5, 0) + Weekday(Date, 7) + Len(Environ(1))',
			'Main = Sqr(4) + Log(0.5) + MonthName(12) & WeekdayName(7)',
			'Main = InStrRev("abc", "b", -1) + InStrRev("abc", "b", 1)',
			'Main = InStr(1, "abc", "b", vbTextCompare)',
			'Main = DateSerial(9999, 12, 31)',
		);
		const diagnostics = analyzeModule(src);
		expect(byCode(diagnostics, ARG)).toHaveLength(0);
		expect(byCode(diagnostics, CONV)).toHaveLength(0);
	});

	it('allows vbDatabaseCompare where Access is the host', () => {
		const src = wrap('Main = InStr(1, "abc", "b", vbDatabaseCompare)');
		expect(byCode(analyzeModule(src, { host: 'access' }), ARG)).toHaveLength(0);
	});
});

describe('runtime-conversion-value - error 13 conversions (issue #118)', () => {
	it.each([
		['Main = CLng("abc")', '"abc"', 'CLng', 'a number'],
		['Main = CBool("yes")', '"yes"', 'CBool', 'Boolean'],
		['Main = CDbl("")', '""', 'CDbl', 'a number'],
		['Main = DateValue("abc")', '"abc"', 'DateValue', 'Date'],
	])('flags %s', (statement, span, name, target) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), CONV, { severity: 'error', span, message: [name, target, "Run-time error '13'"] });
	});

	it('stays quiet on strings the conversions read', () => {
		const src = wrap('Main = CLng("&H10") + CInt(" 12 ") + CDbl("1e3")', 'Main = CBool("True") Or CBool("0")', 'Main = CDate("2020-01-01")');
		expect(byCode(analyzeModule(src), CONV)).toHaveLength(0);
	});
});
