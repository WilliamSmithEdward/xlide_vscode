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
		['Main = DateSerial(10000, 1, 1)', '10000, 1, 1', 'DateSerial'],
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
		// Each on its own: a line after one that raises never runs (issue #430).
		for (const [line, message] of [['Err.Raise 0', 'Err.Raise 0'], ['Err.Raise 65536', 'Err.Raise 65536'], ['Error 0', 'Error 0']]) {
			const src = wrap(line);
			expectDiagnostics(src, analyzeModule(src), ARG, [{ span: line.split(' ').pop()!, message }]);
		}
		expect(byCode(analyzeModule(wrap('Err.Raise 5')), ARG)).toHaveLength(0);
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

	it('flags a bad Like list only when matching reaches it with a character left (issue #193)', () => {
		// Measured in Excel 16.0: each flagged one raises 93, each quiet one is False.
		const flagged = ['"xy" Like "?["', '"ab" Like "*["', '"ab" Like "[a-z]["', '"b" Like "[a"', '"ab" Like "a[z-a]"', '"1b" Like "#["', '"xb" Like "[!a]["'];
		for (const expression of flagged) {
			const src = wrap(`Main = ${expression}`);
			expectDiagnostic(src, analyzeModule(src), ARG, { message: "error '93'" });
		}
		// Not reached, or a string the code does not make plain. A "*" right before
		// the list reaches it with any character left (issue #336).
		const quiet = wrap(
			'Dim s As String',
			'Main = "x" Like "?["',
			'Main = "b" Like "[a-z]["',
			'Main = "a" Like "a[z-a]"',
			'Main = "" Like "[a"',
			'Main = "zb" Like "a[z-a]"',
			'Main = "Ab" Like "a["',
			'Main = InputBox("x") Like "[a"',
		);
		expect(byCode(analyzeModule(quiet), ARG)).toHaveLength(0);
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

	it('rounds a fractional argument half to even before the range, as VBA does (issue #189)', () => {
		// Each runs in Excel 16.0: the argument VBA passes is in range.
		const quiet = wrap(
			'Main = Space(-0.5)',
			'Main = Left("abc", -0.4)',
			'Main = String(-0.5, "a")',
			'Main = Mid("abc", 0.6)',
			'Main = InStr(0.6, "abc", "b")',
			'Main = Chr(255.4)',
			'Main = Chr(-0.5)',
			'Main = MonthName(12.4)',
			'Main = MonthName(1.5)',
			'Main = Round(1.25, -0.4)',
			'Main = Environ(0.6)',
			'Main = Weekday(Date, 7.4)',
			'Main = Log(0.4)',
		);
		expect(byCode(analyzeModule(quiet), ARG)).toHaveLength(0);
		// These round out of range, and a Double parameter is not rounded.
		for (const [statement, shown] of [
			['Main = Space(-0.6)', '-0.6, which VBA rounds to -1'],
			['Main = Mid("abc", 0.5)', '0.5, which VBA rounds to 0'],
			['Main = Chr(255.5)', '255.5, which VBA rounds to 256'],
			['Main = MonthName(0.5)', '0.5, which VBA rounds to 0'],
			['Main = WeekdayName(7.5)', '7.5, which VBA rounds to 8'],
			['Main = Sqr(-0.4)', 'is -0.4;'],
		]) {
			const src = wrap(statement);
			expectDiagnostic(src, analyzeModule(src), ARG, { message: shown });
		}
	});

	it('judges the whole date DateSerial makes, and the Compare, ChrW and Round edges (issue #189)', () => {
		const quiet = wrap(
			'Main = DateSerial(10000, 0, 1)',
			'Main = DateSerial(10000, -11, 1)',
			'Main = DateSerial(9999, 12, 31)',
			'Main = ChrW(-32768)',
			'Main = StrComp("a", "b", 1)',
			'Main = Replace("a", "a", "b", 1, -1, 2)',
			'Main = Round(1.5, 22)',
		);
		expect(byCode(analyzeModule(quiet), ARG)).toHaveLength(0);
		for (const statement of [
			'Main = DateSerial(9999, 13, 1)',
			'Main = DateSerial(9999, 12, 400)',
			'Main = DateSerial(9999, 12, 32)',
			'Main = ChrW(-32769)',
			'Main = StrComp("a", "b", -1)',
			'Main = StrComp("a", "b", 2)',
			'Main = Replace("a", "a", "b", 1, -1, -1)',
			'Main = Round(1.5, 23)',
		]) {
			const src = wrap(statement);
			expect(byCode(analyzeModule(src), ARG), statement).toHaveLength(1);
		}
	});

	it('flags a StrConv Conversion no locale accepts (issue #184)', () => {
		// Measured with the English, Japanese (1041) and Chinese (2052) LCIDs.
		for (const value of ['99', '65', '192', '3 + 256', '48', '12', '-1']) {
			const src = wrap(`Main = StrConv("a", ${value})`);
			expectDiagnostic(src, analyzeModule(src), ARG, { span: value, message: ["'Conversion' of 'StrConv'", "Run-time error '5'"] });
		}
		// vbWide 4, vbNarrow 8, vbKatakana 16 and vbHiragana 32 run under an
		// East Asian locale, alone or with a case: they are not judged.
		const quiet = wrap(
			'Main = StrConv("a", 0)',
			'Main = StrConv("a", vbProperCase)',
			'Main = StrConv("a", vbUnicode)',
			'Main = StrConv("a", vbFromUnicode)',
			'Main = StrConv("a", vbWide)',
			'Main = StrConv("a", 36)',
			'Main = StrConv("a", vbUpperCase + vbHiragana)',
		);
		expect(byCode(analyzeModule(quiet), ARG)).toHaveLength(0);
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

describe('Err.Raise takes negative numbers, Error does not (issue #142)', () => {
	// Measured in Excel 16.0 (build 20326, 2026-09-26): Err.Raise raises
	// -2147220991, -1000 and the &H80004002 HRESULT (Excel maps it to 430);
	// 0 and 65536 raise 5. Error -1 raises 5, Error 65535 raises 65535.
	it('stays quiet on vbObjectError + n, a negative literal and a hex HRESULT Const', () => {
		const src =
			'Option Explicit\nPrivate Const E_NOINTERFACE As Long = &H80004002\n' +
			'Sub Main()\n' +
			'    Err.Raise vbObjectError + 513, "Raised", "custom"\n' +
			'    Err.Raise -1000, "Raised", "negative"\n' +
			'    Err.Raise E_NOINTERFACE\n' +
			'    Err.Raise -2147483648#\n' +
			'End Sub\n';
		expect(byCode(analyzeModule(src), ARG)).toHaveLength(0);
	});

	it('reports 0 and 65536 for Err.Raise and -1 for Error', () => {
		for (const [line, span] of [['Err.Raise 0', '0'], ['Err.Raise 65536', '65536'], ['Error -1', '-1']]) {
			const src = `Option Explicit\nSub Main()\n    ${line}\nEnd Sub\n`;
			expectDiagnostic(src, analyzeModule(src), ARG, { severity: 'error', span, message: "'5'" });
		}
		const quiet = 'Option Explicit\nSub Main()\n    Error 65535\nEnd Sub\n';
		expect(byCode(analyzeModule(quiet), ARG)).toHaveLength(0);
	});
});
