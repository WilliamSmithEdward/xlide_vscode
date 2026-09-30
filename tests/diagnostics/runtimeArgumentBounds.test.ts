// Diagnostics tests: runtime arguments issue #218 measured. Every raising call
// below was run in Excel 16.0 (build 20326, 2026-09-30) and raises the error
// named every time; each quiet neighbour runs clean there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const ARG = 'runtime-argument-value';
const CONV = 'runtime-conversion-value';
const OVERFLOW = 'arithmetic-overflow';
const SHAPE = 'argument-shape-mismatch';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function codes(statement: string, code: string): number {
	return byCode(analyzeModule(wrap(statement)), code).length;
}

describe('Compare is 0, 1 or a locale ID (issue #218)', () => {
	it.each([
		'Main = InStr(1, "abc", "b", 3)',
		'Main = StrComp("a", "b", 1033)',
		'Main = StrComp("a", "b", 16384.4)',
		'Main = InStr(1, "abc", "b", -0.4)',
		'Main = InStrRev("abc", "b", -1, 2)',
		'Main = Replace("abc", "b", "x", 1, -1, 2)',
		'Main = UBound(Split("a,b", ",", -1, 2))',
		'Main = UBound(Filter(Array("a", "b"), "a", True, 1033))',
	])('stays quiet on %s', (statement) => {
		expect(codes(statement, ARG)).toBe(0);
	});

	it.each([
		'Main = InStr(1, "abc", "b", 2)',
		'Main = InStr(1, "abc", "b", 2.4)',
		'Main = StrComp("a", "b", -1)',
		'Main = UBound(Filter(Array("a", "b"), "a", True, 2))',
		'Main = UBound(Filter(Array("a", "b"), "a", True, -1))',
		'Main = InStrRev("abc", "b", -1, -1)',
		'Main = UBound(Split("a,b", ",", -1, -1))',
	])('flags %s', (statement) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), ARG, { message: ['Compare', "Run-time error '5'"] });
	});

	it('takes a Compare of 2 in InStr and StrComp where Access is the host', () => {
		const src = wrap('Main = InStr(1, "abc", "b", 2) + StrComp("a", "b", 2)');
		expect(byCode(analyzeModule(src, { host: 'access' }), ARG)).toHaveLength(0);
	});
});

describe('bounds and overflow of single arguments (issue #218)', () => {
	it.each([
		['Main = IsError(CVErr(65536))', '65536', 'CVErr', '5'],
		['Main = IsError(CVErr(-1))', '-1', 'CVErr', '5'],
		['Main = LeftB("abc", -1)', '-1', 'LeftB', '5'],
		['Main = RightB$("abc", -1)', '-1', 'RightB$', '5'],
		['Main = MidB("abc", 0.4)', '0.4', 'MidB', '5'],
		['Main = AscB("")', '""', 'AscB', '5'],
		['Main = InStrB(0, "abc", "b")', '0', 'InStrB', '5'],
		['Main = Left("abc", 1073741824)', '1073741824', 'Left', '5'],
		['Main = Mid("abc", 1073741825)', '1073741825', 'Mid', '5'],
		['Main = Len(String(1073741824, "a"))', '1073741824', 'String', '5'],
		['Main = Len(String(1E+10, "a"))', '1E+10', 'String', '6'],
		['Main = InStr(1E+10, "abc", "b")', '1E+10', 'InStr', '6'],
		['Main = TimeSerial(32768, 0, 0)', '32768', 'TimeSerial', '6'],
		['Main = TimeSerial(0, 0, -32769)', '-32769', 'TimeSerial', '6'],
		['Main = DateSerial(2020, 32768, 1)', '32768', 'DateSerial', '6'],
		['Main = ChrB(255.5)', '255.5', 'ChrB', '6'],
		['Main = Error(65536)', '65536', 'Error', '6'],
	])('flags %s', (statement, span, name, error) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), ARG, { span, message: [name, `Run-time error '${error}'`] });
	});

	it.each([
		'Main = IsError(CVErr(65535))',
		'Main = IsError(CVErr(-0.5))',
		'Main = LeftB("abc", 0)',
		'Main = InStrB(1, "abc", "b")',
		'Main = Left("abc", 1073741823)',
		'Main = Mid("abc", 1073741824)',
		'Main = TimeSerial(32767, 0, 0)',
		'Main = ChrB(-0.5)',
		'Main = Error(65535)',
		'Main = Error(-70000)',
		'Main = LeftB("abc", 2147483647)',
	])('stays quiet on %s', (statement) => {
		expect(codes(statement, ARG)).toBe(0);
	});

	it('reports DateSerial past the Integer range as the overflow, not a date past 9999', () => {
		const src = wrap('Main = DateSerial(32768, 1, 1)');
		expectDiagnostic(src, analyzeModule(src), ARG, { span: '32768', message: "Run-time error '6'" });
	});

	it('leaves a value past the Long range to the typed signature', () => {
		const src = wrap('Main = Left("abc", 1E+10)');
		expect(byCode(analyzeModule(src), ARG)).toHaveLength(0);
		expect(byCode(analyzeModule(src), 'argument-type-mismatch')).toHaveLength(1);
	});

	it('reads the Error statement once, as the statement', () => {
		const src = wrap('Error (70000)');
		const hits = byCode(analyzeModule(src), ARG);
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('Error 70000');
	});
});

describe('checks across arguments (issue #218)', () => {
	it.each([
		['Main = Partition(5, -1, 10, 2)', 'Start'],
		['Main = Partition(5, 10, 10, 1)', 'Stop'],
		['Main = Partition(5, 0, 0.4, 1)', 'Stop'],
		['Main = Partition(5, 0, 10, 0.4)', 'Interval'],
		['Main = Pmt(0.1, 0, 1000)', 'NPer 0'],
		['Main = IPmt(0.1, 0, 10, 1000)', 'Per is 0'],
		['Main = PPmt(0.1, 11, 10, 1000)', 'past the last'],
		['Main = IPmt(0.1, 1, -10, 1000)', 'NPer is -10'],
		['Main = SLN(1000, 100, 0)', 'Life of 0'],
		['Main = SYD(1000, 100, 5, 5.4)', 'past its Life'],
		['Main = SYD(1000, 100, -1, 1)', 'Life is -1'],
		['Main = DDB(1000, 100, 5, 0)', 'Period is 0'],
		['Main = DDB(1000, 100, 5, 1, 0)', 'Factor is 0'],
		['Main = NPer(-1, 100, 1000)', 'Rate is -1'],
		['Main = NPer(0, 0, 1000)', 'Rate and a Pmt of 0'],
		['Main = NPer(0.1, -100, 1000)', 'logarithm'],
		['Main = NPer(0.1, -50, 1000)', 'logarithm'],
		['Main = NPer(0.1, 0, 1000)', 'logarithm'],
		['Main = Rate(0, -100, 1000)', 'NPer is 0'],
	])('flags %s', (statement, text) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), ARG, { message: [text, "Run-time error '5'"] });
	});

	it.each([
		['Main = PV(-1, 10, 100)', "'11': Division by zero"],
		['Main = LBound(Array(1), 0)', "'9': Subscript out of range"],
		['Main = UBound(Split("a"), 2)', "'9': Subscript out of range"],
		['Main = Join("abc", ",")', "'13': Type mismatch"],
		['Main = UBound(Filter(5, "a"))', "'13': Type mismatch"],
	])('flags %s', (statement, error) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), ARG, { message: error });
	});

	it.each([
		'Main = Partition(5, 0, 10, 1)',
		'Main = Partition(5, -0.4, 10, 0.6)',
		'Main = Partition(-5, 0, 10, 1)',
		'Main = Pmt(0, 10, 1000)',
		'Main = Pmt(0.1, -10, 1000)',
		'Main = IPmt(0.1, 0.4, 10, 1000)',
		'Main = PPmt(0.1, 10.5, 10, 1000)',
		'Main = SLN(1000, 100, -1)',
		'Main = SYD(1000, 100, 5, 5)',
		'Main = DDB(1000, 100, 5, 0.5, 0.5)',
		'Main = NPer(0, 100, 1000)',
		'Main = NPer(0.1, 100, 1000)',
		'Main = NPer(0.1, -150, 1000)',
		'Main = NPer(0.1, -100, 1000, 0, 1)',
		'Main = Rate(10, 100, 1000)',
		'Main = Rate(10, -150, 1000)',
		'Main = FV(0.1, 10, 100, 0, 2)',
		'Main = FV(-1, 10, 100)',
		'Main = UBound(Array(1, 2), 1)',
		'Main = LBound(Array(), 1)',
		'Main = Join(Array(1), ",")',
		'Main = Choose(0, 1, 2)',
		'Main = Switch(False, 1)',
		'Main = Hex(-1)',
	])('stays quiet on %s', (statement) => {
		expect(codes(statement, ARG)).toBe(0);
	});

	it('does not judge LBound of an array it cannot see the rank of', () => {
		const src = wrap('Dim a(1, 1) As Long', 'Main = UBound(a, 2)');
		expect(byCode(analyzeModule(src), ARG)).toHaveLength(0);
		const call = `${wrap('Main = UBound(Grid(), 2)')}Function Grid() As Variant\n    Dim g(1, 1) As Long\n    Grid = g\nEnd Function\n`;
		expect(byCode(analyzeModule(call), ARG)).toHaveLength(0);
	});
});

describe('arguments converted to a Date or a number (issue #218)', () => {
	it.each([
		['Main = Month("")', '""', 'Month'],
		['Main = Year("abc")', '"abc"', 'Year'],
		['Main = Weekday("", 1)', '""', 'Weekday'],
		['Main = DateAdd("m", 1, "abc")', '"abc"', 'DateAdd'],
		['Main = DatePart("m", "abc")', '"abc"', 'DatePart'],
		['Main = DateDiff("d", #1/1/2020#, "")', '""', 'DateDiff'],
		['Main = Year(3000000)', '3000000', 'Year'],
		['Main = Year(2958466)', '2958466', 'Year'],
		['Main = Day(-657435)', '-657435', 'Day'],
		['Main = DateAdd("m", 1, 3000000)', '3000000', 'DateAdd'],
		['Main = Sgn("x")', '"x"', 'Sgn'],
		['Main = Sgn("")', '""', 'Sgn'],
	])('flags %s', (statement, span, name) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), CONV, { span, message: [name, "Run-time error '13'"] });
	});

	it.each([
		'Main = Month("1/2/2020")',
		'Main = Minute("12:30")',
		'Main = Year(2958465.9)',
		'Main = Year(-657434.9)',
		'Main = Hour(2958465.99999)',
		'Main = DateAdd("d", 1, #1/1/2020#)',
		'Main = Sgn(" 5 ")',
		'Main = Sgn(True)',
	])('stays quiet on %s', (statement) => {
		expect(codes(statement, CONV)).toBe(0);
	});
});

describe('arithmetic-overflow: Abs and CDec (issue #218)', () => {
	it('keeps Abs of the smallest Long quiet, which hands it back unchanged', () => {
		expect(codes('Main = Abs(CLng(-2147483647 - 1))', OVERFLOW)).toBe(0);
		const src = wrap('Main = Abs(CInt(-32768))');
		expectDiagnostic(src, analyzeModule(src), OVERFLOW, { message: 'Abs(-32768)' });
	});

	it.each([
		['Main = CDec("1E30")'],
		['Main = CDec(1E+30)'],
		['Main = CDec("79228162514264337593543950336")'],
	])('flags %s', (statement) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), OVERFLOW, { message: ['CDec', 'Decimal'] });
	});

	it.each([
		'Main = CDec("1E28")',
		'Main = CDec("79228162514264337593543950335")',
	])('stays quiet on %s', (statement) => {
		expect(codes(statement, OVERFLOW)).toBe(0);
	});
});

describe('IRR, MIRR and NPV take an array of Double (issue #218)', () => {
	it.each([
		['Main = IRR(Array(-100#, 60#, 60#))', 'Array(-100#, 60#, 60#)'],
		['Main = NPV(0.1, Array(-100#, 60#))', 'Array(-100#, 60#)'],
		['Main = MIRR(Array(-100#, 60#, 60#), 0.1, 0.1)', 'Array(-100#, 60#, 60#)'],
		['Main = IRR(ValueArray:=Array(-100#, 60#, 60#))', 'Array(-100#, 60#, 60#)'],
	])('flags %s', (statement, span) => {
		const src = wrap(statement);
		expectDiagnostic(src, analyzeModule(src), SHAPE, { span, message: 'array or user-defined type expected' });
	});

	it.each([
		['Dim v As Variant', 'v'],
		['Dim v(2) As Variant', 'v'],
		['Dim v(2) As Long', 'v'],
		['Dim v(2) As Double', '(v)'],
	])('flags IRR of %s', (declaration, argument) => {
		const src = wrap(declaration, `Main = IRR(${argument})`);
		expectDiagnostic(src, analyzeModule(src), SHAPE, { span: 'v', message: 'array or user-defined type expected' });
	});

	it('flags an array parameter given (d), in a call and as a statement', () => {
		const procs = 'Private Function TakeD(a() As Double) As Long\nEnd Function\nPrivate Sub PutD(a() As Double)\nEnd Sub\n';
		for (const statement of ['Main = TakeD((d))', 'PutD (d)', 'Call PutD((d))']) {
			const src = wrap('Dim d(2) As Double', statement) + procs;
			expectDiagnostic(src, analyzeModule(src), SHAPE, { span: 'd', message: 'in parentheses' });
		}
		const bare = wrap('Dim d(2) As Double', 'PutD d', 'Main = TakeD(d)') + procs;
		expect(byCode(analyzeModule(bare), SHAPE)).toHaveLength(0);
	});

	it.each([
		['Dim d(2) As Double', 'IRR(d)'],
		['Dim d(2) As Double', 'IRR(d())'],
		['Dim d() As Double', 'IRR(d)'],
		['Dim d(1) As Double', 'NPV(0.1, d)'],
	])('stays quiet on %s then %s', (declaration, call) => {
		const src = wrap(declaration, `Main = ${call}`);
		expect(byCode(analyzeModule(src), SHAPE)).toHaveLength(0);
	});
});
