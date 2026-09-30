// Diagnostics tests: calls to the project's own procedures (issue #223). Each
// case was measured in 64-bit Excel 16.0 (build 20326, 2026-09-30).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { getExcelObjectModel } from '../../src/analyzer/host/excelObjectModel';
import { byCode } from '../helpers/diagnostics';

const HELPERS = [
	'Private Function Two(ByVal a As Long, ByVal b As Long) As Long',
	'End Function',
	'Private Function Opt(ByVal a As Long, Optional ByVal b As Long = 5) As Long',
	'End Function',
	'Private Function Many(ParamArray p() As Variant) As Long',
	'End Function',
	'Private Function TakeColl(ByVal c As Collection) As Long',
	'End Function',
	'Private Function TakeWs(ByVal w As Worksheet) As Long',
	'End Function',
	'Private Function TakeStr(ByVal s As String) As String',
	'End Function',
	'Private Function TakeArr(a() As Long) As Long',
	'End Function',
	'Private Function TakeVar(ByVal v As Variant) As Long',
	'End Function',
	'Private Function TakeBytes(b() As Byte) As Long',
	'End Function',
	'Private Function BytesN(ByVal n As Long) As Byte()',
	'End Function',
	'Private Function Bytes() As Byte()',
	'End Function',
	'Sub TwoSub(ByVal a As Long, Optional ByVal b As Long)',
	'End Sub',
].join('\n');

function codes(...lines: string[]): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n${HELPERS}\n`;
	return analyzeModule(src, { hostModel: getExcelObjectModel() }).filter((d) => d.severity === 'error').map((d) => `${d.code}: ${d.message}`);
}

describe('an object where a value goes, or a value where an object goes (issue #223)', () => {
	it.each([
		['Main = Two(Nothing, 1)', 'Invalid use of object'],
		['Call Two(Nothing, 1)', 'Invalid use of object'],
		['Two Nothing, 1', 'Invalid use of object'],
		['Main = TakeStr(Nothing)', 'Invalid use of object'],
		['Main = Two(New Collection, 1)', 'Argument not optional'],
		['Main = TakeStr(New Collection)', 'Argument not optional'],
		['Main = TakeColl(5)', 'Type mismatch'],
		['Main = TakeColl("x")', 'Type mismatch'],
	])('flags %s at compile time', (line, text) => {
		const hits = codes(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^argument-object-type-mismatch: /);
		expect(hits[0]).toContain(text);
	});

	it.each([
		['Main = Two(Array(1), 1)', 'Array(...), an array'],
		['Main = Two(Split("1"), 1)', 'Split(...), an array'],
		['Main = TakeWs(Range("A1"))', 'not compatible with Worksheet'],
		['Main = TakeWs(ThisWorkbook)', 'not compatible with Worksheet'],
	])('flags %s as error 13', (line, text) => {
		const hits = codes(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toMatch(/^argument-type-mismatch: /);
		expect(hits[0]).toContain(text);
		expect(hits[0]).toContain("'13'");
	});

	it('flags a Long below its range, error 6', () => {
		expect(codes('Main = Two(-2147483649#, 0)')[0]).toContain("'6'");
		expect(codes('Main = Two(-2147483648#, 0)')).toHaveLength(0);
	});

	it.each([
		'Main = TakeVar(New Collection)',
		'Main = TakeColl(Nothing)',
		'Main = TakeWs(ActiveSheet)',
		'Main = TakeWs(Nothing)',
		'Main = TakeWs(Sheets(1))',
		'Main = TakeStr(Range("A1"))',
		'Main = Two(Array(1)(0), 1)',
		'Dim r As Range: Set r = Range("A1"): Main = TakeWs(r)',
	])('stays quiet on %s', (line) => {
		expect(codes(line)).toHaveLength(0);
	});
});

describe('argument lists the VBE refuses (issue #223)', () => {
	it.each([
		'Main = Opt(1, )',
		'Main = Many(1, )',
		'Main = Two(1, 2, )',
		'TwoSub 1,',
		'Call TwoSub(1, )',
		'Main = Left("ab", )',
	])('flags the trailing empty argument in %s', (line) => {
		const hits = codes(line);
		expect(hits).toHaveLength(1);
		expect(hits[0]).toContain('ends in an empty argument');
	});

	it('flags a named ParamArray', () => {
		expect(codes('Main = Many(p:=1)')[0]).toContain('Argument in ParamArray may not be named');
	});

	it.each([
		'Main = Opt(1)',
		'Main = Many(1, 2)',
		'Main = Many(1, , 2)',
		'Debug.Print 1,',
	])('stays quiet on %s', (line) => {
		expect(codes(line)).toHaveLength(0);
	});
});

describe('an array element where an array goes (issue #223)', () => {
	it('flags a(0), and takes a and a()', () => {
		const hits = codes('Dim a(1) As Long', 'Main = TakeArr(a(0))');
		expect(hits).toHaveLength(1);
		expect(hits[0]).toContain("one element of the array 'a'");
		expect(codes('Dim a(1) As Long', 'Main = TakeArr(a)')).toHaveLength(0);
		expect(codes('Dim a(1) As Long', 'Main = TakeArr(a())')).toHaveLength(0);
	});

	it('takes the result of a Function returning an array', () => {
		expect(codes('Main = TakeBytes(Bytes())')).toHaveLength(0);
		expect(codes('Main = TakeBytes(BytesN(1))')).toHaveLength(0);
	});
});
