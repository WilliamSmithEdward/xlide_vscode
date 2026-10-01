// Diagnostics tests: Enum members and String Consts in the value rules, and
// Const declarations the VBE refuses (issue #255). Every case was measured
// in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { rawExpressionTokens } from '../../src/analyzer/diagnostics/walker';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const ENUM = 'Private Enum E\n    eZ\n    eBig = 40000\n    eMax = 2147483647\n    eK = 100000\n    eNeg = -1\n    eA = 3\nEnd Enum';

function source(declarations: string, ...body: string[]): string {
	return `Option Explicit\n${ENUM}\n${declarations}\nFunction Main() As Variant\n${body.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

function expectReport(src: string, code: string, span: string, message: string | readonly string[]): void {
	expectDiagnostic(src, byCode(analyzeModule(src), code), code, { span, message });
}

describe('an Enum member', () => {
	it.each([
		[['Dim x As Integer', 'x = eBig', 'Main = x'], 'eBig', 'stores 40000 in an Integer'],
		[['Dim x As Integer', 'x = E.eBig', 'Main = x'], 'E.eBig', 'stores 40000'],
		[['Dim x As Long', 'x = eMax + 1', 'Main = x'], 'eMax + 1', '2147483648, outside the Long range'],
		[['Dim x As Long', 'x = eK * eK', 'Main = x'], 'eK * eK', '10000000000'],
		[['Dim x As Byte', 'x = eNeg', 'Main = x'], 'eNeg', 'stores -1 in a Byte'],
	])('overflows: %j', (lines, span, message) => {
		expectReport(source('', ...lines), 'arithmetic-overflow', span, message);
	});

	it('counts on from the member before it', () => {
		const src = 'Option Explicit\nPrivate Enum F\n    fTop = 32767\n    fNext\nEnd Enum\nFunction Main() As Variant\n    Dim x As Integer\n    x = fNext\n    Main = x\nEnd Function\n';
		expectReport(src, 'arithmetic-overflow', 'fNext', 'stores 32768 in an Integer');
	});

	it('stays quiet in range, and under a local of its name', () => {
		expect(errors(source('', 'Dim x As Integer', 'x = eA', 'Main = x'))).toEqual([]);
		expect(errors(source('', 'Dim x As Long', 'x = eBig', 'Main = x'))).toEqual([]);
		expect(errors(source('', 'Dim eBig As Long', 'eBig = 1', 'Dim x As Integer', 'x = eBig', 'Main = x'))).toEqual([]);
	});

	it('gives a Const its value', () => {
		expectReport(source('Private Const K As Integer = eBig', 'Main = K'), 'const-overflow', 'eBig', 'value 40000 is outside that range');
		expect(errors(source('Private Const K As Integer = eA', 'Main = K'))).toEqual([]);
	});
});

describe('a String Const', () => {
	it.each([
		['Private Const K = "abc"', ['Dim x As Long', 'x = K', 'Main = x'], 'This string literal cannot be converted'],
		['Private Const K As String = "abc"', ['Dim x As Long', 'x = K', 'Main = x'], "constant 'K' (\"abc\")"],
		['', ['Const K = "abc"', 'Dim x As Long', 'x = K', 'Main = x'], "constant 'K'"],
		['Private Const K = "40000"', ['Dim x As Integer', 'x = K', 'Main = x'], 'outside the Integer range'],
	])('converts as its literal does: %s %j', (declaration, lines, message) => {
		expectReport(source(declaration, ...lines), 'assignment-type-mismatch', 'K', message);
	});

	it('is no number to add', () => {
		expectReport(source('Private Const K = "abc"', 'Main = K + 1'), 'string-arithmetic-coercion', 'K', "constant 'K', which is \"abc\"");
	});

	it('stays quiet where it converts, and in a concatenation', () => {
		expect(errors(source('Private Const K = "12"', 'Dim x As Long', 'x = K', 'Main = x'))).toEqual([]);
		expect(errors(source('Private Const K = "abc"', 'Main = K & 1'))).toEqual([]);
	});
});

describe('Const declarations the VBE refuses', () => {
	it.each(['Now', 'Date', 'Time', 'Timer', 'Rnd'])('a Const of %s', (fn) => {
		expectReport(source(`Private Const K = ${fn}`, 'Main = K'), 'const-value-not-constant', fn, 'a VBA function evaluated as the code runs');
	});

	it('a local Const of Now, and not a VBA constant or a module Const named Now', () => {
		expectReport(source('', 'Const K = Now', 'Main = K'), 'const-value-not-constant', 'Now', 'not constant');
		expect(errors(source('Private Const K = vbCrLf', 'Main = Len(K)'))).toEqual([]);
		expect(errors(source('Private Const Now = 5\nPrivate Const K = Now', 'Main = K'))).toEqual([]);
	});

	it('a Date Const past 12/31/9999', () => {
		expectReport(source('Private Const D1 As Date = #12/31/9999#\nPrivate Const D2 = D1 + 1', 'Main = D2'), 'const-overflow', 'D1 + 1', 'outside the Date range');
		expect(errors(source('Private Const D1 As Date = #12/30/9999#\nPrivate Const D2 = D1 + 1', 'Main = D2'))).toEqual([]);
	});

	it.each([
		['Private Const K = 1', ['Set K = Nothing', 'Main = 1'], 'K'],
		['', ['Set eA = Nothing', 'Main = 1'], 'eA'],
		['', ['Const K = 1', 'Set K = Nothing', 'Main = 1'], 'K'],
	])('Set of a constant: %s %j', (declaration, lines, span) => {
		expectReport(source(declaration, ...lines), 'const-assignment', span, 'Cannot assign to constant');
	});

	it.each([
		['Private Const K = 1', ['ReDim K(3)', 'Main = 1'], 'K'],
		['', ['ReDim eA(3)', 'Main = 1'], 'eA'],
		['', ['Const K = 1', 'ReDim K(3)', 'Main = 1'], 'K'],
	])('ReDim of a constant: %s %j', (declaration, lines, span) => {
		expectReport(source(declaration, ...lines), 'scalar-redim', span, 'Expected array');
	});
});

describe('an expression that opens with a date', () => {
	it('lexes the date, not a directive', () => {
		const toks = rawExpressionTokens('#12/31/9999# + 1');
		expect(toks.map((tok) => [tok.kind, tok.rawText, tok.start])).toEqual([
			['dateLiteral', '#12/31/9999#', 0],
			['operator', '+', 13],
			['integerLiteral', '1', 15],
		]);
	});
});
