// Diagnostics tests: Declare statements (issue #254). Every case was
// measured in Excel 16.0 (build 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const SLEEP = 'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)';
const TICK = 'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long';

function source(declarations: readonly string[], ...body: string[]): string {
	return `Option Explicit\n${declarations.join('\n')}\nFunction Main() As Variant\n${body.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

function expectReport(src: string, code: string, span: string, message: string | readonly string[]): void {
	expectDiagnostic(src, byCode(analyzeModule(src), code), code, { span, message });
}

describe('a Declare in the procedure-name rules', () => {
	it('is a Sub that returns nothing', () => {
		expectReport(source([SLEEP], 'Main = Sleep(1)'), 'sub-used-as-value', 'Sleep', 'Expected Function or variable');
		expect(errors(source([SLEEP], 'Sleep 1', 'Main = 1'))).toEqual([]);
	});

	it('is a Function no assignment can reach', () => {
		expectReport(source([TICK], 'GetTickCount = 5', 'Main = 1'), 'assignment-to-procedure-name', 'GetTickCount', 'Function call on left-hand side');
		expect(errors(source([TICK], 'Main = (GetTickCount() > 0)'))).toEqual([]);
		expect(errors(source([TICK], 'Dim GetTickCount As Long', 'GetTickCount = 1', 'Main = GetTickCount'))).toEqual([]);
	});

	it.each([
		['Private Function GetTickCount() As Long\nEnd Function', 'is already declared'],
		[TICK, 'is already declared'],
		['Private GetTickCount As Long', 'names both a variable and a procedure'],
		['Private Const GetTickCount As Long = 1', 'names both a variable and a procedure'],
	])('shares its name with nothing: %s', (other, message) => {
		expectReport(source([TICK, other], 'Main = 1'), 'duplicate-procedure', 'GetTickCount', ['Ambiguous name detected', message]);
	});
});

describe("a Declare's own line", () => {
	it.each([
		['Private Declare PtrSafe Sub S Lib "kernel32" Alias "Sleep" (ByVal ms As Long) As Long', 'As', 'Expected: end of statement'],
		['Private Declare PtrSafe Function F Lib "kernel32" Alias "GetTickCount" () As String * 4', '*', 'Expected: end of statement'],
		['Private Declare PtrSafe Function F& Lib "kernel32" Alias "GetTickCount" () As Long', 'As', 'Expected: end of statement'],
		['Private Declare PtrSafe Function F Lib "kernel32" Alias "GetTickCount" () As Any', 'Any', 'Expected: type name'],
		['Private Declare PtrSafe Function F Lib K Alias "GetTickCount" () As Long', 'K', 'Expected: string constant'],
		['Private Declare PtrSafe Function F Lib "kernel32" Alias A () As Long', 'A', 'Expected: string constant'],
	])('%s', (declare, span, message) => {
		const consts = 'Private Const K As String = "kernel32"\nPrivate Const A As String = "GetTickCount"';
		expectReport(source([consts, declare], 'Main = 1'), 'invalid-proc-header', span, message);
	});

	it.each([
		'Private Declare PtrSafe Sub F Lib "kernel32" Alias "Sleep" (ByVal s As String * 4)',
		'Private Declare PtrSafe Sub F Lib "kernel32" Alias "Sleep" (ByRef s As String * 4)',
		'Private Declare PtrSafe Sub F Lib "kernel32" Alias "Sleep" (s As String * 4)',
		'Private Sub F(ByVal s As String * 4)\nEnd Sub',
	])('takes no fixed-length String parameter: %s', (declaration) => {
		expectReport(source([declaration], 'Main = 1'), 'invalid-proc-header', 'String * 4', 'Expected array');
	});

	it('stays quiet on what compiles', () => {
		for (const declare of [
			'Private Declare PtrSafe Function F& Lib "kernel32" Alias "GetTickCount" ()',
			'Private Declare PtrSafe Sub F Lib "kernel32" Alias "Sleep" (ByVal s As Any)',
			'Private Declare PtrSafe Function F Lib "kernel32" Alias "GetTickCount" () As Object',
		]) {
			expect(errors(source([declare], 'Main = 1')), declare).not.toContain('invalid-proc-header');
		}
	});
});

describe('a Declare that fails on every call', () => {
	it.each([
		['Private Declare PtrSafe Function F CDecl Lib "kernel32" Alias "GetTickCount" () As Long', 'Main = F()', "'49': Bad DLL calling convention"],
		['Private Declare PtrSafe Sub F CDecl Lib "kernel32" Alias "Sleep" (ByVal ms As Long)', 'F 1', "'49'"],
		['Private Declare PtrSafe Function F Lib "" Alias "GetTickCount" () As Long', 'Main = F()', "'48': File not found"],
		['Private Declare PtrSafe Function F Lib " " Alias "GetTickCount" () As Long', 'Main = F()', "'48'"],
		['Private Declare PtrSafe Sub F Lib "" Alias "Sleep" (ByVal ms As Long)', 'Call F(1)', "'48'"],
		['Private Declare PtrSafe Function F Lib "" Alias "GetTickCount" () As Long', 'Main = F', "'48'"],
		['Private Declare PtrSafe Function F Lib "kernel32" Alias "" () As Long', 'Main = F()', "'453': Can't find DLL entry point"],
		['Private Declare PtrSafe Function F Lib "kernel32" Alias " " () As Long', 'Main = F()', "'453'"],
		['Private Declare PtrSafe Function F Lib "kernel32" Alias "#0" () As Long', 'Main = F()', "'452': Can't find DLL entry point 0"],
		['Private Declare PtrSafe Function F Lib "kernel32" Alias "#00" () As Long', 'Main = F()', "'452'"],
	])('is reported at the call: %s', (declare, call, message) => {
		expectReport(source([declare], call, 'Main = 1'), 'unusable-declare', 'F', message);
	});

	it('stays quiet while nothing calls it, under a local of its name, and on a real Lib and Alias', () => {
		expect(errors(source(['Private Declare PtrSafe Function F Lib "" Alias "GetTickCount" () As Long'], 'Dim F As Long', 'F = 2', 'Main = F'))).toEqual([]);
		expect(errors(source(['Private Declare PtrSafe Function F CDecl Lib "kernel32" Alias "GetTickCount" () As Long'], 'Main = 1'))).toEqual([]);
		expect(errors(source(['Private Declare PtrSafe Function F Lib "" Alias "GetTickCount" () As Long'], 'Main = 1'))).toEqual([]);
		expect(errors(source(['Private Declare PtrSafe Function F Lib "kernel32" Alias "GetTickCount" () As Long'], 'Main = (F() > 0)'))).toEqual([]);
		// A Lib with a space before a real name is a missing file, 53, which depends on the machine.
		expect(errors(source(['Private Declare PtrSafe Function F Lib " kernel32" Alias "GetTickCount" () As Long'], 'Main = (F() > 0)'))).toEqual([]);
	});
});

describe('Nothing into a ByVal Long', () => {
	it('is reported in a single-line If branch too', () => {
		const src = `Option Explicit\nPrivate Sub Sl(ByVal ms As Long)\nEnd Sub\nFunction Main() As Variant\n    Main = 1\n    If False Then Sl Nothing\nEnd Function\n`;
		expectReport(src, 'argument-object-type-mismatch', 'Nothing', 'Invalid use of object');
	});
});
