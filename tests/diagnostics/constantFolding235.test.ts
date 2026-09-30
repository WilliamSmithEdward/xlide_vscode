// Diagnostics tests: constant folding as the VBE folds (issue #235). Measured
// through pyVBAharness in 64-bit Excel 16.0 on 2026-09-30, each Const read
// back with TypeName and CStr.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const MAIN = 'Function Main() As Variant\n    Main = C\nEnd Function\n';

function withConst(decl: string): string {
	return `Option Explicit\nPrivate Const C ${decl}\n${MAIN}`;
}

function main(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

function errors(src: string): string[] {
	return analyzeModule(src).filter((d) => d.severity === 'error').map((d) => d.code ?? '');
}

describe('negating the Long minimum wraps where the VBE folds it (issue #235)', () => {
	it.each([
		'As Long = -&H80000000',
		'= -&H80000000',
		'As Double = -&H80000000',
		'As Long = -(-2147483647 - 1)',
	])('Const C %s compiles', (decl) => {
		expect(errors(withConst(decl))).toEqual([]);
	});

	it.each([
		['Dim l As Long', 'l = -&H80000000', 'Main = l'],
		['Main = -&H80000000&'],
		['Main = Abs(&H80000000)'],
	])('%s runs', (...lines) => {
		expect(errors(main(...lines))).toEqual([]);
	});

	it('still reports it at run time, and the Integer minimum anywhere', () => {
		expect(byCode(analyzeModule(main('Dim l As Long', 'l = &H80000000', 'l = -l', 'Main = l')), 'arithmetic-overflow')).toHaveLength(1);
		expect(byCode(analyzeModule(main('Main = 0 - &H80000000')), 'arithmetic-overflow')).toHaveLength(1);
		expect(byCode(analyzeModule(main('Main = -1 * &H80000000')), 'arithmetic-overflow')).toHaveLength(1);
		expect(byCode(analyzeModule(withConst('As Integer = -&H8000')), 'const-overflow')).toHaveLength(1);
		// A conversion to another type is no constant: CLng runs when the line does.
		expect(byCode(analyzeModule(main('Main = -CLng(-2147483648#)')), 'arithmetic-overflow')).toHaveLength(1);
		expect(byCode(analyzeModule(main('Dim l As Long', 'l = &H80000000', 'Main = -(l + 0)')), 'arithmetic-overflow')).toHaveLength(1);
	});

	it('wraps a Const holding the minimum where a line negates it', () => {
		expect(errors(`Option Explicit\nPrivate Const K = &H80000000\nFunction Main() As Variant\n    Main = -K\nEnd Function\n`)).toEqual([]);
	});

	it('folds away a conversion to the type the constant has, and constant arithmetic', () => {
		expect(errors(main('Main = -CLng(&H80000000)'))).toEqual([]);
		expect(errors(main('Main = -(&H80000000 + 0)'))).toEqual([]);
	});
});

describe('const-evaluation-error: Division by zero (issue #235)', () => {
	it.each(['= 1 / 0', '= 1 \\ 0', '= 1 Mod 0', '= 1 / False', '= 1 / 0&', '= &H7 / 0', 'As Long = 1 \\ 0', '= 1 \\ 0.4', '= 2 * (1 \\ 0)'])('Const C %s', (decl) => {
		const hits = byCode(analyzeModule(withConst(decl)), 'const-evaluation-error');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain('This is a VBE compile error: Division by zero.');
	});

	it('leaves a division by a Const that is not zero, and a runtime division to division-by-zero', () => {
		expect(errors(`Option Explicit\nPrivate Const D = 2\nPrivate Const C = 1 / D\n${MAIN}`)).toEqual([]);
		expect(byCode(analyzeModule(main('Main = 1 / 0')), 'const-evaluation-error')).toEqual([]);
	});
});

describe('a string into a typed Const (issue #235)', () => {
	it.each([
		['As Long = "abc"', 'const-evaluation-error', 'Type mismatch'],
		['As Long = ""', 'const-evaluation-error', 'Type mismatch'],
		['As Double = "abc"', 'const-evaluation-error', 'Type mismatch'],
		['As Boolean = "abc"', 'const-evaluation-error', 'Type mismatch'],
		['= -"abc"', 'const-evaluation-error', 'Type mismatch'],
		['As String = Not ""', 'const-evaluation-error', 'Type mismatch'],
		['As Integer = "40000"', 'const-overflow', 'Overflow'],
	])('Const C %s', (decl, code, vbe) => {
		const hits = byCode(analyzeModule(withConst(decl)), code);
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain(`This is a VBE compile error: ${vbe}.`);
	});

	it.each(['As Long = "12"', 'As Boolean = "True"', '= -"12"', 'As String = "abc"', 'As Variant = "abc"'])('Const C %s compiles', (decl) => {
		expect(errors(withConst(decl))).toEqual([]);
	});
});

describe('Not into a narrow type (issue #235)', () => {
	it.each([
		['As Byte = Not 0', '-1'],
		['As Byte = Not 255', '-256'],
		['As Integer = Not 32768!', '-32769'],
		['As Byte = Not 255 + 256', '-512'],
	])('Const C %s', (decl, value) => {
		const hits = byCode(analyzeModule(withConst(decl)), 'const-overflow');
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain(`its value ${value} is outside that range`);
	});

	it('makes Not of a Single a Long, which then overflows as a Long', () => {
		const src = `Option Explicit\nPrivate Const A = Not 1!\nPrivate Const B = A * 2000000000\n${MAIN}`;
		expect(byCode(analyzeModule(src), 'const-overflow').map((d) => d.message)).toEqual([expect.stringContaining('outside the Long range')]);
	});

	it('leaves Not 32767 into an Integer, which is -32768', () => {
		expect(errors(withConst('As Integer = Not 32767'))).toEqual([]);
	});
});
