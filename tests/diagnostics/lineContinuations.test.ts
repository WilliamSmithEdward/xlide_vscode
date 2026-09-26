// Diagnostics tests: where the VBE refuses a line continuation (issue #126).
// Measured in Excel 16.0 (build 20326, 2026-09-25); each accepted sample
// compiles there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'invalid-line-continuation';

function sum(terms: number): string {
	const lines = Array.from({ length: terms - 1 }, () => '        1 + _');
	return `Option Explicit\nFunction Main() As Variant\n    Dim x As Long\n    x = 1 + _\n${lines.join('\n')}\n        1\n    Main = x\nEnd Function\n`;
}

describe('invalid-line-continuation - limits (issue #126)', () => {
	it('allows 24 continuations in one logical line and refuses the 25th', () => {
		expect(byCode(analyzeModule(sum(24)), CODE)).toHaveLength(0);
		const src = sum(25);
		expectDiagnostic(src, analyzeModule(src), CODE, { message: '25th line continuation' });
	});

	it('refuses a continuation followed by an empty line', () => {
		const src = 'Option Explicit\nFunction Main() As Variant\n    Dim x As Long\n    x = 1 + _\n\n    Main = x\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { message: 'next line is empty' });
	});

	it('refuses a continuation inside an Enum body but not in its header or in a Type', () => {
		for (const body of ['    kOne _\n    = 1\nEnd Enum', '    kOne = _\n    1\nEnd Enum', '    kOne = 1\nEnd _\nEnum']) {
			const src = `Option Explicit\nPrivate Enum K\n${body}\nFunction Main() As Variant\n    Main = kOne\nEnd Function\n`;
			expectDiagnostic(src, analyzeModule(src), CODE, { message: 'Invalid inside Enum' });
		}
		const quiet = 'Option Explicit\nPrivate Type T\n    a As Long\n    b _\n    As Long\nEnd Type\nPrivate _\nEnum K\n    kOne = 1\nEnd Enum\nFunction Main() As Variant\n    Main = kOne\nEnd Function\n';
		expect(byCode(analyzeModule(quiet), CODE)).toHaveLength(0);
	});

	it('refuses a Declare split between the Lib string and the parameter list', () => {
		const src = 'Option Explicit\nPrivate Declare PtrSafe Function GetTickCount Lib "kernel32" _\n    () As Long\nFunction Main() As Variant\n    Main = 1\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), CODE, { message: 'Lib or Alias string' });
		const quiet = 'Option Explicit\nPrivate Declare PtrSafe _\nFunction GetTickCount _\nLib "kernel32" () As Long\nPrivate _\nDeclare PtrSafe Function GetTickCount2 Lib "kernel32" Alias "GetTickCount" ( _\n) As Long\nFunction Main() As Variant\n    Main = 1\nEnd Function\n';
		expect(byCode(analyzeModule(quiet), CODE)).toHaveLength(0);
	});
});
