// Diagnostics tests: ReDim Preserve after Erase sets the bounds afresh, the
// first dimension included (issue #420). Measured in Excel 16.0 (build
// 20326, 2026-10-01).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic } from '../helpers/diagnostics';

const CODE = 'redim-preserve-dimension-change';

function source(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim a() As Long, b() As Long\n${lines.map((line) => `    ${line}`).join('\n')}\n    Main = UBound(a, 1)\nEnd Function\n`;
}

describe('ReDim Preserve after Erase', () => {
	it('runs with new bounds in every dimension', () => {
		for (const lines of [
			['ReDim a(1 To 2, 1 To 3)', 'Erase a', 'ReDim Preserve a(1 To 3, 1 To 3)'],
			['ReDim a(1 To 2, 1 To 3)', 'Erase b, a', 'ReDim Preserve a(1 To 3)'],
		]) {
			expect(byCode(analyzeModule(source(...lines)), CODE), lines.join(' / ')).toEqual([]);
		}
	});

	it('is still reported without the Erase, or after erasing another array', () => {
		for (const lines of [
			['ReDim a(1 To 2, 1 To 3)', 'ReDim Preserve a(1 To 3, 1 To 3)'],
			['ReDim a(1 To 2, 1 To 3)', 'Erase b', 'ReDim Preserve a(1 To 3, 1 To 3)'],
		]) {
			const src = source(...lines);
			expectDiagnostic(src, byCode(analyzeModule(src), CODE), CODE, {});
		}
	});
});
