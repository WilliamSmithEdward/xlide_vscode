import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../src/vbaModuleAnalysis';

const MISMATCH = 'assignment-object-type-mismatch';

function codes(src: string): string[] {
	return analyzeVbaModuleSource({ source: src, moduleName: 'Module1' }).diagnostics.map((d) => d.code);
}

describe('assignment-object-type-mismatch: indexed collection element', () => {
	it('does not flag ThisWorkbook.Sheets("x") assigned to a Worksheet', () => {
		// Sheets(i) is a late-bound single sheet (Worksheet OR Chart); VBA allows
		// `Set ws = Sheets("x")`. Regression for the returnsAnyOf Item resolver.
		const src = 'Sub S()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Sheets("mySheet")\nEnd Sub\n';
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('does not flag ThisWorkbook.Worksheets("x") assigned to a Worksheet', () => {
		const src = 'Sub S()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Worksheets("mySheet")\nEnd Sub\n';
		expect(codes(src)).not.toContain(MISMATCH);
	});

	it('still flags a provable mismatch (Range assigned to a Worksheet)', () => {
		const src = 'Sub S()\n    Dim ws As Worksheet\n    Dim r As Range\n    Set ws = r\nEnd Sub\n';
		expect(codes(src)).toContain(MISMATCH);
	});
});
