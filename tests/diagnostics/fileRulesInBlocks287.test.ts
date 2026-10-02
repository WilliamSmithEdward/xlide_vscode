// Diagnostics tests: the file rules and a self-increment's overflow inside
// a block that always runs (issue #287). Measured in Excel 16.0 64-bit
// (2026-10-02) through pyVBAharness inside If True, For k = 1 To 1 and
// With Application; each raises as it does at the top level.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const BLOCKS: ReadonlyArray<readonly [string, string]> = [['If True Then', 'End If'], ['For k = 1 To 1', 'Next'], ['With Application', 'End With']];

function found(declaration: string, body: string, open: string, close: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim k As Long\n    ${declaration}${open}\n        ${body}\n    ${close}\n    Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('inside a block that always runs', () => {
	it.each(BLOCKS)('reports file number zero inside %s', (open, close) => {
		expect(found('', 'Print #0, "x"', open, close)).toEqual(['file-number-zero']);
	});

	it.each(BLOCKS)('reports a file used after Close inside %s', (open, close) => {
		expect(found('Dim f As String, s As String\n    f = "x.txt"\n    ', 'Open f For Output As #1\n        Close #1\n        Print #1, s', open, close)).toContain('file-used-after-close');
	});

	it.each(BLOCKS)('reports an Integer counted past its range inside %s', (open, close) => {
		expect(found('Dim i As Integer\n    ', 'i = 32767\n        i = i + 1', open, close)).toContain('arithmetic-overflow');
	});
});
