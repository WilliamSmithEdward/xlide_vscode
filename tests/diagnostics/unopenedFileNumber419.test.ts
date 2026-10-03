// Diagnostics tests: a file number no Open in the project names raises 52
// (issue #419). Each case was run through pyVBAharness on 2026-10-02 in Excel
// 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { openedFileNumbersIn, type OpenedFileNumbers } from '../../src/analyzer/diagnostics/openedFileNumbers';

const NONE: OpenedFileNumbers = { any: false, numbers: new Set() };

function errors(body: string, opened: OpenedFileNumbers | 'unknown' = NONE): string[] {
	const projectOpenedFileNumbers = opened === 'unknown' ? undefined : opened;
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src, { projectOpenedFileNumbers }).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a file number nothing opens (issue #419)', () => {
	it('raises 52 in every file statement and function but Close', () => {
		for (const body of ['Print #1, "x"', 'Write #1, "x"', 'Dim s As String\n    Input #1, s', 'Dim s As String\n    Line Input #1, s',
			'Dim n As Long\n    Get #1, , n', 'Dim n As Long\n    Put #1, , n', 'Seek #1, 1', 'Width #1, 10', 'Lock #1', 'Unlock #1',
			'Main = EOF(1)', 'Main = LOF(1)', 'Main = Loc(1)', 'Main = FileAttr(1, 1)', 'Main = Seek(1)', 'Main = Input(1, #1)']) {
			expect(errors(body), body).toEqual(['file-number-zero']);
		}
		expect(errors('Close #1')).toEqual([]);
	});

	it('is not judged where an Open may reach the number', () => {
		expect(errors('Print #1, "x"', { any: false, numbers: new Set([1]) })).toEqual([]);
		expect(errors('Print #1, "x"', { any: true, numbers: new Set() })).toEqual([]);
		expect(errors('Print #1, "x"', 'unknown')).toEqual([]);
		expect(errors('Open Environ("TEMP") & "\\a.txt" For Output As #1\n    Print #1, "x"\n    Close #1')).toEqual([]);
	});

	it('reads the numbers Open names', () => {
		const opened = openedFileNumbersIn('Sub A()\n    Open "C:\\x.txt" For Input As #3\n    If x Then Open p For Output As 4: Open q For Append As #f\nEnd Sub\n');
		expect([...opened.numbers].sort()).toEqual([3, 4]);
		expect(opened.any).toBe(true);
		expect(openedFileNumbersIn('Sub A()\n    \' Open p For Input As #f\n    Name p As q\nEnd Sub\n')).toEqual({ any: false, numbers: new Set() });
	});
});
