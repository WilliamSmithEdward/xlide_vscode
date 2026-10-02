// Diagnostics tests: a LongLong or LongPtr narrowed implicitly in 64-bit
// VBA (issue #298). Measured in 64-bit Excel 16.0 (2026-10-02) through
// pyVBAharness: each is a compile error, "Type mismatch".

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

const D = 'Dim q As LongLong, p As LongPtr, n As Long, i As Integer, s As String, x As Long\n    q = 3\n    p = 3\n    ';

function found(body: string, extra = '', win64 = 1): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${D}${body}\nEnd Function\nPrivate Sub TakeByValLong(ByVal v As Long)\nEnd Sub\n${extra}`;
	return analyzeModule(src, { conditionalCompilation: { projectConstants: { Win64: win64 } } } as never)
		.filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a LongLong where a narrower whole number goes', () => {
	it('is a Type mismatch stored, passed ByVal, indexed or bounded', () => {
		for (const body of ['n = q', 'i = p', 'n = q + 0', 'n = 3^', 'n = VarPtr(x)', 'n = StrPtr(s)', 'TakeByValLong q',
			'Dim a(5) As Long\n    Main = a(q)', 'Dim a() As Long\n    ReDim a(q)', 'For x = 1 To q\n    Next',
			'Main = Mid$("abc", q, 1)', 'Main = Space$(q)', 'Main = InStr(q, "abc", "b")']) {
			expect(found(body), body).toEqual(['longlong-narrowing']);
		}
	});

	it('compiles converted, into a wide type, or as a Double', () => {
		for (const body of ['Dim d As Double\n    d = q', 'Dim v As Variant\n    v = q', 's = q', 'n = CLng(q)', 'q = n', 'p = n', 'Main = (q = n)', 'n = q / 1']) {
			expect(found(body), body).toEqual([]);
		}
	});

	it('is left alone where Win64 is off, as in 32-bit Office', () => {
		expect(found('n = p', '', 0)).toEqual([]);
	});
});
