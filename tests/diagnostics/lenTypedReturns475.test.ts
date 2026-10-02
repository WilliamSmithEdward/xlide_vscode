// Diagnostics tests: Len of an expression on a VBA function with a fixed
// return type (issue #475). Each sample was compiled through pyVBAharness
// on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const REQUIRED = 'variable-required';

function wrap(expr: string): string {
	return `Option Explicit\nFunction Main() As Variant\n    Dim i As Integer, s As String, d As Double, v As Variant\n    Main = Len(${expr}) + i + Len(s) + d + Len(v)\nEnd Function\n`;
}

describe('Len of a typed VBA function (issue #475)', () => {
	it('reports Timer, Err.Number, Asc, Len, LenB, InStr and a $ function in an expression', () => {
		const refused = [
			'Timer < 1', 'Timer + 1', 'Err.Number = 1', 'Asc("a") / 2', 'Asc("a") + 1', 'LenB(s) = 1', 'Len(s) + 1',
			'Format$(1) < i', 'Mid$(s, 1, 1) < i', 'InStr(s, "a") + 1', 'Timer', 'Err.Number + 1',
			'Asc("a")', 'AscW("a")', 'Len(s)', 'InStrRev(s, "a")', 'Err.Number',
		];
		for (const expr of refused) {
			expect(byCode(analyzeModule(wrap(expr)), REQUIRED), expr).toHaveLength(1);
		}
	});

	it('stays quiet where the function returns a Variant or a String', () => {
		const compiled = ['Format(1) < i', 'Mid(s, 1, 1) < i', 'Now', 'Left$(s, 1)', 'Mid$(s, 1, 1)', 'Err.Description', 'v Like "a"', 'Format$(1)', 'Chr$(65)'];
		for (const expr of compiled) {
			expect(byCode(analyzeModule(wrap(expr)), REQUIRED), expr).toHaveLength(0);
		}
	});

	it('follows a procedure of the module that hides the library name', () => {
		const src = 'Option Explicit\nFunction Asc(ByVal s As String) As Variant\n    Asc = 1\nEnd Function\nFunction Main() As Variant\n    Main = Len(Asc("a"))\nEnd Function\n';
		expect(byCode(analyzeModule(src), REQUIRED)).toHaveLength(0);
	});
});
