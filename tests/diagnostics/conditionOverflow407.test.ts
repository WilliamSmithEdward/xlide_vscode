// Diagnostics tests: an If, ElseIf, While or Do condition is evaluated whole,
// and overflows as an assignment of it would (issue #407). Each case was run
// through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

const LARGE = 'Dim d As Double\n    d = 1E10\n    ';

describe('a condition past the Long range under a logical operator (issue #407)', () => {
	it('overflows in each kind of condition', () => {
		expect(errors(`${LARGE}If d And 1 Then Main = 2`)).toEqual(['arithmetic-overflow']);
		expect(errors(`${LARGE}If d And 1 Then\n        Main = 2\n    End If`)).toEqual(['arithmetic-overflow']);
		expect(errors(`${LARGE}If d < 0 Then\n        Main = 1\n    ElseIf d And 1 Then\n        Main = 2\n    End If`)).toEqual(['arithmetic-overflow']);
		expect(errors(`${LARGE}Do While d Xor 1\n        Exit Do\n    Loop`)).toEqual(['arithmetic-overflow']);
		expect(errors('If 1E10 And 1 Then Main = 2')).toEqual(['arithmetic-overflow']);
		expect(errors('If CInt(40000) > 1 Then\n        Main = 2\n    End If')).toEqual(['arithmetic-overflow']);
	});

	it('stays quiet where the values fit or the operands are comparisons', () => {
		expect(errors('Dim d As Double\n    d = 100\n    If d And 1 Then Main = 2')).toEqual([]);
		expect(errors(`${LARGE}If d > 1 And d < 1E11 Then Main = 2`)).toEqual([]);
	});
});
