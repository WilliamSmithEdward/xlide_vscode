// Diagnostics tests: overflow the code proves (issue #116). Each raising
// sample was measured in Excel 16.0 (build 20326, 2026-09-26): it compiles
// and raises error 6 every time it runs, or (for a Const) is refused while
// compiling. Each quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const ARITHMETIC = 'arithmetic-overflow';
const CONST = 'const-overflow';
const COUNTER = 'for-counter-overflow';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('arithmetic-overflow - literal and constant arithmetic (issue #116)', () => {
	it('types two Integer literals as Integer arithmetic even when a Long receives the result', () => {
		const src = wrap('Dim secs As Long', 'secs = 60 * 60 * 24', 'Main = secs');
		expectDiagnostic(src, analyzeModule(src), ARITHMETIC, { span: '60 * 60 * 24', message: ['3600 (Integer) * 24 (Integer)', 'Integer range', "error '6'"] });
	});

	it('flags each measured literal expression', () => {
		const cases: Array<[string, string]> = [
			['32767 + 1', 'Integer range'],
			['1024 * 1024', 'Integer range'],
			['255 * 255', 'Integer range'],
			['24 * 60 * 60', 'Integer range'],
			['&H8000 - 1', '-32768 (Integer) - 1'],
			['2147483647 + 1', 'Long range'],
			['50000 * 50000', 'Long range'],
			['10 ^ 309', 'Double range'],
			['Exp(1000)', 'Double range'],
		];
		for (const [expr, message] of cases) {
			const src = wrap(`Main = ${expr}`);
			expectDiagnostic(src, analyzeModule(src), ARITHMETIC, { span: expr, message });
		}
		// The Long suffix arrives too late: the Integer product before it overflows first.
		const late = wrap('Main = 24 * 60 * 60 * 1000&');
		expectDiagnostic(late, analyzeModule(late), ARITHMETIC, { span: '24 * 60 * 60', message: 'Integer range' });
	});

	it('folds Integer Consts with Integer typing', () => {
		const src =
			'Option Explicit\nPrivate Const HOURS As Integer = 24\nPrivate Const MINUTES As Integer = 60\n' +
			'Function Main() As Variant\n    Main = HOURS * MINUTES * MINUTES\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), ARITHMETIC, { span: 'HOURS * MINUTES * MINUTES', message: '1440 (Integer) * 60 (Integer)' });
	});

	it('stays quiet at the boundaries and when one operand is a Long', () => {
		const src = wrap('Dim l As Long', 'l = 32767 + 1&', 'Main = 181 * 181', 'Main = 32767 + 0', 'Main = -32768 - 0', 'Main = l + 1');
		expect(byCode(analyzeModule(src), ARITHMETIC)).toHaveLength(0);
	});
});

describe('const-overflow (issue #116)', () => {
	it('reports a module Const whose Integer arithmetic overflows as a compile error', () => {
		const src = 'Option Explicit\nPrivate Const SECONDS_PER_DAY = 60 * 60 * 24\nFunction Main() As Variant\n    Main = SECONDS_PER_DAY\nEnd Function\n';
		expectDiagnostic(src, analyzeModule(src), CONST, { span: '60 * 60 * 24', message: ["'SECONDS_PER_DAY'", 'compile error'] });
	});

	it('reports a Const declared As a type its value does not fit, and a local Const', () => {
		const src = 'Option Explicit\nPrivate Const BIG As Integer = 40000\nFunction Main() As Variant\n    Const K = 50000 * 50000\n    Main = BIG + K\nEnd Function\n';
		expectDiagnostics(src, analyzeModule(src), CONST, [
			{ span: '40000', message: 'As Integer' },
			{ span: '50000 * 50000', message: 'Long range' },
		]);
	});
});

describe('for-counter-overflow (issue #116)', () => {
	it('flags a limit the counter type cannot step past', () => {
		const cases: Array<[string, string, string]> = [
			['Dim i As Integer', 'For i = 1 To 32767', '32767'],
			['Dim i As Integer', 'For i = 1 To 40000', '40000'],
			['Dim b As Byte', 'For b = 0 To 255', '255'],
			['Dim i As Integer', 'For i = 100 To -32768 Step -1', '-32768'],
			['Dim l As Long', 'For l = 2147483000 To 2147483647', '2147483647'],
		];
		for (const [decl, header, span] of cases) {
			const src = wrap(decl, header, 'Next', 'Main = 0');
			expectDiagnostic(src, analyzeModule(src), COUNTER, { span, message: ['after its last pass', "error '6'"] });
		}
	});

	it('stays quiet one below the maximum, for a Long counter, and when the step lands short of it', () => {
		const src = wrap(
			'Dim i As Integer, l As Long',
			'For i = 1 To 32766', 'Next',
			'For l = 1 To 32767', 'Next',
			'For i = 100 To -32767 Step -1', 'Next',
			'For i = 1 To 32766 Step 2', 'Next',
			'Main = 0',
		);
		expect(byCode(analyzeModule(src), COUNTER)).toHaveLength(0);
	});
});

describe('arithmetic-overflow - assignments and conversions (issue #116)', () => {
	it('flags a stored value the target type cannot hold, after rounding', () => {
		const cases: Array<[string, string, string]> = [
			['Dim b As Byte', 'b = 200 + 100', 'stores 300 in a Byte'],
			['Dim b As Byte', 'b = 255.5', 'rounds to 256'],
			['Dim i As Integer', 'i = 32767.5', 'rounds to 32768'],
			['Dim d As Date', 'd = 3000000', 'Date'],
			['Dim i As Integer', 'i = Rows.Count', '1048576'],
		];
		for (const [decl, assignment, message] of cases) {
			const src = wrap(decl, assignment, 'Main = 0');
			expectDiagnostic(src, analyzeModule(src), ARITHMETIC, { message });
		}
		const rounding = wrap('Dim b As Byte', 'b = -0.5', 'b = 255.4', 'Main = b');
		expect(byCode(analyzeModule(rounding), ARITHMETIC)).toHaveLength(0);
	});

	it('follows a value through a straight run of statements', () => {
		const src = wrap('Dim i As Integer', 'i = 32767', 'i = i + 1', 'Main = i');
		expectDiagnostic(src, analyzeModule(src), ARITHMETIC, { span: 'i + 1', message: '32767 (Integer) + 1 (Integer)' });
		const branched = wrap('Dim i As Integer', 'i = 32767', 'If Main Then i = 0', 'i = i + 1', 'Main = i');
		expect(byCode(analyzeModule(branched), ARITHMETIC)).toHaveLength(0);
		const negated = wrap('Dim i As Integer', 'i = -32768', 'Main = -i');
		expectDiagnostic(src.replace(src, negated), analyzeModule(negated), ARITHMETIC, { span: '-i', message: 'Negating -32768' });
	});

	it('flags conversions whose argument is outside the target range', () => {
		const cases: Array<[string, string]> = [
			['CInt(40000)', 'CInt(40000) does not fit Integer'],
			['CByte(-1)', 'CByte(-1) does not fit Byte'],
			['CLng(2147483647.5)', 'does not fit Long'],
			['CSng(1E+39)', 'does not fit Single'],
			['Hex(1E+20)', 'outside the Long range'],
			['Abs(CInt(-32768))', 'Abs(-32768) does not fit Integer'],
		];
		for (const [expr, message] of cases) {
			const src = wrap(`Main = ${expr}`);
			expectDiagnostic(src, analyzeModule(src), ARITHMETIC, { span: expr, message });
		}
		const quiet = wrap('Main = CInt(32767.4)', 'Main = CByte(255)', 'Main = CLng(-2147483648#)', 'Main = Abs(CInt(-32767))');
		expect(byCode(analyzeModule(quiet), ARITHMETIC)).toHaveLength(0);
	});

	it('also sees a conversion inside another statement', () => {
		const src = wrap('Debug.Print CInt(40000)');
		expectDiagnostic(src, analyzeModule(src), ARITHMETIC, { span: 'CInt(40000)' });
	});
});
