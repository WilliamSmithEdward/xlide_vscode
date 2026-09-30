// Diagnostics tests: arithmetic-overflow inside call arguments and operands,
// and LongLong and LongPtr bounds (issues #232 and #258). Measured through
// pyVBAharness in 64-bit Excel 16.0 on 2026-09-30: each raising sample
// raises 6, and each quiet one runs.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode } from '../helpers/diagnostics';

const CODE = 'arithmetic-overflow';

function main(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('arithmetic-overflow inside arguments and operands (issues #232, #258)', () => {
	it.each([
		['Main = CStr(CInt(40000))', 'CInt(40000)'],
		['Main = CStr(CLng(3000000000#))', 'CLng(3000000000#)'],
		['Main = CStr(CByte(256))', 'CByte(256)'],
		['Main = CStr(CCur(1E+16))', 'CCur(1E+16)'],
		['Main = CStr(Exp(710))', 'Exp(710)'],
		['Main = CStr(CDate(2958466))', 'CDate(2958466)'],
		['Main = Len(CStr(CInt(40000)))', 'CInt(40000)'],
		['Main = Mid$(String:=CStr(CInt(40000)), Start:=1)', 'CInt(40000)'],
		['Main = IIf(True, 1, CInt(40000))', 'CInt(40000)'],
		['Main = IIf(True, 0, CByte(300))', 'CByte(300)'],
		['Main = IIf(True, 0, 200 * 200)', '200 * 200'],
		['Main = IIf(True, 0, 32767% + 1%)', '32767% + 1%'],
		['Main = "x" & CInt(40000)', 'CInt(40000)'],
		['Main = "x" & 200 * 200', '200 * 200'],
		['Main = Format$(CInt(40000), "0")', 'CInt(40000)'],
		['Main = Choose(1, CInt(40000))', 'CInt(40000)'],
		['Main = Array(CInt(40000))(0)', 'CInt(40000)'],
		['Debug.Print 200 * 200', '200 * 200'],
		['Main = (200 * 200) > 1', '200 * 200'],
		['Main = Not (200 * 200)', '200 * 200'],
		['Main = CStr(Val("1e400"))', 'Val("1e400")'],
		// The sum does not fold, since Len is not folded; its CInt still does.
		['Main = Len("a") + CInt(40000)', 'CInt(40000)'],
	])('%s', (line, marked) => {
		const src = main(line);
		const hits = byCode(analyzeModule(src), CODE);
		expect(hits).toHaveLength(1);
		expect(src.slice(hits[0].span.start, hits[0].span.end)).toBe(marked);
	});

	it('reaches an array index and a project function argument', () => {
		expect(byCode(analyzeModule(main('Dim z(10) As Variant', 'Main = z(CInt(40000))')), CODE)).toHaveLength(1);
		const echo = 'Option Explicit\nFunction Main() As Variant\n    Main = Echo(CInt(40000))\nEnd Function\nPrivate Function Echo(ByVal v As Variant) As Variant\n    Echo = v\nEnd Function\n';
		expect(byCode(analyzeModule(echo), CODE)).toHaveLength(1);
	});

	it('knows a local IIf does not guard: IIf evaluates both branches', () => {
		expect(byCode(analyzeModule(main('Dim n As Long', 'n = 40000', 'Main = IIf(n > 32767, 0, CInt(n))')), CODE)).toHaveLength(1);
	});

	it('reports an overflow once, however many parts reach it', () => {
		expect(byCode(analyzeModule(main('Main = CInt(40000)')), CODE)).toHaveLength(1);
		expect(byCode(analyzeModule(main('Main = 1 + CInt(40000)')), CODE)).toHaveLength(1);
	});

	it.each([
		'Main = CStr(CInt(32767))',
		'Main = IIf(True, 0, CInt(32767))',
		'Main = "x" & 200 * 100',
		// Date arithmetic, not an Integer sum: its tail is not a part of its own.
		'Main = Date - 32767% - 2%',
		'Main = Now + 32767% + 1',
		'Main = CStr(CLngLng(1E+18) * 9)',
	])('leaves %s alone', (line) => {
		expect(byCode(analyzeModule(main(line)), CODE)).toEqual([]);
	});
});

describe('arithmetic-overflow: LongLong and LongPtr (issue #232)', () => {
	it.each([
		['Main = CLngLng(1E+19)', 'does not fit LongLong'],
		['Main = CLngLng("9223372036854775808")', 'does not fit LongLong'],
		['Main = CLngPtr(1E+19)', "does not fit a LongPtr, whose range is at most a LongLong's"],
		['Main = 9223372036854775807^ + 1', '9223372036854775807 (LongLong) + 1 (Integer) is 9223372036854775808, outside the LongLong range'],
		['Main = -9223372036854775807^ - 2', 'is -9223372036854775809, outside the LongLong range'],
		['Main = 4611686018427387904^ * 2', 'is 9223372036854775808, outside the LongLong range'],
	])('%s', (line, text) => {
		const hits = byCode(analyzeModule(main(line)), CODE);
		expect(hits).toHaveLength(1);
		expect(hits[0].message).toContain(text);
	});

	it('judges an assignment to a LongLong or a LongPtr', () => {
		const ll = byCode(analyzeModule(main('Dim x As LongLong', 'x = 1E+19', 'Main = x')), CODE);
		expect(ll.map((d) => d.message)).toEqual([expect.stringContaining('in a LongLong, whose range is -9223372036854775808 to 9223372036854775807')]);
		const lp = byCode(analyzeModule(main('Dim p As LongPtr', 'p = 1E+19', 'Main = p')), CODE);
		expect(lp.map((d) => d.message)).toEqual([expect.stringContaining('in a LongPtr, which holds no more than a LongLong')]);
		expect(byCode(analyzeModule(main('Dim x As LongLong', 'x = 9223372036854775807^ + 1', 'Main = x')), CODE)).toHaveLength(1);
	});

	it.each([
		'Main = CLngLng("9223372036854775807")',
		'Main = CLngLng(9E+18)',
		'Main = &H7FFFFFFFFFFFFFFF^',
		'Main = 9223372036854775807^ - 1',
		'Main = 9223372036854775807^ \\ 2',
		'Main = -9223372036854775807^ - 1',
		// A LongPtr is a Long on 32-bit Office; inside LongLong's range it is not judged.
		'Main = CLngPtr(2147483647)',
		'Main = CLngPtr(3000000000#)',
	])('leaves %s alone', (line) => {
		expect(byCode(analyzeModule(main(line)), CODE)).toEqual([]);
		expect(byCode(analyzeModule(main('Dim p As LongPtr', 'p = 3000000000#', 'Main = p')), CODE)).toEqual([]);
	});

	it('reports a LongLong counter stepping past the top, and not one that stops short', () => {
		const hits = byCode(analyzeModule(main('Dim i As LongLong', 'For i = 9223372036854775806^ To 9223372036854775807^', 'Next', 'Main = 1')), 'for-counter-overflow');
		expect(hits.map((d) => d.message)).toEqual([expect.stringContaining('after its last pass at 9223372036854775807 the loop adds 1')]);
		expect(byCode(analyzeModule(main('Dim i As LongLong', 'For i = 9223372036854775805^ To 9223372036854775806^', 'Next', 'Main = 1')), 'for-counter-overflow')).toEqual([]);
	});

	it('reports a LongLong Const past the range at compile time', () => {
		const src = 'Option Explicit\nPrivate Const C As LongLong = 1E+19\nFunction Main() As Variant\n    Main = 1\nEnd Function\n';
		expect(byCode(analyzeModule(src), 'const-overflow').map((d) => d.message)).toEqual([expect.stringContaining('declared As LongLong but its value 10000000000000000000 is outside that range')]);
		expect(byCode(analyzeModule('Option Explicit\nPrivate Const C As LongLong = 9223372036854775807^\n'), 'const-overflow')).toEqual([]);
	});

	it('reports a float literal past LongLong passed to a LongLong or LongPtr parameter', () => {
		for (const type of ['LongLong', 'LongPtr']) {
			const src = `Option Explicit\nFunction Main() As Variant\n    Main = Take(1E+19)\nEnd Function\nPrivate Function Take(ByVal v As ${type}) As Variant\n    Take = v\nEnd Function\n`;
			const hits = analyzeModule(src).filter((d) => d.message.includes("Run-time error '6': Overflow"));
			expect(hits, type).toHaveLength(1);
			expect(hits[0].message, type).toContain(`outside the ${type} range`);
		}
		const fits = 'Option Explicit\nFunction Main() As Variant\n    Main = Take(9E+18)\nEnd Function\nPrivate Function Take(ByVal v As LongLong) As Variant\n    Take = v\nEnd Function\n';
		expect(analyzeModule(fits).filter((d) => d.message.includes('Overflow'))).toEqual([]);
	});
});
