// Diagnostics tests: the leftovers of #508, #509 and #510 (issue #559). Each
// case was run through pyVBAharness on 2026-10-02 in Excel 16.0 (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

function errors(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', host: 'excel' } as Parameters<typeof analyzeVbaModuleSource>[0]).diagnostics
		.filter((diag) => diag.severity === 'error')
		.map((diag) => diag.code ?? '');
}

describe('issue #559', () => {
	it('follows Item inside a Range chain', () => {
		expect(errors('Main = ActiveSheet.Range("C3:E7").Item(1).Offset(-4).Address')).toEqual(['host-argument-out-of-range']);
		expect(errors('Main = ActiveSheet.Range("XFD1").Item(2).Offset(1048577).Address')).toEqual(['host-argument-out-of-range']);
		expect(errors('Main = ActiveSheet.Range("C3:E7").Item(1).Offset(-2).Address')).toEqual([]);
		expect(errors('Main = ActiveSheet.Range("XFC1048575:XFD1048576").EntireRow.Item(7).Address')).toEqual(['host-argument-out-of-range']);
	});

	it('indexes Split of a String local', () => {
		expect(errors('Dim t As String\n    t = ""\n    Main = Split(t, ",")(2)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim s As String\n    s = "q,r"\n    Main = Split(s, ",")(3)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim s As String\n    s = "q,r"\n    Main = Split(s & "abc", ",")(-1)')).toEqual(['array-subscript-out-of-bounds']);
		expect(errors('Dim s As String\n    s = "q,r"\n    Main = Split(s, ",")(1)')).toEqual([]);
	});

	it('judges dates past the range', () => {
		expect(errors('Main = DateSerial(-10000, 1, 1)')).toEqual(['runtime-argument-value']);
		expect(errors('Main = DateSerial(-100, 1, 1)')).toEqual([]);
		expect(errors('Main = DateAdd("yyyy", -10000, CDate(-10000))')).toEqual(['runtime-argument-value']);
		expect(errors('Main = DateAdd("yyyy", 1, CDate(-10000))')).toEqual([]);
		expect(errors('Dim t As Date\n    t = #3/15/2023#\n    Main = DateAdd("m", 2958465, t)')).toEqual(['runtime-argument-value']);
		expect(errors('Dim t As Date\n    t = #3/15/2023#\n    Main = DateAdd("m", 1, t)')).toEqual([]);
		expect(errors('Dim t As Date\n    t = #3/15/2023#\n    Main = CDate(DateDiff("s", #2/29/2000#, t))')).toEqual(['runtime-conversion-value']);
		expect(errors('Dim t As Date\n    t = #3/15/2023#\n    Main = CDate(DateDiff("d", #2/29/2000#, t))')).toEqual([]);
	});
});
