// Diagnostics tests: the last of issue #415, objects read as values. Each case
// was run through pyVBAharness on 2026-10-02 in Excel 16.0 64-bit (build
// 20430).

import { describe, it, expect } from 'vitest';
import { hostTokenForFileName } from '../../src/analyzer/host/hostRegistry';
import { analyzeVbaModuleSource } from '../../src/vbaModuleAnalysis';

function errors(body: string): string[] {
	const source = `Option Explicit\nFunction Main() As Variant\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeVbaModuleSource({ source, moduleName: 'Module1', host: hostTokenForFileName('Book1.xlsm'), referencedHosts: [] }).diagnostics
		.filter((diag) => diag.severity === 'error').map((diag) => String(diag.code));
}

const APP = 'Dim x As Application\n    Set x = Application\n    ';
const NAMES = 'Dim x As Names\n    Set x = ActiveWorkbook.Names\n    ';

describe('the Application read as a value (issue #415)', () => {
	it('is its Name, which is no number', () => {
		expect(errors(`${APP}Main = x + 1`)).toEqual(['assignment-type-mismatch']);
		expect(errors(`${APP}If x Then Main = 2`)).toEqual(['assignment-type-mismatch']);
		expect(errors(`${APP}If x = 0 Then Main = 2`)).toEqual(['assignment-type-mismatch']);
		expect(errors(`${APP}Main = x(1)`)).toEqual(['argument-count']);
	});

	it('reads as a String', () => {
		for (const body of ['Main = x', 'Main = x & "a"', 'If x = "Microsoft Excel" Then Main = 2', 'Main = x + "a"', 'Main = x.Name']) {
			expect(errors(`${APP}${body}`)).toEqual([]);
		}
	});
});

describe('Names read as a value (issue #415)', () => {
	it('needs the argument of its Item', () => {
		expect(errors(`${NAMES}Main = x`)).toEqual(['object-default-value']);
		expect(errors(`${NAMES}Main = x & "a"`)).toEqual(['object-default-value']);
		expect(errors(`${NAMES}Main = x.Count`)).toEqual([]);
	});
});

describe('a Collection that holds one (issue #415)', () => {
	it('raises 450 as a condition once set', () => {
		expect(errors('Dim x As Collection\n    Set x = New Collection\n    If x Then Main = 2')).toEqual(['object-default-value']);
		expect(errors('Dim x As Collection\n    Set x = New Collection\n    If x.Count = 0 Then Main = 2')).toEqual([]);
	});

	it('raises 5 indexed empty through an Object', () => {
		expect(errors('Dim x As Object\n    Set x = New Collection\n    Main = x(1)')).toEqual(['collection-index-out-of-range']);
		expect(errors('Dim x As Object\n    Set x = New Collection\n    x.Add 1\n    Main = x(1)')).toEqual([]);
	});
});
