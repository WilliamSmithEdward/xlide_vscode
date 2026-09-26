// Diagnostics tests: array bounds the code makes plain (issue #120). Each
// raising sample was measured in Excel 16.0 (build 20326, 2026-09-26); each
// quiet one runs there.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';
import { byCode, expectDiagnostic, expectDiagnostics } from '../helpers/diagnostics';

const CODE = 'array-subscript-out-of-bounds';

function wrap(...lines: string[]): string {
	return `Option Explicit\nFunction Main() As Variant\n${lines.map((line) => `    ${line}`).join('\n')}\nEnd Function\n`;
}

describe('array-subscript-out-of-bounds - bounds from values (issue #120)', () => {
	it('flags an index past the parts Split on literals returns, inline and through a local', () => {
		const src = wrap(
			'Dim a() As String',
			'Main = Split("abc", ",")(1)',
			'a = Split("a,b", ",")',
			'Main = a(2)',
			'Main = a(1)',
		);
		expectDiagnostics(src, analyzeModule(src), CODE, [
			{ span: '1', message: 'Split returns' },
			{ span: '2', message: ['Split(...)', 'upper bound 1'] },
		]);
	});

	it('flags an index past Array(...) and any index into Array()', () => {
		const src = wrap('Dim v As Variant, w As Variant', 'v = Array(1, 2)', 'Main = v(2)', 'w = Array()', 'Main = w(0)', 'Main = v(1)');
		expectDiagnostics(src, analyzeModule(src), CODE, [
			{ span: '2', message: ['Array(...)', 'upper bound 1'] },
			{ span: '0', message: 'empty' },
		]);
	});

	it('flags index 0 into the two-dimensional array a Range literal gives', () => {
		const src = wrap('Dim v As Variant', 'v = Range("A1:B2").Value', 'Main = v(0, 0)', 'Main = v(2, 2)');
		expectDiagnostic(src, analyzeModule(src), CODE, { span: '0', message: ['Range(...).Value', 'lower bound 1'] });
	});

	it('follows Option Base for Dim and Array, and not for VBA.Array', () => {
		const src =
			'Option Explicit\nOption Base 1\nFunction Main() As Variant\n' +
			'    Dim a(3) As Long, v As Variant, w As Variant\n' +
			'    Main = a(0)\n' +
			'    Main = a(3)\n' +
			'    v = VBA.Array(1, 2, 3)\n' +
			'    Main = v(3)\n' +
			'    w = Array(1, 2, 3)\n' +
			'    Main = w(3)\n' +
			'End Function\n';
		expectDiagnostics(src, analyzeModule(src), CODE, [
			{ span: '0', message: 'Option Base 1' },
			{ span: '3', message: 'VBA.Array(...)' },
		]);
	});

	it('checks every dimension of a multi-dimensional Dim', () => {
		const src = wrap('Dim a(1 To 3, 1 To 2) As Long', 'Main = a(2, 3)', 'Main = a(3, 2)');
		expectDiagnostic(src, analyzeModule(src), CODE, { span: '3', message: 'dimension 2' });
	});

	it('flags UBound and LBound of a dimension the array has not got', () => {
		const src = wrap('Dim a(1 To 3) As Long', 'Main = UBound(a, 2)', 'Main = LBound(a, 1)');
		expectDiagnostic(src, analyzeModule(src), CODE, { span: '2', message: ['UBound', '1 dimension'] });
	});

	it('flags the last pass of a For whose counter indexes past the end', () => {
		const src = wrap('Dim a(2) As Long, i As Long', 'For i = 0 To 3', '    a(i) = i', 'Next', 'Main = a(2)');
		expectDiagnostic(src, analyzeModule(src), CODE, { span: 'i', message: ['last pass', 'reaches 3'] });
		const inRange = wrap('Dim a(2) As Long, i As Long', 'For i = 0 To 2', '    a(i) = i', 'Next', 'Main = a(2)');
		expect(byCode(analyzeModule(inRange), CODE)).toHaveLength(0);
	});

	it('stays quiet once something else could have shaped the array', () => {
		const src = wrap(
			'Dim v As Variant, a() As String',
			'v = Array(1, 2)',
			'If Main Then v = Array(1, 2, 3)',
			'Main = v(2)',
			'a = Split("a,b", ",")',
			'ReDim Preserve a(5)',
			'Main = a(4)',
		);
		expect(byCode(analyzeModule(src), CODE)).toHaveLength(0);
	});
});

describe('redim-impossible-bounds - implicit lower bound (issue #120)', () => {
	it('flags ReDim a(-1) against the Option Base lower bound', () => {
		const src = wrap('Dim a() As Long', 'ReDim a(-1)');
		expectDiagnostic(src, analyzeModule(src), 'redim-impossible-bounds', { span: '-1', message: 'Option Base 0' });
	});

	it('stays quiet for ReDim a(0) and for an explicit lower bound', () => {
		const src = wrap('Dim a() As Long', 'ReDim a(0)', 'ReDim a(-3 To -1)');
		expect(byCode(analyzeModule(src), 'redim-impossible-bounds')).toHaveLength(0);
	});
});
