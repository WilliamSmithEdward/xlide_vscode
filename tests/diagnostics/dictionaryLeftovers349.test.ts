// Diagnostics tests: Dictionary state through Keys read whole, a For Each, and
// a negative Items index (issue #349). Each case was run through pyVBAharness
// on 2026-10-02 in Excel 16.0 64-bit (build 20430).

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function errors(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    Dim d As Object, x As Variant\n    Set d = CreateObject("Scripting.Dictionary")\n    ${body}\n    If IsEmpty(Main) Then Main = 1\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('an early-bound Dictionary without the Scripting reference (issue #349)', () => {
	const EXCEL_DEFAULTS = ['VBA', 'Excel', 'stdole', 'Office'];
	const found = (body: string, referencedLibraries: readonly string[] | undefined): string[] => {
		const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
		return analyzeModule(src, { referencedLibraries }).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
	};

	it('does not compile in a project of Excel\'s default references', () => {
		expect(found('Dim d As Scripting.Dictionary\n    Set d = New Scripting.Dictionary\n    Main = d.Count', EXCEL_DEFAULTS)).toEqual(['missing-library-reference']);
		expect(found('Dim d As New Dictionary\n    Main = d.Count', EXCEL_DEFAULTS)).toEqual(['missing-library-reference']);
	});

	it('compiles with the reference, and is not judged without the list', () => {
		expect(found('Dim d As New Dictionary\n    Main = d.Count', [...EXCEL_DEFAULTS, 'Scripting'])).toEqual([]);
		expect(found('Dim d As New Dictionary\n    Main = d.Count', undefined)).toEqual([]);
		expect(found('Dim d As Object\n    Set d = CreateObject("Scripting.Dictionary")\n    Main = d.Count', EXCEL_DEFAULTS)).toEqual([]);
	});

	it('offers the reference fix for qualified and unqualified Scripting types', () => {
		for (const type of ['Dictionary', 'Scripting.Dictionary', 'FileSystemObject', 'TextStream']) {
			const src = `Option Explicit\nPrivate value As ${type}\n`;
			const diagnostic = analyzeModule(src, { referencedLibraries: EXCEL_DEFAULTS })
				.find((diag) => diag.code === 'missing-library-reference');
			expect(diagnostic?.data?.addLibraryReference, type).toEqual({ library: 'scripting' });
		}
	});

	it('leaves a Dictionary the project declares itself', () => {
		const src = 'Option Explicit\nPrivate Type Dictionary\n    n As Long\nEnd Type\nFunction Main() As Variant\n    Dim d As Dictionary\n    Main = d.n\nEnd Function\n';
		expect(analyzeModule(src, { referencedLibraries: EXCEL_DEFAULTS }).filter((diag) => diag.code === 'missing-library-reference')).toEqual([]);
	});
});

describe('a Dictionary read whole (issue #349)', () => {
	it('keeps its keys through UBound(d.Keys) and a For Each', () => {
		expect(errors('Main = UBound(d.Keys)\n    d.Remove "A"')).toEqual(['collection-key-not-found']);
		expect(errors('For Each x In d\n    Next\n    d.Remove "A"')).toEqual(['collection-key-not-found']);
		expect(errors('d.Add "a", 1\n    For Each x In d\n    Next\n    d.Remove "a"\n    Main = d.Count')).toEqual([]);
	});

	it('forgets them where the loop body may change it', () => {
		expect(errors('d.Add "a", 1\n    For Each x In d.Keys\n        d.Remove x\n    Next\n    d.Remove "b"')).toEqual([]);
		expect(errors('d.Add "a", 1\n    For Each x In d\n        d.RemoveAll\n    Next\n    d.Remove "a"')).toEqual([]);
	});

	it('refuses a negative index into Keys or Items', () => {
		expect(errors('d.Add "a", 1\n    Main = d.Items()(-1)')).toEqual(['collection-index-out-of-range']);
		expect(errors('Main = d("a")\n    Main = d.Items()(-1)')).toEqual(['collection-index-out-of-range']);
		expect(errors('d.Add "a", 1\n    Main = d.Keys()(0)')).toEqual([]);
	});
});
