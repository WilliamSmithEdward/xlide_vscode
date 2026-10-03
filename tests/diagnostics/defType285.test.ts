// Diagnostics tests: a DefType line types the names declared with no type
// (issue #285). Measured in Excel 16.0 64-bit (2026-10-02) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(defType: string, body: string, extra = ''): string[] {
	const src = `Option Explicit\n${defType}\nFunction Main() As Variant\n    ${body}\nEnd Function\n${extra}`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a name a DefType line types', () => {
	it('overflows as its type does', () => {
		expect(found('DefInt A-Z', 'Dim i\n    i = 40000\n    Main = i')).toContain('arithmetic-overflow');
		expect(found('DefInt I-K', 'Dim k\n    k = 40000\n    Main = k')).toContain('arithmetic-overflow');
		expect(found('DefByte B', 'Dim b\n    b = -1\n    Main = b')).toContain('arithmetic-overflow');
		expect(found('DefInt A-Z', 'Main = Big()', 'Function Big()\n    Big = 40000\nEnd Function\n')).toContain('arithmetic-overflow');
		expect(found('DefInt A-Z', 'Main = Twice(40000)', 'Function Twice(n)\n    Twice = n * 2\nEnd Function\n')).toHaveLength(1);
	});

	it('takes what its type takes', () => {
		expect(found('DefDate D', 'Dim d\n    d = "abc"\n    Main = d')).toEqual(['assignment-type-mismatch']);
		expect(found('DefObj O', 'Dim o\n    Main = o.Count')).toEqual(['object-variable-not-set']);
	});

	it('leaves other letters, a declared type and a type character alone', () => {
		expect(found('DefInt I-K', 'Dim a\n    a = 40000\n    Main = a')).toEqual([]);
		expect(found('DefInt A-Z', 'Dim i As Long\n    i = 40000\n    Main = i')).toEqual([]);
		expect(found('DefInt A-Z', 'Dim i&\n    i = 40000\n    Main = i')).toEqual([]);
		expect(found('DefLng A-Z', 'Dim i\n    i = 40000\n    Main = i')).toEqual([]);
	});
});

describe('a name assigned with no Dim and no Option Explicit', () => {
	function implicit(defType: string, body: string): string[] {
		const src = `${defType}\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
		return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
	}

	it('is a local of the DefType type', () => {
		expect(implicit('DefInt A-Z', 'i = 40000\n    Main = i')).toContain('arithmetic-overflow');
		expect(implicit('DefInt A-Z', 'For i = 1 To 40000\n    Next')).toContain('for-counter-overflow');
		expect(implicit('DefInt A-Z', 'n = 10\n    n = n * 4000\n    Main = n')).toContain('arithmetic-overflow');
		expect(implicit('DefDate D', 'd = "abc"\n    Main = d')).toEqual(['assignment-type-mismatch']);
	});

	it('stays quiet where the value fits, the letter has no DefType, or the name is a host global', () => {
		expect(implicit('DefInt A-Z', 'i = 1.5\n    Main = i')).toEqual([]);
		expect(implicit('DefLng A-Z', 'i = 40000\n    Main = i')).toEqual([]);
		expect(implicit('DefInt I', 'k = 40000\n    Main = k')).toEqual([]);
		expect(implicit('DefInt A-Z', 'StatusBar = False\n    Main = 1')).toEqual([]);
		expect(implicit('DefStr S', 's = 5\n    Main = s & "x"')).toEqual([]);
	});
});
