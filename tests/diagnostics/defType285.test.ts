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
