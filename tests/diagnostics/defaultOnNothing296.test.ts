// Diagnostics tests: the default member of a never-set object, reached
// without its name (issue #296). Measured in Excel 16.0 64-bit (2026-10-02)
// through pyVBAharness: each raises 91.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a never-set object indexed through its default member', () => {
	it('raises 91, in a Set\'s value too', () => {
		expect(found('Dim o As Collection\n    Main = o(1)')).toEqual(['object-variable-not-set']);
		expect(found('Dim o As Collection, p As Collection\n    Set p = o(1)\n    Main = 1')).toEqual(['object-variable-not-set']);
		expect(found('Dim r As Range\n    Main = r(1)')).toEqual(['object-variable-not-set']);
	});

	it('stays quiet once the object is set', () => {
		expect(found('Dim o As Collection\n    Set o = New Collection\n    o.Add 5\n    Main = o(1)')).toEqual([]);
		expect(found('Dim o As Collection, p As Collection\n    Set o = New Collection\n    o.Add New Collection\n    Set p = o(1)\n    Main = 1')).toEqual([]);
	});
});
