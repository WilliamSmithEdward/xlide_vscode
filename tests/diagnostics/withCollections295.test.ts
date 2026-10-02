// Diagnostics tests: a Collection or Dictionary followed into a With block
// on it (issue #295). Measured in Excel 16.0 64-bit (2026-10-02) through
// pyVBAharness.

import { describe, it, expect } from 'vitest';
import { analyzeModule } from '../../src/analyzer';

function found(body: string): string[] {
	const src = `Option Explicit\nFunction Main() As Variant\n    ${body}\nEnd Function\n`;
	return analyzeModule(src).filter((diag) => diag.severity === 'error').map((diag) => diag.code);
}

describe('a With block on a Collection', () => {
	it('reads the subject\'s state through a leading dot', () => {
		expect(found('Dim c As New Collection\n    c.Add 1\n    With c\n        Main = .Item(2)\n    End With')).toEqual(['collection-index-out-of-range']);
		expect(found('With New Collection\n        .Add 1\n        Main = .Item(2)\n    End With')).toEqual(['collection-index-out-of-range']);
		expect(found('With New Collection\n        .Remove 1\n    End With')).toEqual(['collection-index-out-of-range']);
		expect(found('With New Collection\n        .Add 1, "a"\n        Main = .Item("b")\n    End With')).toEqual(['collection-key-not-found']);
		expect(found('With New Collection\n        .Add 1, "a"\n        .Add 2, "a"\n    End With')).toEqual(['collection-key-in-use']);
	});

	it('counts what the block adds', () => {
		expect(found('With New Collection\n        .Add 1\n        .Add 2\n        Main = .Item(2)\n    End With')).toEqual([]);
		expect(found('Dim c As New Collection\n    c.Add 1\n    With c\n        Main = .Item(1)\n    End With')).toEqual([]);
	});
});

describe('a With block on a Dictionary', () => {
	it('refuses a key it already holds', () => {
		expect(found('With CreateObject("Scripting.Dictionary")\n        .Add "a", 1\n        .Add "a", 2\n    End With')).toHaveLength(1);
	});
});
