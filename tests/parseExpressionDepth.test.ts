import { describe, it, expect } from 'vitest';
import { parseModule } from '../src/analyzer/parser/parseModule';

// The recursive-descent expression parser documents a "Never throws" contract.
// Pathological nesting must not overflow the JS stack with a RangeError.
describe('expression parser recursion-depth guard', () => {
	it('does not throw on deeply nested parentheses', () => {
		const src = `Sub Foo()\n    x = ${'('.repeat(3000)}1${')'.repeat(3000)}\nEnd Sub`;
		expect(() => parseModule(src)).not.toThrow();
	});

	it('does not throw on a long unary chain', () => {
		const src = `Sub Foo()\n    x = ${'-'.repeat(20000)}1\nEnd Sub`;
		expect(() => parseModule(src)).not.toThrow();
	});
});
