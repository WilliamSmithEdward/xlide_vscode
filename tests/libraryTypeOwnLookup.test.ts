import {describe, expect, it, vi} from 'vitest';
import {libraryTypeNames} from '../src/analyzer/host/libraryTypeNames';
import {analyzeModule} from '../src/analyzer';

describe('referenced library type lookup', () => {
	it.each(['constructor', 'CONSTRUCTOR', '__proto__'])('treats inherited name %s as an unknown library', name => {
		expect(libraryTypeNames(name)).toBeUndefined();
	});

	it.each(['constructor', '__proto__'])('analyzes an unknown declared type with reference %s', reference => {
		const source = 'Option Explicit\nPrivate value As NeverDefinedType\nSub Main()\nDebug.Print value\nEnd Sub\n';
		const failures: unknown[] = [];
		expect(analyzeModule(source, {referencedLibraries: ['VBA', 'Excel', reference], onInternalError(error, where) { failures.push({error, where}); }}).filter(d => d.severity === 'error')).toEqual([]);
		expect(failures).toEqual([]);
	});

	it('normalizes the library spelling once per lookup', () => {
		libraryTypeNames('Excel');
		const original = String.prototype.toLowerCase;
		let calls = 0;
		const spy = vi.spyOn(String.prototype, 'toLowerCase').mockImplementation(function(this: string) {
			if (String(this) === 'EXCEL') calls++;
			return original.call(this);
		});
		try {expect(libraryTypeNames('EXCEL')?.has('worksheet')).toBe(true);} finally {spy.mockRestore();}
		expect(calls).toBe(1);
	});

	it('keeps complete cached sets case insensitive and leaves unknown lookups isolated', () => {
		for (const library of ['VBA', 'Excel', 'Word', 'PowerPoint', 'Access', 'Office', 'stdole', 'MSForms']) {
			const first = libraryTypeNames(library);
			expect(first?.size).toBeGreaterThan(0);
			expect(libraryTypeNames(library.toUpperCase())).toBe(first);
			expect(libraryTypeNames('unknown_' + library)).toBeUndefined();
			expect(libraryTypeNames(library)).toBe(first);
		}
	});
});
