import { afterEach, describe, expect, it, vi } from 'vitest';
import { arrayValueShape } from '../src/analyzer/diagnostics/rules/arrays';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';

function tokens(text: string, delimiter: string, limit = -1, compare = 'vbTextCompare') {
	const quote = (value: string) => `"${value.replaceAll('"', '""')}"`;
	return tokenizeCached(`Split(${quote(text)}, ${quote(delimiter)}, ${limit}, ${compare})`);
}
function shape(text: string, delimiter: string, limit = -1, compare = 'vbTextCompare') {
	return arrayValueShape(tokens(text, delimiter, limit, compare), 'v', 0);
}
afterEach(() => { vi.restoreAllMocks(); });

describe('Split caseless search preparation', () => {
	it('folds the source once rather than once per delimiter', () => {
		const text = 'AbX'.repeat(1000);
		const input = tokens(text, 'x');
		const original = String.prototype.toLowerCase;
		let foldedCharacters = 0;
		vi.spyOn(String.prototype, 'toLowerCase').mockImplementation(function (this: string) {
			if (String(this) === text) foldedCharacters += text.length;
			return original.call(this);
		});
		const result = arrayValueShape(input, 'v', 0);
		expect(result?.values).toEqual([...Array(1000).fill('Ab'), '']);
		expect(result?.dims[0].upper).toBe(1000);
		expect(foldedCharacters).toBeLessThanOrEqual(text.length * 2);
	});

	it.each([
		['xAAXbX', 'x', -1, ['', 'AA', 'b', '']],
		['xAAXbX', 'x', 2, ['', 'AAXbX']],
		['ABabAB', 'ab', -1, ['', '', '', '']],
		['NoMatch', 'xyz', -1, ['NoMatch']],
		['A"BXc', 'x', -1, ['A"B', 'c']],
	] as const)('retains original text parts for %j / %j / limit %i', (text, delimiter, limit, values) => {
		expect(shape(text, delimiter, limit)?.values).toEqual(values);
	});

	it('leaves binary comparison case-sensitive', () => {
		expect(shape('AbXcdx', 'x', -1, 'vbBinaryCompare')?.values).toEqual(['AbXcd', '']);
	});

	it('keeps non-ASCII comparison conservative, including delimiters', () => {
		expect(shape('ÉXtail', 'x')).toBeUndefined();
		expect(shape('ascii', 'é')).toBeUndefined();
		expect(shape('ÉXtail', 'X', -1, 'vbBinaryCompare')?.values).toEqual(['É', 'tail']);
	});

	it('does not compare text for limit one, limit zero or an empty input', () => {
		expect(shape('ÉXtail', 'é', 1)?.values).toEqual(['ÉXtail']);
		expect(shape('ÉXtail', 'é', 0)?.dims[0].upper).toBe(-1);
		expect(shape('', 'é')?.dims[0].upper).toBe(-1);
	});
});
