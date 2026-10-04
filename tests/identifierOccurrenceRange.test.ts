import { describe, expect, it } from 'vitest';
import { findIdentifierOccurrences, findIdentifierOccurrencesForNames } from '../src/vbaSourceScan';

describe('identifier occurrence range', () => {
	it.each(['\n', '\r\n', '\r'])('retains source positions and comment context for %j', eol => {
		const source = ["' comment _", 'café = 9', 'café = 1: Debug.Print CAFÉ', 'Rem café _', 'café = 8', 'café = 2', ''].join(eol);
		const start = source.indexOf('café = 1'), end = source.indexOf('café = 2');
		const range = Object.freeze({start, end});
		const expected = [
			{line: 2, column: 0, offset: start, text: 'café'},
			{line: 2, column: 22, offset: start + 22, text: 'CAFÉ'},
			{line: 5, column: 0, offset: end, text: 'café'},
		];
		expect(findIdentifierOccurrences(source, 'CAFÉ', range)).toEqual(expected);
		expect(findIdentifierOccurrencesForNames(source, ['CAFÉ', 'missing', 'café'], range)).toEqual(new Map([['café', expected], ['missing', []]]));
	});

	it('bounds occurrence starts without truncating words at either end', () => {
		const source = 'target target target';
		expect(findIdentifierOccurrences(source, 'target', {start: 1, end: 7})).toEqual([{line: 0, column: 7, offset: 7, text: 'target'}]);
		expect(findIdentifierOccurrences(source, 'target', {start: 0, end: 0})).toEqual([{line: 0, column: 0, offset: 0, text: 'target'}]);
		expect(findIdentifierOccurrences(source, 'target', {start: 14, end: 13})).toEqual([]);
		expect(findIdentifierOccurrences(source, 'target', {start: source.length, end: source.length + 20})).toEqual([]);
		expect(findIdentifierOccurrences(source, 'target')).toHaveLength(3);
	});
});
