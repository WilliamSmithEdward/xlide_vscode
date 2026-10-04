import { describe, expect, it, vi } from 'vitest';
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

it('keeps whole-source stripping cached while scoped scans can opt out', () => {
    const source = "' occurrence-cache control _\r\nTarget = 0\r\nTarget = Target + Other\r\n";
    const range = { start: source.indexOf('Target = Target'), end: source.length };
    const expected = [
        { line: 2, column: 0, offset: range.start, text: 'Target' },
        { line: 2, column: 9, offset: range.start + 9, text: 'Target' },
    ];
    const originalSplit = String.prototype.split;
    let sourceSplits = 0;
    const spy = vi.spyOn(String.prototype, 'split').mockImplementation(function (...args: Parameters<typeof originalSplit>) {
        if (String(this) === source) { sourceSplits++; }
        return Reflect.apply(originalSplit, this, args);
    });
    try {
        expect(findIdentifierOccurrences(source, 'target', range)).toEqual(expected);
        expect(findIdentifierOccurrences(source, 'target', range)).toEqual(expected);
        expect(sourceSplits).toBe(1);
        expect(findIdentifierOccurrences(source, 'target', range, { cacheStrippedSource: false })).toEqual(expected);
        expect(findIdentifierOccurrencesForNames(source, ['target', 'missing'], range, { cacheStrippedSource: false }))
            .toEqual(new Map([['target', expected], ['missing', []]]));
        expect(sourceSplits).toBe(3);
        expect(findIdentifierOccurrences(source, 'target', range)).toEqual(expected);
        expect(sourceSplits).toBe(3);
    } finally { spy.mockRestore(); }
});
