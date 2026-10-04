import {describe, expect, it} from 'vitest';
import {classifyReferenceKinds} from '../src/analyzer/references/referenceKinds';
import {tokenizeCached} from '../src/analyzer/lexer/tokenize';

describe('sparse reference statement work', () => {
	it.each(['\n', '\r\n', '\r'].flatMap(eol => ['first', 'middle', 'last', 'separated'].map(position => ({eol, position}))))('skips unrelated token kinds at $position with EOL $eol', ({eol, position}) => {
		const rows = Array.from({length: 1000}, (_, i) => `unrelated${i} = 1`);
		const at = position === 'first' || position === 'separated' ? 0 : position === 'last' ? rows.length : 500;
		rows.splice(at, 0, 'target = target + 1');
		if (position === 'separated') rows.push('target = 7');
		const source = rows.join(eol), start = source.indexOf('target'), read = start + 9;
		const offsets = position === 'separated' ? [source.lastIndexOf('target'), read, start] : [read, start];
		const tokens = tokenizeCached(source), descriptors = tokens.map(token => Object.getOwnPropertyDescriptor(token, 'kind')!);
		let unrelatedReads = 0;
		for (const token of tokens) {
			const kind = token.kind;
			const isUnrelated = token.line !== at && !(position === 'separated' && token.line === rows.length - 1);
			Object.defineProperty(token, 'kind', {configurable: true, get() {if (isUnrelated) unrelatedReads++; return kind;}});
		}
		let result;
		try {result = classifyReferenceKinds(source, Object.freeze(offsets));} finally {
			for (let i = 0; i < tokens.length; i++) Object.defineProperty(tokens[i], 'kind', descriptors[i]);
		}
		const expected = [[start, 'write'], ...(position === 'separated' ? [[source.lastIndexOf('target'), 'write']] : []), [read, 'read']];
		expect([...result]).toEqual(expected);
		expect(unrelatedReads).toBeLessThanOrEqual(6);
	});

	it('keeps unmatched, duplicate and nonfinite offsets in their original default order', () => {
		const source = 'a = 1\n'.repeat(100), last = source.lastIndexOf('a');
		expect([...classifyReferenceKinds(source, [NaN, last, -1, last, Infinity, -Infinity, 1.5])]).toEqual([
			[last, 'write'], [NaN, 'read'], [-1, 'read'], [Infinity, 'read'], [-Infinity, 'read'], [1.5, 'read'],
		]);
	});

	it('classifies the complete continued statement and each colon-separated neighbor', () => {
		const prefix = 'unrelated = 1\n'.repeat(100);
		const source = prefix + 'If condition Then target = _\n 7 Else target = 8: target = target + 1';
		const starts = [...source.matchAll(/target/g)].map(match => match.index);
		expect([...classifyReferenceKinds(source, starts.toReversed())]).toEqual([
			[starts[0], 'write'], [starts[1], 'write'], [starts[2], 'write'], [starts[3], 'read'],
		]);
	});
});

it('keeps complete ordering when a later long statement requires the streaming fallback', () => {
	const prefix = 'first = 1\n' + 'unrelated = 1\n'.repeat(100);
	const source = prefix + 'target = f(' + Array(1000).fill('arg').join(',') + ')\nlast = target';
	const arg = source.lastIndexOf('arg'), target = source.indexOf('target'), last = source.indexOf('last');
	expect([...classifyReferenceKinds(source, [last, arg, 0, target])]).toEqual([
		[0, 'write'], [target, 'write'], [last, 'write'], [arg, 'read'],
	]);
});
