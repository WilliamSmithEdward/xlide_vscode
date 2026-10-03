import { afterEach, describe, expect, it, vi } from 'vitest';
import { writtenNamesIn } from '../src/analyzer/diagnostics/moduleState';
import * as lexer from '../src/analyzer/lexer/tokenize';

afterEach(() => { vi.restoreAllMocks(); });

describe('module write enclosing-callee scan', () => {
	it('reads a wide call token stream linearly', () => {
		const count = 1000;
		const source = `Fill(${Array.from({ length: count }, (_, i) => `v${i}`).join(',')})`;
		let reads = 0;
		const tokens = lexer.tokenizeCached(source).map(token => new Proxy(token, {
			get(target, key, receiver) {
				if (key === 'rawText') reads++;
				return Reflect.get(target, key, receiver);
			},
		}));
		vi.spyOn(lexer, 'tokenizeCached').mockReturnValue(tokens);
		const names = writtenNamesIn(source);
		expect(names.size).toBe(count);
		expect(names.has('v0')).toBe(true);
		expect(names.has(`v${count - 1}`)).toBe(true);
		expect(reads).toBeLessThan(count * 100);
	});

	it('keeps nested runtime and unknown callees separate', () => {
		expect([...writtenNamesIn('Fill(Abs(x), Inner(y), z)')]).toEqual(['y', 'z']);
		expect([...writtenNamesIn('result = Abs(x) + Inner(y)')]).toEqual(['result', 'y']);
	});

	it('handles bare calls, receivers and parenthesized arguments', () => {
		expect([...writtenNamesIn('Fill x, y\nobj.Fill z, q')]).toEqual(['x', 'y', 'z', 'q']);
		expect([...writtenNamesIn('Fill (x), y')]).toEqual(['x', 'y']);
	});

	it('recognizes source procedures that shadow runtime functions', () => {
		expect([...writtenNamesIn('Sub Abs()\nEnd Sub\nCall Abs(x, y)')]).toEqual(['x', 'y']);
		expect([...writtenNamesIn('Call Abs(x, y)')]).toEqual([]);
	});

	it('retains conservative results for unmatched parentheses', () => {
		expect([...writtenNamesIn('Fill(x, y')]).toEqual(['x', 'y']);
		expect([...writtenNamesIn('Fill(x)), y')]).toEqual(['x', 'y']);
		expect([...writtenNamesIn('Fill(Abs(x), Inner(y')]).toEqual(['y']);
	});
});
