import { afterEach, describe, expect, it, vi } from 'vitest';
import { elementsWrittenIn } from '../src/analyzer/diagnostics/rules/arrays';
import { parseModule } from '../src/analyzer/parser/parseModule';
import { tokenizeCached } from '../src/analyzer/lexer/tokenize';
import * as walker from '../src/analyzer/diagnostics/walker';

function fixture(statement: string) {
	const source = `Sub P()\n${statement}\nEnd Sub`;
	const proc = parseModule(source).members.find(member => member.kind === 'Procedure');
	if (!proc || proc.kind !== 'Procedure') throw new Error('Missing procedure');
	return { source, proc };
}
function writes(statement: string) {
	const { source, proc } = fixture(statement);
	return [...elementsWrittenIn(source, proc, undefined)];
}
afterEach(() => { vi.restoreAllMocks(); });

describe('array element write equality depth', () => {
	it('scans nested comparisons once rather than rescanning their prefixes', () => {
		const count = 1000;
		const statement = `Fill(${Array(count).fill('(a = a)').join(' And ')}, target(0))`;
		const { source, proc } = fixture(statement);
		let reads = 0;
		const tokens = tokenizeCached(statement).map(token => new Proxy(token, {
			get(target, key, receiver) {
				if (key === 'rawText') reads++;
				return Reflect.get(target, key, receiver);
			},
		}));
		vi.spyOn(walker, 'statementTokensAfterLeadingLabel').mockReturnValue(tokens);
		expect([...elementsWrittenIn(source, proc, undefined)]).toEqual(['target']);
		expect(reads).toBeLessThan(count * 100);
	});

	it.each([
		['Fill target(0)', ['target']],
		['Fill((a = a), target(0))', ['target']],
		['target(0) = 1', ['target']],
		['Let target(0) = 1', ['target']],
		['result = target(0)', []],
		['result = Fill((a = a), target(0))', ['target']],
		['Fill("=", target(0))', ['target']],
		['Fill((a = a), target(0)', ['target']],
		['Fill((a = a)), target(0)', ['target']],
	] as const)('preserves writes in %s', (statement, expected) => {
		expect(writes(statement)).toEqual(expected);
	});
});
