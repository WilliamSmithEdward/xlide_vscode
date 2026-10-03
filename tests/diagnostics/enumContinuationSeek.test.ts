import { describe, expect, it, vi } from 'vitest';
import { checkLineContinuationLimits } from '../../src/analyzer/diagnostics/rules/lineContinuations';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { tokenizeCached } from '../../src/analyzer/lexer/tokenize';
import type { Span } from '../../src/analyzer/parser/nodes';

vi.mock('../../src/analyzer/lexer/tokenize', async () => {
	const actual = await vi.importActual<typeof import('../../src/analyzer/lexer/tokenize')>('../../src/analyzer/lexer/tokenize');
	return { ...actual, tokenizeCached: vi.fn(actual.tokenizeCached) };
});

describe('Enum continuation seeking', () => {
	it('does not compare every continuation with every Enum', () => {
		const count = 250;
		const source = Array.from({ length: count }, (_, i) => 'Public Enum E' + i + '\nV' + i + ' = _\n 1\nEnd Enum\n').join('');
		const module = parseModule(source);
		let reads = 0;
		const tokens = tokenizeCached(source).map(token => ({ ...token,
			leadingTrivia: token.leadingTrivia?.map(trivia => ({ ...trivia,
				get start() { reads++; return trivia.start; },
			})),
		}));
		vi.mocked(tokenizeCached).mockReturnValueOnce(tokens);
		const push = vi.fn();
		checkLineContinuationLimits(source, module, push);
		expect(push).toHaveBeenCalledTimes(count);
		expect(reads).toBeLessThan(count * 30);
	});

	it.each(['\n', '\r\n', '\r'])('keeps header and unrelated continuations out of Enum bodies for %j', eol => {
		const source = ['Public _', 'Enum One', 'First = _', ' 1', 'End Enum', 'Enum Two', 'Second = 2', 'End Enum', 'Sub Main()', 'Dim value As Long', 'value = _', ' 1', 'End Sub'].join(eol);
		const found: Array<{ message: string; span: Span }> = [];
		checkLineContinuationLimits(source, parseModule(source), (_kind, message, span) => { found.push({ message, span }); });
		expect(found).toHaveLength(1);
		expect(found[0].message).toContain("inside Enum 'One'");
		expect(found[0].span.start).toBe(source.indexOf('_', source.indexOf('First')) - 1);
	});
});
