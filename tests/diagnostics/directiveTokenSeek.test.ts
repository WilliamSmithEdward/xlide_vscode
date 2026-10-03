import { describe, expect, it, vi } from 'vitest';
import { checkDirectiveForms } from '../../src/analyzer/diagnostics/rules/directiveForms';
import { parseModule } from '../../src/analyzer/parser/parseModule';
import { tokenizeCached } from '../../src/analyzer/lexer/tokenize';
import type { Span } from '../../src/analyzer/parser/nodes';

vi.mock('../../src/analyzer/lexer/tokenize', async () => {
	const actual = await vi.importActual<typeof import('../../src/analyzer/lexer/tokenize')>('../../src/analyzer/lexer/tokenize');
	return { ...actual, tokenizeCached: vi.fn(actual.tokenizeCached) };
});

describe('directive token seeking', () => {
	it('does not revisit every preceding token for each directive', () => {
		const count = 500;
		const source = Array.from({ length: count }, (_, i) => '#Const FLAG' + i + ' = 1\n').join('');
		const module = parseModule(source);
		const tokens = tokenizeCached(source);
		let reads = 0;
		const counted = new Proxy(tokens, { get(target, key, receiver) {
			if (typeof key === 'string' && /^\d+$/.test(key)) { reads++; }
			return Reflect.get(target, key, receiver);
		} });
		vi.mocked(tokenizeCached).mockReturnValueOnce(counted);
		const push = vi.fn();
		checkDirectiveForms(source, module, undefined, push);
		expect(push).not.toHaveBeenCalled();
		expect(reads).toBeLessThan(count * 30);
	});

	it.each(['\n', '\r\n', '\r'])('preserves line bounds and ignores comment-only colons for %j', eol => {
		const source = ['#Const A = 1: Debug.Print 1', "#Const B = 2: 'comment", '#Const C = 3', '#Const D = 4:Debug.Print 4'].join(eol);
		const found: Array<{ kind: string; span: Span }> = [];
		checkDirectiveForms(source, parseModule(source), undefined, (kind, _message, span) => { found.push({ kind, span }); });
		expect(found.map(item => item.kind)).toEqual(['directiveTrailingStatement', 'directiveTrailingStatement']);
		expect(found.map(item => source.slice(item.span.start, item.span.end))).toEqual(['Debug.Print 1', 'Debug.Print 4']);
	});
});
