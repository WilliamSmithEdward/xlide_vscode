import { beforeEach, describe, expect, it, vi } from 'vitest';
import { formatVbaModule, tokenStreamDifference } from '../src/analyzer/format/formatModule';
import { tokenize } from '../src/analyzer/lexer/tokenize';

vi.mock('../src/analyzer/lexer/tokenize', async () => {
	const actual = await vi.importActual<typeof import('../src/analyzer/lexer/tokenize')>('../src/analyzer/lexer/tokenize');
	return { ...actual, tokenize: vi.fn(actual.tokenize) };
});

beforeEach(() => { vi.mocked(tokenize).mockClear(); });

describe('formatter safety-check token reuse', () => {
	it('lexes the input once and validates the separately lexed output', () => {
		const source = 'sub t()\nx=1\nend sub';
		const result = formatVbaModule(source, { tabSize: 4, insertSpaces: true });
		expect(result.refusal).toBeUndefined();
		expect(result.text).toBe('Sub t()\n    x = 1\nEnd Sub');
		expect(vi.mocked(tokenize).mock.calls.map(call => call[0])).toEqual([source, result.text]);
	});

	it('still lexes both sources for standalone token comparison', () => {
		expect(tokenStreamDifference('x = 1', 'x = 2')).toBe('token 2 changed from "1" to "2" at line 1');
		expect(vi.mocked(tokenize).mock.calls.map(call => call[0])).toEqual(['x = 1', 'x = 2']);
	});

	it('still refuses formatting that moves continuation trivia', () => {
		const result = formatVbaModule(' _\r\n_\r\n', { tabSize: 0, insertSpaces: false });
		expect(result.text).toBeUndefined();
		expect(result.refusal).toContain('moved');
		expect(tokenize).toHaveBeenCalledTimes(2);
	});
});
