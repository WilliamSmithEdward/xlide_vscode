import { describe, it, expect } from 'vitest';
import { completionCursorContext, identifierSpanEndingAt } from '../src/analyzer';

/** Context at the | marker in `src`. */
function contextAtMarker(src: string) {
	const offset = src.indexOf('|');
	if (offset < 0) {
		throw new Error('Missing | marker');
	}
	return completionCursorContext(src.replace('|', ''), offset);
}

describe('completionCursorContext', () => {
	it('peels the trailing partial identifier and exposes the preceding token', () => {
		const ctx = contextAtMarker('Sub T()\n    ws.Ran|\nEnd Sub\n');
		expect(ctx.partial).toBe('Ran');
		expect(ctx.partialToken?.rawText).toBe('Ran');
		expect(ctx.before?.rawText).toBe('.');
	});

	it('reports no partial when the cursor is not finishing a word', () => {
		const ctx = contextAtMarker('Sub T()\n    x = Foo |\nEnd Sub\n');
		expect(ctx.partial).toBe('');
		expect(ctx.partialToken).toBeUndefined();
		expect(ctx.before?.rawText).toBe('Foo');
	});

	it('filters comments from significant tokens but keeps newlines', () => {
		const ctx = contextAtMarker("x = 1 ' note\ny|");
		expect(ctx.significantTokens.some((t) => t.kind === 'comment')).toBe(false);
		expect(ctx.significantTokens.some((t) => t.kind === 'newline')).toBe(true);
		expect(ctx.tokens.some((t) => t.kind === 'comment')).toBe(true);
	});

	it('detects cursor positions inside comments and strings', () => {
		expect(contextAtMarker("x = 1 ' comm|").inComment).toBe(true);
		expect(contextAtMarker('x = "ab|').inString).toBe(true);
		expect(contextAtMarker('x = y|').inComment).toBe(false);
		expect(contextAtMarker('x = y|').inString).toBe(false);
	});

	it('tracks statement starts across newlines and colon separators', () => {
		expect(contextAtMarker('x = 1\ny = |').statementStart).toBe(6);
		const src = 'a = 1: b = |';
		expect(contextAtMarker(src).statementStart).toBe(src.indexOf(':') + 1);
		expect(contextAtMarker('x = |').statementStart).toBe(0);
	});

	it('agrees with the char-level identifier span scanner', () => {
		const cases = [
			'Sub T()\n    ws.Ran|\nEnd Sub\n',
			'Dim x As Stri|',
			'x = Foo |',
			'y = 5|',
			'|',
		];
		for (const marked of cases) {
			const offset = marked.indexOf('|');
			const src = marked.replace('|', '');
			const ctx = completionCursorContext(src, offset);
			const span = identifierSpanEndingAt(src, offset);
			if (ctx.partialToken) {
				expect(span).toEqual({ start: ctx.partialToken.start, end: ctx.partialToken.end });
			} else if (span) {
				// The char-level scanner cannot see token kinds; it may report a
				// span for non-identifier tokens (e.g. the digits of a literal).
				expect(src.slice(span.start, span.end)).toMatch(/^[A-Za-z_][A-Za-z0-9_]*$/);
			}
		}
	});

	it('clamps offsets outside the source', () => {
		expect(completionCursorContext('x = 1', -3).tokens).toHaveLength(0);
		expect(completionCursorContext('x = 1', 99).offset).toBe(5);
	});
});
