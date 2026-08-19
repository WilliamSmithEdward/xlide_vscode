import { describe, it, expect } from 'vitest';
import {
	completionCursorContext,
	identifierSpanEndingAt,
	spaceTriggerMayComplete,
	tokenize,
} from '../src/analyzer';

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

	it('reuses one tokenize pass for repeated lookups at the same position', () => {
		const src = 'Sub T()\n    ws.Ran\nEnd Sub\n';
		const offset = src.indexOf('Ran') + 3;
		expect(completionCursorContext(src, offset)).toBe(completionCursorContext(src, offset));
		expect(completionCursorContext(src, offset)).not.toBe(completionCursorContext(src, offset - 1));
	});
});

describe('spaceTriggerMayComplete', () => {
	it('keeps the grammar positions that space completion powers', () => {
		// Declaration / New type positions.
		expect(spaceTriggerMayComplete('    Dim x As ')).toBe(true);
		expect(spaceTriggerMayComplete('    Set wb = New ')).toBe(true);
		// Keyword grammar: End/Exit/Option/On Error/Then/To and friends.
		expect(spaceTriggerMayComplete('End ')).toBe(true);
		expect(spaceTriggerMayComplete('    Exit ')).toBe(true);
		expect(spaceTriggerMayComplete('Option ')).toBe(true);
		expect(spaceTriggerMayComplete('    On Error ')).toBe(true);
		expect(spaceTriggerMayComplete('    If x > 1 ')).toBe(true);
		expect(spaceTriggerMayComplete('    For i = 1 ')).toBe(true);
		expect(spaceTriggerMayComplete('    For Each item ')).toBe(true);
		// Procedure labels.
		expect(spaceTriggerMayComplete('    GoTo ')).toBe(true);
		expect(spaceTriggerMayComplete('    Resume ')).toBe(true);
		// Declaration headers (parameter/return As positions, event stubs).
		expect(spaceTriggerMayComplete('Private Sub ')).toBe(true);
		expect(spaceTriggerMayComplete('Function F(ByVal x As Long) ')).toBe(true);
		// Statement starts and member dots.
		expect(spaceTriggerMayComplete('    ')).toBe(true);
		expect(spaceTriggerMayComplete('')).toBe(true);
		expect(spaceTriggerMayComplete('    ws. ')).toBe(true);
		expect(spaceTriggerMayComplete('x = 1: ')).toBe(true);
		expect(spaceTriggerMayComplete('Handler: Resume ')).toBe(true);
		expect(spaceTriggerMayComplete('10 GoTo ')).toBe(true);
		expect(spaceTriggerMayComplete('#If VBA7 ')).toBe(true);
		// Statements continued from the previous physical line.
		expect(spaceTriggerMayComplete('    y ', true)).toBe(true);
	});

	it('bails on ordinary code, comments, and strings', () => {
		expect(spaceTriggerMayComplete('    x = y ')).toBe(false);
		expect(spaceTriggerMayComplete('    MySub arg1, ')).toBe(false);
		expect(spaceTriggerMayComplete('    ws.Range("A1").Select ')).toBe(false);
		expect(spaceTriggerMayComplete("    x = 1 ' a comment ")).toBe(false);
		expect(spaceTriggerMayComplete('    Rem old-style comment ')).toBe(false);
		expect(spaceTriggerMayComplete('    x = "inside a string ')).toBe(false);
		expect(spaceTriggerMayComplete('    .Cells(1, 1).Value = x ')).toBe(false);
	});
});

describe('prefix token derivation from the cached module stream', () => {
	// The cursor context now slices the memoized full-module token stream
	// instead of re-lexing the prefix per keystroke. This must be
	// indistinguishable from lexing the truncated prefix, at EVERY offset:
	// cuts inside strings, comments, identifiers, multi-char operators, line
	// continuations, and unicode identifiers all included.
	const FIXTURE = [
		'Attribute VB_Name = "Módulo"',
		'Option Explicit',
		'',
		"' A comment with \"quotes\" and trailing spaces   ",
		'Public Sub Пример()',
		"    Dim s As String: s = \"a \"\"quoted\"\" string with ' inside\"",
		'    If x <= 42 And y >= 3.14 Then',
		'        total = total + _',
		'            offset(1, 2)',
		'    End If',
		'    Rem old-style comment, with commas',
		'End Sub',
		'',
	].join('\r\n');

	function tokenSnapshot(tokens: readonly { kind: string; rawText: string; start: number; end: number }[]) {
		return tokens.map((t) => `${t.kind}:${t.start}:${t.end}:${t.rawText}`);
	}

	it('matches a fresh prefix lex at every offset of the fixture', () => {
		for (let offset = 0; offset <= FIXTURE.length; offset++) {
			const derived = completionCursorContext(FIXTURE, offset);
			const direct = tokenize(FIXTURE.slice(0, offset));
			expect(tokenSnapshot(derived.tokens), `offset ${offset}`).toEqual(tokenSnapshot(direct));
		}
	});
});
