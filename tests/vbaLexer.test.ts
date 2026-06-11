import { describe, expect, it } from 'vitest';
import { tokenize } from '../src/analyzer/lexer/tokenize';
import { TokenKind, VbaToken } from '../src/analyzer/lexer/tokenKinds';

/** Reconstruct the source from a token stream to prove round-trippability. */
function roundTrip(tokens: VbaToken[]): string {
	let out = '';
	for (const t of tokens) {
		for (const tv of t.leadingTrivia ?? []) {
			out += tv.text;
		}
		out += t.rawText;
		for (const tv of t.trailingTrivia ?? []) {
			out += tv.text;
		}
	}
	return out;
}

/** Token kinds excluding trivia-only details, for compact assertions. */
function kinds(tokens: VbaToken[]): TokenKind[] {
	return tokens.map((t) => t.kind);
}

function raws(tokens: VbaToken[]): string[] {
	return tokens.map((t) => t.rawText);
}

describe('tokenize - round-trip (Phase 1 acceptance)', () => {
	const samples = [
		'Dim x As Long',
		'  Option   Explicit  ',
		'x = 1 \' trailing comment',
		'If a > 0 Then b = 1 Else b = 2',
		'Dim x _\r\n    As Long',
		's = "a""b""c"',
		'd = #1/2/2020#',
		'n = &HFF& : m = 1.5e3#',
		'Sub Foo()\r\nEnd Sub\r\n',
		'\r\n\r\n   \r\n',
		'Set r = [My Named Range]',
	];

	for (const src of samples) {
		it(`reproduces source exactly: ${JSON.stringify(src)}`, () => {
			expect(roundTrip(tokenize(src))).toBe(src);
		});
	}
});

describe('tokenize - keywords and identifiers', () => {
	it('classifies a Dim statement', () => {
		const t = tokenize('Dim x As Long');
		expect(kinds(t)).toEqual(['keyword', 'identifier', 'keyword', 'keyword']);
		expect(raws(t)).toEqual(['Dim', 'x', 'As', 'Long']);
	});

	it('attaches canonical casing to lowercase keywords', () => {
		const t = tokenize('dim x as long');
		expect(t[0].canonicalText).toBe('Dim');
		expect(t[2].canonicalText).toBe('As');
		expect(t[3].canonicalText).toBe('Long');
		// the identifier carries no canonical casing
		expect(t[1].canonicalText).toBeUndefined();
	});

	it('does not normalize non-keyword identifiers', () => {
		const t = tokenize('myWorksheet.Range');
		expect(t[0].kind).toBe('identifier');
		expect(t[0].canonicalText).toBeUndefined();
	});

	it('lexes non-Latin identifiers without keyword canonicalization', () => {
		const t = tokenize('Dim 価格 As Long\n価格 = 1');
		const names = t.filter((token) => token.rawText === '価格');

		expect(names).toHaveLength(2);
		expect(names.every((token) => token.kind === 'identifier')).toBe(true);
		expect(names.every((token) => token.canonicalText === undefined)).toBe(true);
	});

	it('does not keyword-case reserved-for-implementation-use attribute names', () => {
		const t = tokenize('Attribute VB_Name = "Module1"');

		expect(kinds(t)).toEqual(['identifier', 'identifier', 'operator', 'stringLiteral']);
		expect(raws(t).slice(0, 2)).toEqual(['Attribute', 'VB_Name']);
		expect(t[0].canonicalText).toBeUndefined();
		expect(t[1].canonicalText).toBeUndefined();
	});
});

describe('tokenize - comments', () => {
	it('lexes an apostrophe comment to end of line', () => {
		const t = tokenize('x = 1 \' Sub Function Dim');
		const last = t[t.length - 1];
		expect(last.kind).toBe('comment');
		expect(last.rawText).toBe("' Sub Function Dim");
		// keywords inside the comment are never normalized
		expect(last.canonicalText).toBeUndefined();
	});

	it('lexes a Rem comment at statement start', () => {
		const t = tokenize('Rem this is a remark');
		expect(t).toHaveLength(1);
		expect(t[0].kind).toBe('comment');
		expect(t[0].rawText).toBe('Rem this is a remark');
	});

	it('treats Rem after code as a keyword token, not a comment', () => {
		// Rem is a reserved identifier (rem-keyword); only at statement start does
		// it begin a comment. Mid-statement it is just a keyword token.
		const t = tokenize('x = Rem');
		expect(t[t.length - 1].kind).toBe('keyword');
		expect(t[t.length - 1].rawText).toBe('Rem');
	});

	it('treats Remark as an identifier (not a Rem comment)', () => {
		const t = tokenize('Remark = 1');
		expect(t[0].kind).toBe('identifier');
		expect(t[0].rawText).toBe('Remark');
	});
});

describe('tokenize - strings', () => {
	it('handles doubled-quote escaping as one token', () => {
		const t = tokenize('s = "a""b"');
		const str = t[t.length - 1];
		expect(str.kind).toBe('stringLiteral');
		expect(str.rawText).toBe('"a""b"');
	});

	it('tolerates an unterminated string ending at line end', () => {
		const t = tokenize('s = "abc\r\n');
		const str = t.find((x) => x.kind === 'stringLiteral');
		expect(str?.rawText).toBe('"abc');
		// the newline is a separate token, not swallowed by the string
		expect(t.some((x) => x.kind === 'newline')).toBe(true);
	});
});

describe('tokenize - numbers (MS-VBAL 3.3.2)', () => {
	it('lexes a decimal integer', () => {
		expect(tokenize('123')[0].kind).toBe('integerLiteral');
	});

	it('lexes a hex integer with type suffix', () => {
		const t = tokenize('&HFF&');
		expect(t[0].kind).toBe('integerLiteral');
		expect(t[0].rawText).toBe('&HFF&');
	});

	it('lexes an octal integer', () => {
		const t = tokenize('&O17');
		expect(t[0].kind).toBe('integerLiteral');
		expect(t[0].rawText).toBe('&O17');
	});

	it('lexes floats with fraction, exponent, and suffix', () => {
		expect(tokenize('1.5')[0].kind).toBe('floatLiteral');
		expect(tokenize('1.5e3')[0].kind).toBe('floatLiteral');
		expect(tokenize('.5')[0].kind).toBe('floatLiteral');
		expect(tokenize('1.5#')[0].kind).toBe('floatLiteral');
		expect(tokenize('5!')[0].kind).toBe('floatLiteral');
	});

	it('treats integer type suffixes as integer literals', () => {
		expect(tokenize('100&')[0].kind).toBe('integerLiteral');
		expect(tokenize('100%')[0].kind).toBe('integerLiteral');
	});

	it('does not eat a member-access dot as a decimal point', () => {
		const t = tokenize('a.b');
		expect(kinds(t)).toEqual(['identifier', 'punctuation', 'identifier']);
	});
});

describe('tokenize - date literals (MS-VBAL 3.3.3)', () => {
	it('lexes a hash-delimited date literal', () => {
		const t = tokenize('d = #1/2/2020#');
		const date = t[t.length - 1];
		expect(date.kind).toBe('dateLiteral');
		expect(date.rawText).toBe('#1/2/2020#');
	});

	it('lexes time and month-name date-or-time bodies', () => {
		expect(tokenize('t = #3:15:30 PM#')[2].kind).toBe('dateLiteral');
		expect(tokenize('t = #12:30#')[2].kind).toBe('dateLiteral');
		expect(tokenize('d = #March 15, 2026#')[2].kind).toBe('dateLiteral');
		expect(tokenize('d = #1/1/2026 3:30 am#')[2].kind).toBe('dateLiteral');
	});

	it('lexes a file-number # as an operator, not a date-literal opener', () => {
		// The '#' pair encloses `ff, "csv", 1, ` -- not a date-or-time body, so
		// the file-number '#' must not swallow the statement (file-I/O syntax,
		// e.g. Open/Close/Print/Write/Get/Put, MS-VBAL 5.4.5).
		const kinds = tokenize('Write #ff, "csv", 1, #1/1/2026#').map((t) => `${t.kind}:${t.rawText}`);
		expect(kinds).toEqual([
			'keyword:Write',
			'operator:#',
			'identifier:ff',
			'punctuation:,',
			'stringLiteral:"csv"',
			'punctuation:,',
			'integerLiteral:1',
			'punctuation:,',
			'dateLiteral:#1/1/2026#',
		]);
	});

	it('does not pair a type-suffix # with a later file-number #', () => {
		const kinds = tokenize('x# = 1: Print #1, "y"').map((t) => `${t.kind}:${t.rawText}`);
		expect(kinds).not.toContain('dateLiteral:# = 1: Print #');
		expect(kinds).toContain('stringLiteral:"y"');
	});
});

describe('tokenize - directives (MS-VBAL 3.4)', () => {
	it('marks a leading # as a directive', () => {
		const t = tokenize('#If DEBUGGING Then');
		expect(t[0].kind).toBe('directive');
		expect(t[0].rawText).toBe('#');
		expect(t[1].kind).toBe('keyword');
		expect(t[1].rawText).toBe('If');
	});
});

describe('tokenize - operators and separators (MS-VBAL 3.3.1)', () => {
	it('lexes multi-character operators', () => {
		expect(tokenize('a <= b')[1].rawText).toBe('<=');
		expect(tokenize('a <> b')[1].rawText).toBe('<>');
		expect(tokenize('a >= b')[1].rawText).toBe('>=');
		expect(tokenize('Foo x:=1')[2].rawText).toBe(':=');
	});

	it('treats a bare colon as a statement separator and resets statement start', () => {
		const t = tokenize('x = 1 : Rem hi');
		const colon = t.find((x) => x.kind === 'colon');
		expect(colon).toBeDefined();
		// Rem is recognized as a comment because the colon restarts the statement
		expect(t[t.length - 1].kind).toBe('comment');
	});
});

describe('tokenize - bracketed identifiers (MS-VBAL 3.3.5.3)', () => {
	it('lexes a foreign name with spaces', () => {
		const t = tokenize('Set r = [My Named Range]');
		const bracket = t[t.length - 1];
		expect(bracket.kind).toBe('bracketedIdentifier');
		expect(bracket.rawText).toBe('[My Named Range]');
	});
});

describe('tokenize - line continuation (MS-VBAL 3.2.2)', () => {
	it('captures a line continuation as leading trivia of the next token', () => {
		const t = tokenize('Dim x _\r\n    As Long');
		const asTok = t.find((x) => x.rawText === 'As');
		expect(asTok).toBeDefined();
		const hasContinuation = (asTok?.leadingTrivia ?? []).some((tv) => tv.kind === 'lineContinuation');
		expect(hasContinuation).toBe(true);
	});

	it('keeps the lexer on one logical line across a continuation', () => {
		const t = tokenize('Dim x _\r\n    As Long');
		// no newline token appears because the only line break is a continuation
		expect(t.some((x) => x.kind === 'newline')).toBe(false);
	});
});

describe('tokenize - positions', () => {
	it('reports zero-based line and character for each token', () => {
		const t = tokenize('Dim x\r\nSet y = 1');
		const setTok = t.find((x) => x.rawText === 'Set');
		expect(setTok?.line).toBe(1);
		expect(setTok?.character).toBe(0);
	});
});
