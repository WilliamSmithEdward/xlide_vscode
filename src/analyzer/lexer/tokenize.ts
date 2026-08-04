// VBA tokenizer.
//
// Verified against MS-VBAL.pdf, v20250520 (Release: May 20, 2025):
//   - 3.2.2 Logical Line Grammar (WSC, line-continuation, line terminators)
//   - 3.3.1 Separator and Special Tokens (special-token, comments, EOS/colon)
//   - 3.3.2 Number Tokens (INTEGER, FLOAT, type suffixes)
//   - 3.3.3 Date Tokens (date-or-time, '#'-delimited)
//   - 3.3.4 String Tokens (doubled-quote escaping, end-at-LINE-END)
//   - 3.3.5 Identifier Tokens (lex-identifier, FOREIGN-NAME, reserved-identifier)
//   - 3.4   Conditional Compilation (#Const / #If / #ElseIf / #Else / #EndIf)
//
// The lexer is loss-aware and round-trippable: re-joining every token's leading
// trivia + rawText (+ any trailing trivia on the final token) reproduces the
// source exactly. Reserved/contextual keywords carry canonical capitalization.

import { canonicalKeyword } from './keywordTable';
import { isLineTerminator, isWsc, TokenKind, Trivia, VbaToken } from './tokenKinds';
import { scanLeadingTrivia } from './trivia';

const SPECIAL_DECISION = {
	/** Type-suffix chars that make an integer literal (MS-VBAL 3.3.2). */
	integerSuffix: new Set(['%', '&', '^']),
	/** Type-suffix chars that make a float literal (MS-VBAL 3.3.2). */
	floatSuffix: new Set(['!', '#', '@']),
} as const;

function isDigit(ch: string): boolean {
	return ch >= '0' && ch <= '9';
}

function isOctalDigit(ch: string): boolean {
	return ch >= '0' && ch <= '7';
}

function isHexDigit(ch: string): boolean {
	return isDigit(ch) || (ch >= 'a' && ch <= 'f') || (ch >= 'A' && ch <= 'F');
}

function isExponentLetter(ch: string): boolean {
	return ch === 'e' || ch === 'E' || ch === 'd' || ch === 'D';
}

function isIdentStart(ch: string): boolean {
	if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')) {
		return true;
	}
	// Non-Latin identifier forms (MS-VBAL 3.3.5.1): permit Unicode letters.
	return ch.charCodeAt(0) >= 0x80 && /\p{L}/u.test(ch);
}

function isIdentPart(ch: string): boolean {
	if (ch === '_' || isDigit(ch)) {
		return true;
	}
	if (isIdentStart(ch)) {
		return true;
	}
	// Combining marks continue a name but can never begin one. Thai writes a
	// letter as base + tone mark and Devanagari as base + matra - categories
	// Mn/Mc, not L - and the VBE compiles both, so stopping at the mark split
	// valid identifiers in half. At position 0 a mark has nothing to combine
	// with, so isIdentStart deliberately still rejects it.
	return ch.charCodeAt(0) >= 0x80 && /\p{M}/u.test(ch);
}

// Hot editor paths (hover, canonical casing) re-tokenize the same full module
// text several times per request; a value-keyed memo collapses those scans to
// one. Sized to hold a handful of recent modules (matching the parser cache) so
// a workbook pass touching sibling modules does not evict the active one.
// Callers must not mutate the returned array or its tokens.
const TOKENIZE_CACHE_MAX = 8;
const tokenizeCache: { src: string; tokens: VbaToken[] }[] = [];

/** Cached variant of {@link tokenize} for read-only consumers on hot paths. */
export function tokenizeCached(src: string): VbaToken[] {
	for (let i = 0; i < tokenizeCache.length; i += 1) {
		if (tokenizeCache[i].src === src) {
			const hit = tokenizeCache[i];
			if (i > 0) {
				tokenizeCache.splice(i, 1);
				tokenizeCache.unshift(hit);
			}
			return hit.tokens;
		}
	}
	const tokens = tokenize(src);
	tokenizeCache.unshift({ src, tokens });
	if (tokenizeCache.length > TOKENIZE_CACHE_MAX) {
		tokenizeCache.pop();
	}
	return tokens;
}

/**
 * Tokenize a VBA module body into a flat, round-trippable token stream.
 *
 * @param src The raw module source (may contain CR, LF, or CRLF terminators).
 */
export function tokenize(src: string): VbaToken[] {
	const tokens: VbaToken[] = [];
	const len = src.length;
	let pos = 0;
	let line = 0;
	let character = 0;
	// EOS/statement-start tracking: true at file start and after a newline or a
	// ':' separator (MS-VBAL 3.3.1 EOS). Governs Rem-comment and directive lexing.
	let atStatementStart = true;

	while (pos < len) {
		const tr = scanLeadingTrivia(src, pos, line, character);
		const leading: Trivia[] = tr.trivia;
		pos = tr.pos;
		line = tr.line;
		character = tr.character;

		if (pos >= len) {
			// Trailing trivia at EOF: attach to the last token so the stream stays
			// round-trippable. If there are no tokens, the input was trivia-only.
			if (leading.length > 0 && tokens.length > 0) {
				tokens[tokens.length - 1].trailingTrivia = leading;
			}
			break;
		}

		const startPos = pos;
		const startLine = line;
		const startChar = character;
		const ch = src[pos];
		let kind: TokenKind;
		let canonical: string | undefined;
		let isNewline = false;

		if (isLineTerminator(ch)) {
			// newline token (line-terminator). Consume CRLF / CR / LF.
			if (ch === '\r' && pos + 1 < len && src[pos + 1] === '\n') {
				pos += 2;
			} else {
				pos += 1;
			}
			kind = 'newline';
			isNewline = true;
			atStatementStart = true;
		} else if (ch === "'") {
			// Apostrophe comment to end of physical line (MS-VBAL 3.3.1). The VBE
			// does not continue comments across line-continuations, so we stop at
			// the line terminator.
			pos++;
			while (pos < len && !isLineTerminator(src[pos])) {
				pos++;
			}
			kind = 'comment';
		} else if (isIdentStart(ch)) {
			pos++;
			while (pos < len && isIdentPart(src[pos])) {
				pos++;
			}
			const word = src.slice(startPos, pos);
			if (word.toLowerCase() === 'rem' && atStatementStart) {
				// Rem comment (MS-VBAL 3.3.5.2 rem-keyword): rest of line is comment.
				while (pos < len && !isLineTerminator(src[pos])) {
					pos++;
				}
				kind = 'comment';
			} else {
				canonical = canonicalKeyword(word);
				kind = canonical ? 'keyword' : 'identifier';
				atStatementStart = false;
			}
		} else if (
			isDigit(ch) ||
			(ch === '.' && pos + 1 < len && isDigit(src[pos + 1])) ||
			(ch === '&' &&
				pos + 1 < len &&
				(src[pos + 1] === 'h' || src[pos + 1] === 'H' || src[pos + 1] === 'o' || src[pos + 1] === 'O'))
		) {
			kind = lexNumber(src, () => pos, (p) => (pos = p));
			atStatementStart = false;
		} else if (ch === '"') {
			// String literal (MS-VBAL 3.3.4): doubled-quote escaping; may end at
			// the closing quote or at LINE-END (unterminated tolerated).
			pos++;
			while (pos < len) {
				const c = src[pos];
				if (c === '"') {
					if (pos + 1 < len && src[pos + 1] === '"') {
						pos += 2; // escaped quote
						continue;
					}
					pos++; // closing quote
					break;
				}
				if (isLineTerminator(c)) {
					break; // unterminated string ends at LINE-END
				}
				pos++;
			}
			kind = 'stringLiteral';
			atStatementStart = false;
		} else if (ch === '[') {
			// FOREIGN-NAME (MS-VBAL 3.3.5.3): "[" 1*non-line-termination-character "]".
			pos++;
			while (pos < len && src[pos] !== ']' && !isLineTerminator(src[pos])) {
				pos++;
			}
			if (pos < len && src[pos] === ']') {
				pos++;
			}
			kind = 'bracketedIdentifier';
			atStatementStart = false;
		} else if (ch === '#') {
			if (atStatementStart) {
				// Conditional-compilation directive marker (MS-VBAL 3.4).
				pos++;
				kind = 'directive';
				atStatementStart = false;
			} else {
				// Candidate '#'-delimited date literal on the same physical line. The
				// '#' pair only forms a DATE token when the enclosed text is a valid
				// date-or-time body (MS-VBAL 3.3.3); otherwise this '#' is a
				// file-number marker (`Write #ff, ...`, MS-VBAL 5.4.5 file statements)
				// or a stray type-suffix and lexes as an operator. Without the body
				// check, a file-number '#' would pair with any later '#' on the line
				// and swallow the intervening tokens as one bogus date literal.
				let scan = pos + 1;
				while (scan < len && src[scan] !== '#' && !isLineTerminator(src[scan])) {
					scan++;
				}
				if (scan < len && src[scan] === '#' && isDateLiteralBody(src.slice(pos + 1, scan))) {
					pos = scan + 1;
					kind = 'dateLiteral';
				} else {
					pos++;
					kind = 'operator';
				}
				atStatementStart = false;
			}
		} else {
			kind = lexSymbol(src, ch, () => pos, (p) => (pos = p));
			atStatementStart = kind === 'colon';
		}

		const rawText = src.slice(startPos, pos);
		const token: VbaToken = {
			kind,
			rawText,
			start: startPos,
			end: pos,
			line: startLine,
			character: startChar,
		};
		if (canonical !== undefined) {
			token.canonicalText = canonical;
		}
		if (leading.length > 0) {
			token.leadingTrivia = leading;
		}
		tokens.push(token);

		if (isNewline) {
			line++;
			character = 0;
		} else {
			character += pos - startPos;
		}
	}

	return tokens;
}

/**
 * Lex a numeric literal starting at the current position. Returns the token kind
 * ('integerLiteral' or 'floatLiteral') and advances the cursor via `setPos`.
 * MS-VBAL 3.3.2.
 */
function lexNumber(src: string, getPos: () => number, setPos: (p: number) => void): TokenKind {
	const len = src.length;
	let p = getPos();
	const ch = src[p];

	if (ch === '&') {
		// Hex (&H) or octal (&O) integer literal.
		const radixCh = src[p + 1];
		p += 2; // consume '&' and the radix letter
		if (radixCh === 'h' || radixCh === 'H') {
			while (p < len && isHexDigit(src[p])) {
				p++;
			}
		} else {
			while (p < len && isOctalDigit(src[p])) {
				p++;
			}
		}
		// Optional integer type-suffix (MS-VBAL 3.3.2): % & ^.
		if (p < len && SPECIAL_DECISION.integerSuffix.has(src[p])) {
			p++;
		}
		setPos(p);
		return 'integerLiteral';
	}

	let isFloat = false;
	// integer-digits
	while (p < len && isDigit(src[p])) {
		p++;
	}
	// Optional decimal point. Only consume '.' as part of the number when it is
	// followed by a digit or an exponent letter; otherwise leave it as a member-
	// access dot (a numeric literal cannot have a member, MS-VBAL 3.3.2).
	if (p < len && src[p] === '.') {
		const after = p + 1 < len ? src[p + 1] : '';
		if (isDigit(after) || (isExponentLetter(after) && hasExponentTail(src, p + 1))) {
			isFloat = true;
			p++; // consume '.'
			while (p < len && isDigit(src[p])) {
				p++;
			}
		}
	}
	// Optional exponent.
	if (p < len && isExponentLetter(src[p]) && hasExponentTail(src, p)) {
		isFloat = true;
		p++; // exponent-letter
		if (p < len && (src[p] === '+' || src[p] === '-')) {
			p++;
		}
		while (p < len && isDigit(src[p])) {
			p++;
		}
	}
	// Type suffix.
	if (p < len) {
		if (SPECIAL_DECISION.floatSuffix.has(src[p])) {
			isFloat = true;
			p++;
		} else if (!isFloat && SPECIAL_DECISION.integerSuffix.has(src[p])) {
			p++;
		}
	}
	setPos(p);
	return isFloat ? 'floatLiteral' : 'integerLiteral';
}

/** True if `pos` begins a valid exponent tail: [DdEe] [sign] 1*digit. */
function hasExponentTail(src: string, pos: number): boolean {
	const len = src.length;
	if (pos >= len || !isExponentLetter(src[pos])) {
		return false;
	}
	let p = pos + 1;
	if (p < len && (src[p] === '+' || src[p] === '-')) {
		p++;
	}
	return p < len && isDigit(src[p]);
}

// ---------------------------------------------------------------------------
// Date-literal body validation (MS-VBAL 3.3.3 Date Tokens).
//
//   DATE          = "#" *WSC [date-or-time *WSC] "#"
//   date-or-time  = (date-value 1*WSC time-value) / date-value / time-value
//   date-value    = left-date-value date-separator middle-date-value
//                   [date-separator right-date-value]
//   left/middle/right-date-value = decimal-literal / month-name
//   date-separator = 1*WSC / (*WSC ("/" / "-" / ",") *WSC)
//   time-value    = (hour-value ampm) / (hour-value time-separator minute-value
//                   [time-separator second-value] [ampm])
//   time-separator = *WSC (":" / ".") *WSC
//   ampm          = *WSC ("am" / "pm" / "a" / "p")
//
// Whitespace alone is a date-separator, so the grammar is ambiguous (in
// `#1/1 3:30#` the `3` could be a right-date-value or an hour-value); the
// matchers below return every candidate end position and the caller accepts
// the body if any reading consumes it exactly.

/** month-name = English-month-name / English-month-abbreviation (3.3.3). */
const MONTH_NAMES = new Set([
	'january', 'february', 'march', 'april', 'may', 'june', 'july', 'august',
	'september', 'october', 'november', 'december',
	'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec',
]);

/**
 * True when the text between a '#' pair is a valid date-literal body:
 * `*WSC [date-or-time *WSC]`. A non-matching body means the opening '#' was
 * never a date literal (e.g. a file-number marker) and must lex on its own.
 */
function isDateLiteralBody(body: string): boolean {
	const start = skipWsc(body, 0);
	if (start === body.length) {
		return true; // "#" *WSC "#": the empty date literal
	}
	for (const end of dateOrTimeEnds(body, start)) {
		if (skipWsc(body, end) === body.length) {
			return true;
		}
	}
	return false;
}

/** Candidate end positions of date-or-time at `pos`. */
function dateOrTimeEnds(s: string, pos: number): number[] {
	const ends: number[] = [];
	for (const dateEnd of dateValueEnds(s, pos)) {
		ends.push(dateEnd); // date-value alone
		const wsEnd = skipWsc(s, dateEnd);
		if (wsEnd > dateEnd) {
			ends.push(...timeValueEnds(s, wsEnd)); // date-value 1*WSC time-value
		}
	}
	ends.push(...timeValueEnds(s, pos)); // time-value alone
	return ends;
}

/** Candidate end positions of date-value at `pos` (two- and three-part). */
function dateValueEnds(s: string, pos: number): number[] {
	const ends: number[] = [];
	const left = datePartEnd(s, pos);
	if (left < 0) {
		return ends;
	}
	const sep1 = dateSeparatorEnd(s, left);
	if (sep1 < 0) {
		return ends;
	}
	const middle = datePartEnd(s, sep1);
	if (middle < 0) {
		return ends;
	}
	ends.push(middle);
	const sep2 = dateSeparatorEnd(s, middle);
	if (sep2 >= 0) {
		const right = datePartEnd(s, sep2);
		if (right >= 0) {
			ends.push(right);
		}
	}
	return ends;
}

/** End of a date part (decimal-literal / month-name) at `pos`, or -1. */
function datePartEnd(s: string, pos: number): number {
	const digitsEnd = decimalEnd(s, pos);
	if (digitsEnd >= 0) {
		return digitsEnd;
	}
	let p = pos;
	while (p < s.length && ((s[p] >= 'A' && s[p] <= 'Z') || (s[p] >= 'a' && s[p] <= 'z'))) {
		p++;
	}
	return p > pos && MONTH_NAMES.has(s.slice(pos, p).toLowerCase()) ? p : -1;
}

/**
 * End of a date-separator at `pos`, or -1. Greedily takes the separator-char
 * form when one follows the whitespace; a date part never starts with one of
 * those chars, so the whitespace-only reading could not succeed where this
 * fails.
 */
function dateSeparatorEnd(s: string, pos: number): number {
	const afterWs = skipWsc(s, pos);
	const ch = afterWs < s.length ? s[afterWs] : '';
	if (ch === '/' || ch === '-' || ch === ',') {
		return skipWsc(s, afterWs + 1);
	}
	return afterWs > pos ? afterWs : -1;
}

/** Candidate end positions of time-value at `pos`. */
function timeValueEnds(s: string, pos: number): number[] {
	const ends: number[] = [];
	const hour = decimalEnd(s, pos);
	if (hour < 0) {
		return ends;
	}
	const hourAmPm = ampmEnd(s, hour);
	if (hourAmPm >= 0) {
		ends.push(hourAmPm); // hour-value ampm
	}
	const sep1 = timeSeparatorEnd(s, hour);
	if (sep1 < 0) {
		return ends;
	}
	const minute = decimalEnd(s, sep1);
	if (minute < 0) {
		return ends;
	}
	ends.push(minute);
	const minuteAmPm = ampmEnd(s, minute);
	if (minuteAmPm >= 0) {
		ends.push(minuteAmPm);
	}
	const sep2 = timeSeparatorEnd(s, minute);
	if (sep2 >= 0) {
		const second = decimalEnd(s, sep2);
		if (second >= 0) {
			ends.push(second);
			const secondAmPm = ampmEnd(s, second);
			if (secondAmPm >= 0) {
				ends.push(secondAmPm);
			}
		}
	}
	return ends;
}

/** End of a time-separator (*WSC (":" / ".") *WSC) at `pos`, or -1. */
function timeSeparatorEnd(s: string, pos: number): number {
	const afterWs = skipWsc(s, pos);
	const ch = afterWs < s.length ? s[afterWs] : '';
	return ch === ':' || ch === '.' ? skipWsc(s, afterWs + 1) : -1;
}

/** End of an ampm marker (*WSC ("am" / "pm" / "a" / "p")) at `pos`, or -1. */
function ampmEnd(s: string, pos: number): number {
	const p = skipWsc(s, pos);
	const first = p < s.length ? s[p].toLowerCase() : '';
	if (first !== 'a' && first !== 'p') {
		return -1;
	}
	const second = p + 1 < s.length ? s[p + 1].toLowerCase() : '';
	return second === 'm' ? p + 2 : p + 1;
}

/** End of a decimal-literal (1*DIGIT) at `pos`, or -1. */
function decimalEnd(s: string, pos: number): number {
	let p = pos;
	while (p < s.length && isDigit(s[p])) {
		p++;
	}
	return p > pos ? p : -1;
}

/** First non-WSC position at or after `pos`. */
function skipWsc(s: string, pos: number): number {
	let p = pos;
	while (p < s.length && isWsc(s[p])) {
		p++;
	}
	return p;
}

/**
 * Lex an operator, punctuation, or colon token starting at the current position.
 * MS-VBAL 3.3.1 special-token. Handles the multi-character operators :=, <=, >=,
 * and <>.
 */
function lexSymbol(
	src: string,
	ch: string,
	getPos: () => number,
	setPos: (p: number) => void,
): TokenKind {
	const len = src.length;
	let p = getPos();
	const next = p + 1 < len ? src[p + 1] : '';

	// Multi-character operators.
	if ((ch === ':' && next === '=') || (ch === '<' && (next === '=' || next === '>')) || (ch === '>' && next === '=')) {
		setPos(p + 2);
		return 'operator';
	}

	p++;
	setPos(p);
	switch (ch) {
		case ':':
			return 'colon';
		case ',':
		case '.':
		case '(':
		case ')':
		case ';':
			return 'punctuation';
		case '=':
		case '<':
		case '>':
		case '+':
		case '-':
		case '*':
		case '/':
		case '\\':
		case '^':
		case '&':
		case '!':
		case '?':
			return 'operator';
		default:
			return 'unknown';
	}
}
