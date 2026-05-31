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
import { isLineTerminator, TokenKind, Trivia, VbaToken } from './tokenKinds';
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
	return isIdentStart(ch);
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
				// Try a '#'-delimited date literal on the same physical line.
				let scan = pos + 1;
				while (scan < len && src[scan] !== '#' && !isLineTerminator(src[scan])) {
					scan++;
				}
				if (scan < len && src[scan] === '#') {
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
