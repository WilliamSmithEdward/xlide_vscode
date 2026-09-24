// Leading-trivia scanner for the VBA lexer.
//
// Verified against MS-VBAL.pdf, v20250520, section 3.2.2 (WSC, line-continuation)
// and 3.3 ("Lexical tokens encompass any white space characters that immediately
// precede them.").

import { isLineTerminator, isWsc, Trivia } from './tokenKinds';

/** Result of scanning a run of trivia: the trivia plus the advanced cursor. */
export interface TriviaScan {
	trivia: Trivia[];
	pos: number;
	line: number;
	character: number;
}

/**
 * Consume whitespace and line-continuation trivia starting at `pos`. Stops at
 * the first character that begins a real token (including a line terminator,
 * which is a significant 'newline' token, not trivia).
 *
 * line-continuation = 1*WSC underscore line-terminator (MS-VBAL 3.2.2): a run of
 * whitespace, an underscore, and the line terminator are merged into a single
 * lineContinuation trivia so that the logical line is preserved while the raw
 * text round-trips. The VBE also continues a line whose underscore has
 * whitespace after it, and drops that whitespace when it stores the line, so
 * the trivia takes it in too (issue #83).
 */
export function scanLeadingTrivia(
	src: string,
	pos: number,
	line: number,
	character: number,
): TriviaScan {
	const trivia: Trivia[] = [];
	const len = src.length;
	while (pos < len) {
		const ch = src[pos];
		if (!isWsc(ch)) {
			break;
		}
		const start = pos;
		const startChar = character;
		while (pos < len && isWsc(src[pos])) {
			pos++;
			character++;
		}
		// A line-continuation is whitespace + '_' + line terminator.
		const terminator = continuationTerminator(src, pos);
		if (terminator >= 0) {
			pos = terminator; // consume '_' and any whitespace after it
			// consume the line terminator (CRLF, CR, or LF)
			if (src[pos] === '\r' && pos + 1 < len && src[pos + 1] === '\n') {
				pos += 2;
			} else {
				pos += 1;
			}
			line++;
			character = 0;
			trivia.push({ kind: 'lineContinuation', text: src.slice(start, pos), start, end: pos });
		} else {
			trivia.push({ kind: 'whitespace', text: src.slice(start, pos), start, end: pos });
			// character already advanced; startChar retained for clarity only.
			void startChar;
		}
	}
	return { trivia, pos, line, character };
}

/**
 * Where the line terminator starts when `pos` holds the underscore of a line
 * continuation, whitespace after the underscore allowed; otherwise -1.
 */
export function continuationTerminator(src: string, pos: number): number {
	if (src[pos] !== '_') {
		return -1;
	}
	let at = pos + 1;
	while (at < src.length && isWsc(src[at])) {
		at++;
	}
	return at < src.length && isLineTerminator(src[at]) ? at : -1;
}
