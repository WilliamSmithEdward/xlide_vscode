// Lexer-derived stripped-line substrate (audit #74).
//
// Produces the same per-line contract as the legacy regex scanner
// (vbaSourceScan.stripVba): every physical line with string-literal and
// comment spans blanked to spaces, length and column alignment preserved.
// The spans come from the analyzer lexer (tokenize.ts) instead of an ad-hoc
// regex pass, so a single lexer defines string/comment semantics for the
// consumers built on this substrate (Smart Enter and keyword completion).
//
// Deliberate divergence from stripVba: the lexer recognizes `Rem` comments at
// any statement start (MS-VBAL 3.3.5.2), so a trailing `: Rem ...` comment is
// blanked here but leaks through stripVba, which only blanks whole-line Rem
// comments. tests/smartEnterSubstrateComparison.test.ts diffs the two
// substrates over every VBA sample in the repository and pins that as the
// only allowed difference.

import { tokenize } from './tokenize';

/**
 * Every physical line of `source` with string-literal and comment token spans
 * blanked to spaces, preserving length and column alignment. Comments and
 * string literals never span physical lines (MS-VBAL 3.3.1 / 3.3.4), so
 * blanking on the token's start line covers the whole token.
 */
export function lexerStrippedLines(source: string): string[] {
	const chars = source.split(/\r\n|\r|\n/).map((line) => line.split(''));
	for (const token of tokenize(source)) {
		if (token.kind !== 'comment' && token.kind !== 'stringLiteral') {
			continue;
		}
		const lineChars = chars[token.line];
		if (!lineChars) {
			continue;
		}
		const end = Math.min(token.character + token.rawText.length, lineChars.length);
		for (let col = token.character; col < end; col++) {
			lineChars[col] = ' ';
		}
	}
	return chars.map((lineChars) => lineChars.join(''));
}

/**
 * Single-line variant for call sites that strip one physical line in
 * isolation (the per-line stripVba contract).
 */
export function lexerStrippedLine(line: string): string {
	return lexerStrippedLines(line)[0] ?? line;
}
