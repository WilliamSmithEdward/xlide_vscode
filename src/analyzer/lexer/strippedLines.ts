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

import { tokenizeCached } from './tokenize';

/**
 * Every physical line of `source` with string-literal and comment token spans
 * blanked to spaces, preserving length and column alignment. A string literal
 * never spans physical lines (MS-VBAL 3.3.4); a comment does when it runs on
 * through ` _` (MS-VBAL 3.3.1), and every line it covers is blanked.
 */
export function lexerStrippedLines(source: string): string[] {
	const chars = source.split(/\r\n|\r|\n/).map((line) => line.split(''));
	for (const token of tokenizeCached(source)) {
		if (token.kind !== 'comment' && token.kind !== 'stringLiteral') {
			continue;
		}
		token.rawText.split(/\r\n|\r|\n/).forEach((segment, index) => {
			const lineChars = chars[token.line + index];
			if (!lineChars) {
				return;
			}
			const from = index === 0 ? token.character : 0;
			const end = Math.min(from + segment.length, lineChars.length);
			for (let col = from; col < end; col++) {
				lineChars[col] = ' ';
			}
		});
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
