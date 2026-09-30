// Rule family: characters and lines the VBE refuses outright (issues #132
// and #133). Measured in Excel 16.0 (build 20326, 2026-09-26):
//
//  - stray-character: `n = 1;` (a semicolon outside a Print or Write list),
//    a backtick, braces, a lone `@`, `~`, `|` -> "Syntax error". A
//    non-breaking space (U+00A0, what a web page or Word document pastes) is
//    not whitespace to the VBE: between tokens it is a Syntax error and at
//    the start of a line it becomes part of the name ("Variable not
//    defined"). The lexer gives all of these the `unknown` kind, and `;` the
//    punctuation kind.
//  - line-too-long: the VBE reads a physical line in pieces of 1023
//    characters, each its own line (issue #187, measured 2026-09-29). A
//    1024-character statement is refused. A line long only because of
//    blanks compiles: 2635 spaces, a statement with 1100 trailing blanks, or
//    one after 1100 leading blanks, since the other pieces are blank lines.
//    So the rule is that the code between the first and last non-blank
//    character fits one piece.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { tokenizeCached } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { PushFn } from '../analysisContext';
import { tokenText } from '../walker';

const MAX_LINE_LENGTH = 1023;

/** Statements whose lists take `;` and `,` as output separators. */
const PRINT_LIKE: ReadonlySet<string> = new Set(['print', 'write', 'debug']);

export function checkStrayCharacters(
	source: string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const tokens = tokenizeCached(source);
	let statementHead: VbaToken | undefined;
	let afterHash = false;
	let previous: VbaToken | undefined;
	for (const tok of tokens) {
		if (tok.kind === 'newline' || tok.kind === 'colon') {
			statementHead = undefined;
			afterHash = false;
			previous = undefined;
			continue;
		}
		if (tok.kind === 'comment') {
			continue;
		}
		const prior = previous;
		previous = tok;
		if (!statementHead) {
			// A line number is not the statement: `10 Debug.Print "a"; "b"`
			// (issue #143). The next token is the head.
			if (tok.kind === 'integerLiteral' && prior === undefined && /^\d+$/.test(tok.rawText)) {
				previous = undefined;
				continue;
			}
			statementHead = tok;
			// `#Const`, `#If` lines are directives; `Print #1, x` names a file.
			afterHash = tok.kind === 'directive';
		} else if (tok.kind === 'keyword' && (tokenText(tok) === 'then' || tokenText(tok) === 'else')) {
			// A one-line If runs a statement after Then and another after Else:
			// `If x Then Debug.Print a; b` is a Print list (issue #143).
			statementHead = undefined;
			continue;
		}
		const span = { start: tok.start, end: tok.end };
		if (tok.kind === 'unknown') {
			// A run of such characters (four NBSPs as an indent) is one finding.
			if (activity?.isInactive(span) || (prior?.kind === 'unknown' && prior.end === tok.start)) {
				continue;
			}
			// `1.5%`: the Integer suffix glued to a fractional literal is the
			// literal rule's finding, not a stray character. A type-declaration
			// character glued to a name (`Left$(`, `n%`, `x@`) is the name's suffix.
			if (tok.rawText === '%' && prior?.kind === 'floatLiteral' && prior.end === tok.start) {
				continue;
			}
			if (/^[$%&!#@]$/.test(tok.rawText) && prior && prior.end === tok.start
				&& (prior.kind === 'identifier' || prior.kind === 'keyword' || prior.kind === 'bracketedIdentifier')) {
				continue;
			}
			const text = tok.rawText;
			if (/^[\u00A0]+$/.test(text)) {
				push(
					'strayCharacter',
					'A non-breaking space (U+00A0) is not whitespace in VBA: the VBE reads it as part of a name or refuses the line. Replace it with an ordinary space. This is a VBE compile error.',
					span,
				);
			} else {
				push(
					'strayCharacter',
					`'${text}' is not a character VBA uses here. This is a VBE compile error: Syntax error.`,
					span,
				);
			}
			continue;
		}
		if (tok.kind === 'punctuation' && tok.rawText === ';' && !afterHash && !activity?.isInactive(span)) {
			const head = statementHead ? tokenText(statementHead) : '';
			if (!PRINT_LIKE.has(head)) {
				push(
					'strayCharacter',
					"A ';' ends no VBA statement: it separates items only in a Print or Write list. Remove it. This is a VBE compile error: Syntax error.",
					span,
				);
			}
		}
	}
	let lineStart = 0;
	let lineIndex = 0;
	while (lineStart <= source.length) {
		let lineEnd = source.indexOf('\n', lineStart);
		if (lineEnd < 0) {
			lineEnd = source.length;
		}
		let visibleEnd = lineEnd;
		if (visibleEnd > lineStart && source[visibleEnd - 1] === '\r') {
			visibleEnd--;
		}
		const length = visibleEnd - lineStart;
		if (length > MAX_LINE_LENGTH) {
			let first = lineStart;
			while (first < visibleEnd && (source[first] === ' ' || source[first] === '\t')) {
				first++;
			}
			let last = visibleEnd - 1;
			while (last >= first && (source[last] === ' ' || source[last] === '\t')) {
				last--;
			}
			const piece = Math.floor((first - lineStart) / MAX_LINE_LENGTH);
			const crosses = first <= last && Math.floor((last - lineStart) / MAX_LINE_LENGTH) !== piece;
			const span = { start: lineStart + (piece + 1) * MAX_LINE_LENGTH, end: last + 1 };
			if (crosses && !activity?.isInactive(span)) {
				push(
					'lineTooLong',
					`Line ${lineIndex + 1} is ${length} characters long. The VBE reads a line in pieces of ${MAX_LINE_LENGTH} characters, and this line's code runs past the end of one. Break it with a line continuation. This is a VBE compile error.`,
					span,
				);
			}
		}
		if (lineEnd >= source.length) {
			break;
		}
		lineStart = lineEnd + 1;
		lineIndex++;
	}
}
