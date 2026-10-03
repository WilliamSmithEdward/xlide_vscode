// Rule: places where the VBE refuses a line continuation (issue #126).
// Measured in Excel 16.0 (build 20326, 2026-09-25):
//
//  - 25 continuations in one logical line: AddFromString refuses the module,
//    "Too many line continuations". 24 compile.
//  - A continuation followed by a blank line: `x = 1 + _` then an empty
//    line. Syntax error.
//  - Any continuation inside an Enum body, on a member line or before
//    `End Enum`: "Invalid inside Enum". The same splits inside a Type compile,
//    and so does one after `Private` on the Enum header line.
//
// A Declare continued between its Lib (or Alias) string and its parameter
// list compiles. It was reported until a recheck on 2026-10-01: the
// "Syntax error" came from VBComponents.CodeModule.AddFromString, which
// stores a stray `()` line after such a Declare. The same module
// imported from a .bas file, saved and reopened, compiles, and so does
// VBA-JSON, which has this shape in its Windows branch.
//
// The analyzer follows continuations everywhere else the VBE does; those
// places are covered by tests/diagnostics/lineContinuations.test.ts.

import { tokenizeCached } from '../../lexer/tokenize';
import { firstTokenAtOrAfter } from '../../lexer/tokenHelpers';
import type { Trivia } from '../../lexer/tokenKinds';
import type { ModuleNode } from '../../parser/nodes';
import type { PushFn } from '../analysisContext';

const MAX_CONTINUATIONS = 24;

export function checkLineContinuationLimits(source: string, mod: ModuleNode, push: PushFn): void {
	const tokens = tokenizeCached(source);
	const continuations: Trivia[] = [];
	let inLogicalLine = 0;
	for (const tok of tokens) {
		for (const trivia of tok.leadingTrivia ?? []) {
			if (trivia.kind !== 'lineContinuation') {
				continue;
			}
			continuations.push(trivia);
			inLogicalLine++;
			if (inLogicalLine === MAX_CONTINUATIONS + 1) {
				push(
					'invalidLineContinuation',
					`This is the ${MAX_CONTINUATIONS + 1}th line continuation in one logical line; the VBE allows ${MAX_CONTINUATIONS} ("Too many line continuations").`,
					{ start: trivia.start, end: trivia.end },
				);
			}
			// The continuation joined this line to the next; a newline token
			// right after it means the next physical line is empty.
			if (tok.kind === 'newline' && onlyWhitespaceBetween(source, trivia.end, tok.start)) {
				push(
					'invalidLineContinuation',
					'A line continuation must be followed by more of the statement; the next line is empty. This is a VBE compile error: Syntax error.',
					{ start: trivia.start, end: trivia.end },
				);
			}
		}
		if (tok.kind === 'newline') {
			inLogicalLine = 0;
		}
	}
	if (continuations.length === 0) {
		return;
	}
	for (const member of mod.members) {
		if (member.kind === 'Enum') {
			const headerEnd = lineEndAfter(source, member.nameSpan?.end ?? member.span.start);
			for (let i = firstTokenAtOrAfter(continuations, headerEnd + 1); i < continuations.length; i++) {
				const trivia = continuations[i];
				if (trivia.start >= member.span.end) { break; }
				push(
					'invalidLineContinuation',
					`A line continuation is not allowed inside Enum '${member.name}': neither on a member line nor before End Enum ("Invalid inside Enum").`,
					{ start: trivia.start, end: trivia.end },
				);
			}
		}
	}
}

function onlyWhitespaceBetween(source: string, from: number, to: number): boolean {
	return /^[ \t]*$/.test(source.slice(from, to));
}

/** Offset of the first line terminator at or after `from`. */
function lineEndAfter(source: string, from: number): number {
	for (let i = from; i < source.length; i++) {
		if (source[i] === '\r' || source[i] === '\n') {
			return i;
		}
	}
	return source.length;
}
