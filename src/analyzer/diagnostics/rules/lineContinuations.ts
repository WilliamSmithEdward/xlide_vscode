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
//  - In a Declare, a continuation between the Lib (or Alias) string and the
//    parameter list. Syntax error. After PtrSafe, after the name, after
//    Private and inside the parameter list all compile.
//
// The analyzer follows continuations everywhere else the VBE does; those
// places are covered by tests/diagnostics/lineContinuations.test.ts.

import { tokenizeCached } from '../../lexer/tokenize';
import type { Trivia, VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode, Span } from '../../parser/nodes';
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
			for (const trivia of continuations) {
				if (trivia.start > headerEnd && trivia.start < member.span.end) {
					push(
						'invalidLineContinuation',
						`A line continuation is not allowed inside Enum '${member.name}': neither on a member line nor before End Enum ("Invalid inside Enum").`,
						{ start: trivia.start, end: trivia.end },
					);
				}
			}
		} else if (member.kind === 'Declare') {
			const gap = declareLibToParamsGap(tokens, member.span);
			if (!gap) {
				continue;
			}
			for (const trivia of continuations) {
				if (trivia.start >= gap.start && trivia.end <= gap.end) {
					push(
						'invalidLineContinuation',
						`Declare '${member.name}' cannot break the line between its Lib or Alias string and the parameter list. This is a VBE compile error: Syntax error.`,
						{ start: trivia.start, end: trivia.end },
					);
				}
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

/** The source between a Declare's last Lib/Alias string literal and its `(`. */
function declareLibToParamsGap(tokens: readonly VbaToken[], span: Span): Span | undefined {
	let lastString: VbaToken | undefined;
	for (const tok of tokens) {
		if (tok.start < span.start) {
			continue;
		}
		if (tok.start >= span.end) {
			break;
		}
		if (tok.kind === 'stringLiteral') {
			lastString = tok;
		} else if (tok.rawText === '(' && lastString) {
			return { start: lastString.end, end: tok.start };
		}
	}
	return undefined;
}
