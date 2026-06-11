// Rule family: lexical source-shape rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: unterminated string literals and
// invalid line continuations.

import { tokenizeCached } from '../../lexer/tokenize';
import type { PushFn } from '../analysisContext';

/** Counts double-quote characters; an odd count means the string is unterminated. */
function countQuotes(text: string): number {
	let n = 0;
	for (const ch of text) {
		if (ch === '"') {
			n++;
		}
	}
	return n;
}

/** Rule: a string literal with an odd number of quotes is never closed. */
export function checkUnterminatedStrings(source: string, push: PushFn): void {
	for (const tok of tokenizeCached(source)) {
		if (tok.kind === 'stringLiteral' && countQuotes(tok.rawText) % 2 === 1) {
			push(
				'unterminatedString',
				'Unterminated string literal.',
				{ start: tok.start, end: tok.end },
			);
		}
	}
}

/**
 * Rule: VBA line-continuation trivia is strictly `1*WSC "_" line-terminator`.
 * A likely continuation underscore with trailing text/comment, or without the
 * required whitespace before it, is a settled compile-time syntax error.
 */
export function checkInvalidLineContinuations(source: string, push: PushFn): void {
	let lineStart = 0;
	while (lineStart < source.length) {
		let lineEnd = lineStart;
		while (lineEnd < source.length && source[lineEnd] !== '\r' && source[lineEnd] !== '\n') {
			lineEnd++;
		}
		checkInvalidLineContinuationOnLine(source, lineStart, lineEnd, push);
		if (lineEnd >= source.length) {
			break;
		}
		lineStart = source[lineEnd] === '\r' && source[lineEnd + 1] === '\n'
			? lineEnd + 2
			: lineEnd + 1;
	}
}

function checkInvalidLineContinuationOnLine(
	source: string,
	lineStart: number,
	lineEnd: number,
	push: PushFn,
): void {
	const commentStart = physicalLineCommentStart(source, lineStart, lineEnd) ?? lineEnd;
	const codeLast = lastNonWscOffset(source, lineStart, commentStart);
	if (codeLast === undefined) {
		return;
	}
	const visibleLineEnd = lastNonWscOffset(source, lineStart, lineEnd);
	const spanEnd = visibleLineEnd === undefined ? lineEnd : visibleLineEnd + 1;

	for (const underscore of underscoresOutsideStrings(source, lineStart, commentStart)) {
		const prev = source[underscore - 1];
		const next = source[underscore + 1];
		const prevIsWsc = underscore > lineStart && isVbaWsc(prev);
		const nextStartsIdentifier = next !== undefined && isIdentifierPartChar(next);
		const hasTrailingText = firstNonWscOffset(source, underscore + 1, lineEnd) !== undefined;

		if (prevIsWsc && hasTrailingText && !nextStartsIdentifier) {
			push(
				'invalidLineContinuation',
				"Line continuation '_' must be the final non-whitespace character on the physical line.",
				{ start: underscore, end: Math.max(underscore + 1, spanEnd) },
			);
			return;
		}

		if (
			underscore === codeLast &&
			lineEnd < source.length &&
			!prevIsWsc &&
			!isIdentifierPartChar(prev)
		) {
			push(
				'invalidLineContinuation',
				"Line continuation '_' must be preceded by whitespace.",
				{ start: underscore, end: underscore + 1 },
			);
			return;
		}
	}
}

function physicalLineCommentStart(
	source: string,
	lineStart: number,
	lineEnd: number,
): number | undefined {
	let inString = false;
	let statementStart = true;
	for (let i = lineStart; i < lineEnd; i++) {
		const ch = source[i];
		if (inString) {
			if (ch === '"') {
				if (source[i + 1] === '"') {
					i++;
				} else {
					inString = false;
				}
			}
			continue;
		}
		if (ch === '"') {
			inString = true;
			statementStart = false;
			continue;
		}
		if (ch === "'") {
			return i;
		}
		if (isVbaWsc(ch)) {
			continue;
		}
		if (ch === ':') {
			statementStart = true;
			continue;
		}
		if (statementStart && startsRemComment(source, i, lineEnd)) {
			return i;
		}
		statementStart = false;
	}
	return undefined;
}

function underscoresOutsideStrings(
	source: string,
	start: number,
	end: number,
): number[] {
	const offsets: number[] = [];
	let inString = false;
	for (let i = start; i < end; i++) {
		const ch = source[i];
		if (inString) {
			if (ch === '"') {
				if (source[i + 1] === '"') {
					i++;
				} else {
					inString = false;
				}
			}
			continue;
		}
		if (ch === '"') {
			inString = true;
			continue;
		}
		if (ch === '_') {
			offsets.push(i);
		}
	}
	return offsets;
}

function startsRemComment(source: string, offset: number, end: number): boolean {
	return offset + 3 <= end &&
		source.slice(offset, offset + 3).toLowerCase() === 'rem' &&
		!isIdentifierPartChar(source[offset + 3]);
}

function firstNonWscOffset(
	source: string,
	start: number,
	end: number,
): number | undefined {
	for (let i = start; i < end; i++) {
		if (!isVbaWsc(source[i])) {
			return i;
		}
	}
	return undefined;
}

function lastNonWscOffset(
	source: string,
	start: number,
	end: number,
): number | undefined {
	for (let i = end - 1; i >= start; i--) {
		if (!isVbaWsc(source[i])) {
			return i;
		}
	}
	return undefined;
}

function isVbaWsc(ch: string | undefined): boolean {
	return ch === '\t' || ch === '\u0019' || ch === ' ' || ch === '\u3000';
}

function isIdentifierPartChar(ch: string | undefined): boolean {
	return ch !== undefined && /[A-Za-z0-9_]/.test(ch);
}
