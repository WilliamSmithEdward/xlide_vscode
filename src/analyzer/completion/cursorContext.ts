// Shared cursor-context detection for the completion stack.
//
// Every completion-adjacent resolver needs the same prefix analysis: tokenize
// the text before the cursor, peel off the trailing partial identifier being
// typed, classify the token that precedes it, and know whether the cursor sits
// inside a comment or string literal. Keeping that dance here (one tokenize
// pass per cursor position) stops the per-resolver reimplementations from
// drifting and gives the resolvers a single seam for sharing token state.

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { isIdentLike } from '../lexer/tokenHelpers';

export interface CompletionCursorContext {
	/** Cursor offset clamped into the source. */
	offset: number;
	/** Raw token stream of the text before the cursor, comments included. */
	tokens: VbaToken[];
	/** Prefix tokens without comments; newlines kept as statement boundaries. */
	significantTokens: VbaToken[];
	/** Identifier-like token being typed when it ends exactly at the cursor. */
	partialToken?: VbaToken;
	/** Text of `partialToken` ('' when the cursor is not finishing a word). */
	partial: string;
	/** Significant token preceding the partial identifier (or the cursor). */
	before?: VbaToken;
	/** Offset where the statement containing the cursor begins. */
	statementStart: number;
	/** True when the cursor sits at the end of a comment token. */
	inComment: boolean;
	/** True when the cursor sits at the end of a string-literal token. */
	inString: boolean;
}

// One completion/signature/hover request fans out to many resolvers (and to
// per-item helpers) that all ask for the same cursor context, so a tiny memo
// keyed on (source, offset) collapses the prefix tokenizations to one.
const CURSOR_CONTEXT_CACHE_MAX = 4;
const cursorContextCache: {
	source: string;
	offset: number;
	context: CompletionCursorContext;
}[] = [];

/**
 * Analyzes the cursor position in one tokenize pass over the prefix. Callers
 * must not mutate the returned token arrays or their tokens.
 */
export function completionCursorContext(
	source: string,
	offset: number,
): CompletionCursorContext {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	for (let i = 0; i < cursorContextCache.length; i += 1) {
		const entry = cursorContextCache[i];
		if (entry.offset === safeOffset && entry.source === source) {
			if (i > 0) {
				cursorContextCache.splice(i, 1);
				cursorContextCache.unshift(entry);
			}
			return entry.context;
		}
	}
	const context = buildCursorContext(source, safeOffset);
	cursorContextCache.unshift({ source, offset: safeOffset, context });
	if (cursorContextCache.length > CURSOR_CONTEXT_CACHE_MAX) {
		cursorContextCache.pop();
	}
	return context;
}

function buildCursorContext(
	source: string,
	safeOffset: number,
): CompletionCursorContext {
	const tokens = tokenize(source.slice(0, safeOffset));
	const significantTokens = tokens.filter((t) => t.kind !== 'comment');
	const last = tokens[tokens.length - 1];
	const lastSignificant = significantTokens[significantTokens.length - 1];
	const partialToken =
		lastSignificant && isIdentLike(lastSignificant) && lastSignificant.end === safeOffset
			? lastSignificant
			: undefined;
	return {
		offset: safeOffset,
		tokens,
		significantTokens,
		partialToken,
		partial: partialToken?.rawText ?? '',
		before: partialToken
			? significantTokens[significantTokens.length - 2]
			: lastSignificant,
		statementStart: statementStartOffset(tokens),
		inComment: last?.kind === 'comment' && last.end === safeOffset,
		inString: last?.kind === 'stringLiteral' && last.end === safeOffset,
	};
}

/** Start of the statement containing the cursor (after the last newline/':'). */
function statementStartOffset(tokens: readonly VbaToken[]): number {
	for (let i = tokens.length - 1; i >= 0; i -= 1) {
		if (tokens[i].kind === 'newline' || tokens[i].kind === 'colon') {
			return tokens[i].end;
		}
	}
	return 0;
}

/**
 * Char-level twin of the partial-identifier peel: the span of the identifier
 * ending exactly at `offset`, or undefined when the cursor does not end a
 * word. Used by per-keystroke paths that cannot afford a tokenize pass.
 */
export function identifierSpanEndingAt(
	source: string,
	offset: number,
): { start: number; end: number } | undefined {
	const end = Math.max(0, Math.min(offset, source.length));
	let start = end;
	while (start > 0 && /[A-Za-z0-9_]/.test(source[start - 1])) {
		start -= 1;
	}
	if (start === end || !/[A-Za-z_]/.test(source[start])) {
		return undefined;
	}
	return { start, end };
}
