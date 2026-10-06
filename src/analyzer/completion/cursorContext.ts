// Shared cursor-context detection for the completion stack.
//
// Every completion-adjacent resolver needs the same prefix analysis: tokenize
// the text before the cursor, peel off the trailing partial identifier being
// typed, classify the token that precedes it, and know whether the cursor sits
// inside a comment or string literal. Keeping that dance here (one tokenize
// pass per cursor position) stops the per-resolver reimplementations from
// drifting and gives the resolvers a single seam for sharing token state.

import { assignmentTargetFromTokens } from './assignmentTarget';
import { tokenize, tokenizeCached } from '../lexer/tokenize';
import { isWsc } from '../lexer/tokenKinds';
import { lineStartAtAnyBreak, lineEndAtOrAfter } from '../../vbaSourceScan';
import type { VbaToken, Trivia } from '../lexer/tokenKinds';
import { isIdentLike, tokensWithoutLeadingLineNumber } from '../lexer/tokenHelpers';

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
			// Adopt the caller's instance so later lookups settle on the pointer
			// rather than comparing the whole module again (issue #45).
			entry.source = source;
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

type PrefixWindow = 'all' | 'line' | 'type';

// Physical rows ending in an underscore may belong to the same logical line.
// Include them conservatively: comments can continue as well as code, and an
// underscore inside a token may only be distinguishable after lexing the row.
function beforePhysicalBreak(source: string, start: number): number {
	return source[start - 1] === '\n' && source[start - 2] === '\r' ? start - 2 : start - 1;
}

function possibleContinuationLineStart(source: string, offset: number): number {
	let start = lineStartAtAnyBreak(source, offset);
	while (start > 0) {
		const priorEnd = beforePhysicalBreak(source, start);
		const priorStart = lineStartAtAnyBreak(source, priorEnd);
		let at = priorEnd;
		while (at > priorStart && isWsc(source[at - 1])) { at--; }
		if (at <= priorStart || source[at - 1] !== '_') { break; }
		start = priorStart;
	}
	return start;
}

function boundedLineTokens(source: string, offset: number): VbaToken[] {
	if (Number.isNaN(offset)) { return []; }
	// A cut inside CRLF belongs to the preceding physical line.
	let anchor = Math.trunc(offset);
	if (source[anchor] === '\n' && source[anchor - 1] === '\r') { anchor--; }
	const currentStart = possibleContinuationLineStart(source, anchor);
	// Include one preceding logical line to preserve the leading trivia and
	// column of the newline that starts the returned grammar window.
	const start = currentStart > 0
		? possibleContinuationLineStart(source, beforePhysicalBreak(source, currentStart))
		: 0;
	let end = anchor;
	while (true) {
		const lineStart = lineStartAtAnyBreak(source, end);
		const lineEnd = lineEndAtOrAfter(source, end);
		let at = lineEnd;
		while (at > lineStart && isWsc(source[at - 1])) { at--; }
		const continues = at > lineStart && source[at - 1] === '_';
		const breakWidth = source[lineEnd] === '\r' && source[lineEnd + 1] === '\n' ? 2 : 1;
		end = lineEnd < source.length ? lineEnd + breakWidth : lineEnd;
		if (!continues || end >= source.length) { break; }
	}
	// Typing resolvers consume spans and grammar, not absolute line numbers.
	// Count the prefix only if a caller explicitly asks for line metadata, and
	// share that result across every token in this window.
	let lineBase: number | undefined;
	const absoluteLineBase = (): number => {
		if (lineBase !== undefined) { return lineBase; }
		lineBase = 0;
		for (let i = 0; i < start; i++) {
			if (source[i] === '\n' || (source[i] === '\r' && source[i + 1] !== '\n')) { lineBase++; }
		}
		return lineBase;
	};
	const shiftTrivia = (items: Trivia[]): Trivia[] => items.map(item => ({
		...item, start: item.start + start, end: item.end + start,
	}));
	return tokenize(source.slice(start, end)).map(token => {
		const shifted: VbaToken = {
			kind: token.kind, rawText: token.rawText,
			start: token.start + start, end: token.end + start,
			get line() { return token.line + absoluteLineBase(); },
			character: token.character,
		};
		if (token.canonicalText !== undefined) { shifted.canonicalText = token.canonicalText; }
		if (token.leadingTrivia) { shifted.leadingTrivia = shiftTrivia(token.leadingTrivia); }
		if (token.trailingTrivia) { shifted.trailingTrivia = shiftTrivia(token.trailingTrivia); }
		return shifted;
	});
}

/**
 * Full-prefix/type windows derive from the memoized full-module stream; line
 * windows lex their own logical boundary. Each avoids re-lexing the prefix.
 * Every keystroke asks
 * for a new offset, and lexing a large module's prefix per keystroke was
 * the completion path's dominant cost. Tokens ending at or before the
 * offset are shared with the cached stream verbatim; when the offset lands
 * inside a token (a string, comment, identifier, or operator being typed),
 * only that token's remainder is re-lexed, which reproduces exactly what
 * lexing the truncated prefix produces since tokenization is local from a
 * token boundary onward.
 */
function prefixTokens(source: string, safeOffset: number, window: PrefixWindow = 'all'): VbaToken[] {
	const all = window === 'line'
		? boundedLineTokens(source, safeOffset)
		: tokenizeCached(source);
	// First token that ends after the offset.
	let lo = 0;
	let hi = all.length;
	while (lo < hi) {
		const mid = (lo + hi) >> 1;
		if (all[mid].end <= safeOffset) {
			lo = mid + 1;
		} else {
			hi = mid;
		}
	}
	// Keep only the grammar window the caller needs. Newlines in the cached
	// stream are logical boundaries: absorbed continuations must stay intact.
	let headStart = 0;
	if (window === 'line') {
		headStart = Math.max(0, lo - 1);
		while (headStart > 0 && all[headStart].kind !== 'newline') {
			headStart -= 1;
		}
	} else if (window === 'type') {
		headStart = lo;
		let remaining = 5;
		while (headStart > 0 && remaining > 0) {
			const token = all[--headStart];
			if (token.kind !== 'comment' && token.kind !== 'newline') {
				remaining -= 1;
			}
		}
	}
	const rebase = (base: number) => (token: VbaToken): VbaToken => ({
		...token,
		start: token.start + base,
		end: token.end + base,
	});
	const boundary = all[lo];
	if (!boundary || boundary.start >= safeOffset) {
		// The cut lands between tokens. That gap is whitespace or a line
		// continuation in the full stream, but TRUNCATED it can lex
		// differently: a prefix ending right after a continuation's `_`
		// materializes a dangling token the full stream absorbed. Re-lex the
		// residue (at most a few whitespace characters) to reproduce exactly
		// what lexing the prefix produces.
		const head = all.slice(headStart, lo);
		const residueStart = lo > 0 ? all[lo - 1].end : 0;
		if (residueStart < safeOffset) {
			const residue = tokenize(source.slice(residueStart, safeOffset)).map(rebase(residueStart));
			if (residue.length > 0) {
				return [...head, ...residue];
			}
		}
		return head;
	}
	const tail = tokenize(source.slice(boundary.start, safeOffset)).map(rebase(boundary.start));
	return [...all.slice(headStart, lo), ...tail];
}

function buildCursorContext(
	source: string,
	safeOffset: number,
	window: PrefixWindow = 'all',
): CompletionCursorContext {
	const tokens = prefixTokens(source, safeOffset, window);
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

/** Last five grammar tokens, preserving type detection's existing newline semantics. */
export function completionTypeTokens(source: string, offset: number): VbaToken[] {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	return prefixTokens(source, safeOffset, 'type')
		.filter(token => token.kind !== 'comment' && token.kind !== 'newline')
		.slice(-5);
}

const lineContextCache: { source: string; offset: number; context: CompletionCursorContext }[] = [];

/**
 * Cursor analysis limited to the current logical line (including its leading
 * newline). Callers must not mutate the returned arrays or look into earlier
 * statements. Fresh source versions lex only this logical line and its
 * preceding boundary; absolute line metadata is calculated on demand.
 * Truncated tokens and continuation residue are re-lexed just as
 * in the full-prefix API.
 */
export function completionLineCursorContext(source: string, offset: number): CompletionCursorContext {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	const index = lineContextCache.findIndex(entry => entry.source === source && entry.offset === safeOffset);
	if (index >= 0) {
		const [entry] = lineContextCache.splice(index, 1);
		lineContextCache.unshift(entry);
		return entry.context;
	}
	const context = buildCursorContext(source, safeOffset, 'line');
	lineContextCache.unshift({ source, offset: safeOffset, context });
	if (lineContextCache.length > CURSOR_CONTEXT_CACHE_MAX) {
		lineContextCache.pop();
	}
	return context;
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
 * Cheap line-local gate for space-triggered completion requests. Every typed
 * space fires the completion provider; only a few grammar positions can
 * actually produce results there (a statement led by a keyword or directive,
 * a member-access dot, a fresh statement, or a statement continued from the
 * previous physical line). Returns false when the space landed in ordinary
 * code (expressions, argument lists, comments, strings) so the provider can
 * bail before running any full-source resolver.
 */
export function spaceTriggerMayComplete(
	linePrefix: string,
	continuedFromPreviousLine = false,
): boolean {
	if (continuedFromPreviousLine) {
		return true; // the statement starts on an earlier physical line
	}
	if (/^[ \t]*$/.test(linePrefix) || /\.[ \t]*$/.test(linePrefix)) {
		return true; // statement start, or a member-access dot
	}
	const tokens = tokenize(linePrefix);
	const last = tokens[tokens.length - 1];
	if (
		(last?.kind === 'comment' || last?.kind === 'stringLiteral') &&
		last.end === linePrefix.length
	) {
		return false;
	}
	const significant = tokens.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	let start = 0;
	for (let i = significant.length - 1; i >= 0; i -= 1) {
		if (significant[i].kind === 'colon') {
			start = i + 1;
			break;
		}
	}
	const statement = tokensWithoutLeadingLineNumber(significant.slice(start));
	const head = statement[0];
	if (!head) {
		return true; // fresh statement after a ':' separator
	}
	return Boolean(assignmentTargetFromTokens(statement)) || head.kind === 'keyword' || head.kind === 'directive';
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
	// Marks are walked over as well as letters, or a Thai name would be cut at
	// its tone mark; the first character must still be a letter, since a mark
	// cannot begin a name.
	while (start > 0 && /[\p{L}\p{M}\p{N}_]/u.test(source[start - 1])) {
		start -= 1;
	}
	if (start === end || !/[\p{L}_]/u.test(source[start])) {
		return undefined;
	}
	return { start, end };
}

/** Line-local gate for automatic value suggestions on '=' (including ':='). */
export function assignmentValueTriggerMayComplete(linePrefix: string, continued = false): boolean {
	const cursor = completionLineCursorContext(linePrefix, linePrefix.length);
	if (cursor.inComment || cursor.inString) { return false; }
	if (continued) { return true; }
	const tokens = cursor.significantTokens.filter(t => t.start >= cursor.statementStart);
	return tokens.at(-1)?.rawText === ':=' || Boolean(assignmentTargetFromTokens(tokens));
}
