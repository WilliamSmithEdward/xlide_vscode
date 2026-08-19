// Small token utilities shared across analyzer features so statement-level
// token handling (comment/newline filtering, identifier extraction, leading
// line numbers, paren matching) cannot drift between surfaces.

import type { VbaToken } from './tokenKinds';
import { tokenize, tokenizeCached } from './tokenize';

// \p{L}/\p{M}: VBA identifiers may use any locale letter, and a combining mark
// continues a name. The ASCII-only form made `Dim g As Прибор` resolve to no
// type at all, so that receiver offered no members - not even its ASCII ones.
export const IDENT_RE = /^[\p{L}_][\p{L}\p{M}\p{N}_]*$/u;

export function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

/**
 * Significant tokens of a source span, excluding comments and newlines.
 *
 * Re-lexes the slice on every call. Use this only for one-off or derived
 * strings (where {@link statementTokensCached} would thrash its by-source
 * cache); every surface walking a module's own statements should use the
 * cached variant instead.
 */
export function statementTokens(
	source: string,
	span: { start: number; end: number },
): VbaToken[] {
	return tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
}

// Statement-token cache (audit #5, hoisted here from the diagnostics engine
// so every surface shares one implementation): independent consumers
// re-tokenized the same statement 25-40 times per pass. Tokens are cached
// per source string (value identity, LRU of 2 like tokenizeCached) and per
// statement span, so one pass lexes each statement once. Callers must not
// mutate the returned arrays or their tokens.
const STATEMENT_TOKEN_CACHE_MAX = 2;
const statementTokenCache: { src: string; byStart: Map<number, Map<number, VbaToken[]>> }[] = [];

/** Significant tokens of a statement span, excluding comments and newlines (memoized per pass). */
export function statementTokensCached(
	source: string,
	span: { start: number; end: number },
): VbaToken[] {
	let entry: { src: string; byStart: Map<number, Map<number, VbaToken[]>> } | undefined;
	for (let i = 0; i < statementTokenCache.length; i += 1) {
		if (statementTokenCache[i].src === source) {
			entry = statementTokenCache[i];
			if (i > 0) {
				statementTokenCache.splice(i, 1);
				statementTokenCache.unshift(entry);
			}
			break;
		}
	}
	if (!entry) {
		entry = { src: source, byStart: new Map() };
		statementTokenCache.unshift(entry);
		if (statementTokenCache.length > STATEMENT_TOKEN_CACHE_MAX) {
			statementTokenCache.pop();
		}
	}
	// Two number-keyed levels instead of a `${start}:${end}` string key: a
	// full pass looks statements up hundreds of thousands of times, and the
	// per-call key allocation plus string hashing was the analysis pass's
	// single largest profile frame. Integer keys also cannot lose precision
	// the way the ancient `start * 2^32 + end` numeric scheme did.
	let byEnd = entry.byStart.get(span.start);
	if (!byEnd) {
		byEnd = new Map();
		entry.byStart.set(span.start, byEnd);
	}
	let toks = byEnd.get(span.end);
	if (!toks) {
		toks = deriveStatementTokens(source, span) ?? statementTokens(source, span);
		byEnd.set(span.end, toks);
	}
	return toks;
}

/**
 * Derives a statement's span-relative significant tokens from the module's
 * shared token stream instead of re-lexing the statement's text. The whole
 * module is already tokenized once (tokenizeCached); re-running the lexer per
 * statement was the analysis pass's largest remaining cost on big modules.
 * Statement spans start at statement boundaries, which are line-start lexer
 * contexts in both the module stream and an isolated slice, so the token
 * streams agree; if a module token ever straddles the span boundary (which a
 * well-formed statement span never produces), we return undefined and the
 * caller falls back to lexing the slice.
 */
function deriveStatementTokens(
	source: string,
	span: { start: number; end: number },
): VbaToken[] | undefined {
	const all = tokenizeCached(source);
	// Binary search: first token ending after the span starts.
	let lo = 0;
	let hi = all.length - 1;
	let first = all.length;
	while (lo <= hi) {
		const mid = (lo + hi) >> 1;
		if (all[mid].end > span.start) {
			first = mid;
			hi = mid - 1;
		} else {
			lo = mid + 1;
		}
	}
	const out: VbaToken[] = [];
	for (let i = first; i < all.length; i += 1) {
		const tok = all[i];
		if (tok.start >= span.end) {
			break;
		}
		if (tok.start < span.start || tok.end > span.end) {
			return undefined;
		}
		if (tok.kind === 'comment' || tok.kind === 'newline') {
			continue;
		}
		out.push({ ...tok, start: tok.start - span.start, end: tok.end - span.start });
	}
	return out;
}

/** Identifier-like name of a token (unwraps bracketed identifiers). */
export function tokenName(token: VbaToken | undefined): string | undefined {
	if (!token) {
		return undefined;
	}
	if (token.kind === 'identifier' || token.kind === 'keyword') {
		return token.rawText;
	}
	if (token.kind === 'bracketedIdentifier') {
		// Strip the surrounding brackets, but only when both are present: the
		// tokenizer still emits a bracketedIdentifier for an unterminated
		// `[name` at line end (tokenize.ts), where a blind slice(1, -1) would
		// drop a real character.
		const raw = token.rawText;
		return raw.startsWith('[') && raw.endsWith(']') ? raw.slice(1, -1) : raw;
	}
	return undefined;
}

/**
 * Canonical (case-folded) text of a token used for keyword matching. For
 * keyword tokens the lexer already provides canonicalText; otherwise we lower
 * the raw text because VBA is case-insensitive (MS-VBAL 3.3.5).
 */
export function tokenWord(token: VbaToken | undefined): string {
	return (token?.canonicalText ?? token?.rawText ?? '').toLowerCase();
}

/** True when the token is a decimal line-number literal. */
export function isDecimalLineNumber(token: VbaToken | undefined): boolean {
	return token?.kind === 'integerLiteral' && /^\d+$/.test(token.rawText);
}

/** Drops the leading line-number token when one prefixes the statement. */
export function tokensWithoutLeadingLineNumber(tokens: readonly VbaToken[]): VbaToken[] {
	return tokens.length > 1 && isDecimalLineNumber(tokens[0])
		? [...tokens.slice(1)]
		: [...tokens];
}

/** Index of the ')' matching the '(' at `open`, or -1 when unmatched. */
export function matchParenFrom(tokens: readonly VbaToken[], open: number): number {
	let depth = 0;
	for (let i = open; i < tokens.length; i += 1) {
		const raw = tokens[i].rawText;
		if (raw === '(') {
			depth += 1;
		} else if (raw === ')') {
			depth -= 1;
			if (depth === 0) {
				return i;
			}
		}
	}
	return -1;
}

/** Split tokens[from, to) into separator-delimited groups at paren depth 0. */
export function splitTopLevelTokenGroups(
	tokens: readonly VbaToken[],
	from: number,
	separator: string,
	to = tokens.length,
): VbaToken[][] {
	const groups: VbaToken[][] = [];
	let current: VbaToken[] = [];
	let depth = 0;
	for (let i = from; i < to; i++) {
		const raw = tokens[i].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth = Math.max(0, depth - 1);
		}
		if (depth === 0 && raw === separator) {
			groups.push(current);
			current = [];
			continue;
		}
		current.push(tokens[i]);
	}
	groups.push(current);
	return groups;
}
