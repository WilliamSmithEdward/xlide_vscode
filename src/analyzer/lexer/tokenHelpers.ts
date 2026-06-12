// Small token utilities shared across analyzer features so statement-level
// token handling (comment/newline filtering, identifier extraction, leading
// line numbers, paren matching) cannot drift between surfaces.

import type { VbaToken } from './tokenKinds';
import { tokenize } from './tokenize';

export const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

export function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

/** Significant tokens of a source span, excluding comments and newlines. */
export function statementTokens(
	source: string,
	span: { start: number; end: number },
): VbaToken[] {
	return tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
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
		return token.rawText.slice(1, -1);
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
