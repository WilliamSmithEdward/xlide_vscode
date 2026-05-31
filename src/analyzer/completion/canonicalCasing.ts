import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import {
	resolveIdentifierCompletions,
	type IdentifierCompletionContext,
} from './identifierCompletion';
import {
	resolveMemberCompletions,
	type MemberCompletionContext,
} from './memberAccess';

export interface CanonicalCaseContext {
	member?: MemberCompletionContext;
	identifier?: IdentifierCompletionContext;
}

export interface CanonicalCaseEdit {
	start: number;
	end: number;
	text: string;
}

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

export function resolveCanonicalCaseEdit(
	source: string,
	offset: number,
	ctx: CanonicalCaseContext = {},
): CanonicalCaseEdit | undefined {
	const span = identifierSpanEndingAt(source, offset);
	if (!span) {
		return undefined;
	}
	const word = source.slice(span.start, span.end);
	const token = tokenAtSpan(source, span.start, span.end);
	if (!token || !isIdentifierToken(token)) {
		return undefined;
	}

	const canonical =
		token.canonicalText ??
		canonicalFromMemberCompletion(source, span.end, word, ctx.member) ??
		canonicalFromIdentifierCompletion(source, span.end, word, ctx.identifier);
	if (!canonical || canonical === word) {
		return undefined;
	}
	return { ...span, text: canonical };
}

function identifierSpanEndingAt(
	source: string,
	offset: number,
): { start: number; end: number } | undefined {
	const end = Math.max(0, Math.min(offset, source.length));
	let start = end;
	while (start > 0 && /[A-Za-z0-9_]/.test(source[start - 1])) {
		start -= 1;
	}
	if (start === end) {
		return undefined;
	}
	const word = source.slice(start, end);
	return IDENT_RE.test(word) ? { start, end } : undefined;
}

function tokenAtSpan(source: string, start: number, end: number): VbaToken | undefined {
	return tokenize(source).find((token) => token.start === start && token.end === end);
}

function isIdentifierToken(token: VbaToken): boolean {
	return token.kind === 'identifier' || token.kind === 'keyword';
}

function canonicalFromMemberCompletion(
	source: string,
	offset: number,
	word: string,
	ctx: MemberCompletionContext = {},
): string | undefined {
	return resolveMemberCompletions(source, offset, ctx).find(
		(item) => item.name.toLowerCase() === word.toLowerCase(),
	)?.name;
}

function canonicalFromIdentifierCompletion(
	source: string,
	offset: number,
	word: string,
	ctx: IdentifierCompletionContext = {},
): string | undefined {
	return resolveIdentifierCompletions(source, offset, ctx).find(
		(item) => item.name.toLowerCase() === word.toLowerCase(),
	)?.name;
}
