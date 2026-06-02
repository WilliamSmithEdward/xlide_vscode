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
import {
	resolveTypeCompletions,
	type TypeCompletionContext,
} from './typeCompletion';

export interface CanonicalCaseContext {
	member?: MemberCompletionContext;
	identifier?: IdentifierCompletionContext;
	type?: TypeCompletionContext;
}

export interface CanonicalCaseEdit {
	start: number;
	end: number;
	text: string;
}

export interface CanonicalCaseSpan {
	start: number;
	end: number;
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
		canonicalFromTypeCompletion(source, span.end, word, ctx.type) ??
		canonicalFromMemberCompletion(source, span.end, word, ctx.member) ??
		canonicalFromIdentifierCompletion(source, span.end, word, ctx.identifier);
	if (!canonical || canonical === word) {
		return undefined;
	}
	return { ...span, text: canonical };
}

export function resolveCanonicalCaseEdits(
	source: string,
	span: CanonicalCaseSpan,
	ctx: CanonicalCaseContext = {},
): CanonicalCaseEdit[] {
	const safeStart = Math.max(0, Math.min(span.start, source.length));
	const safeEnd = Math.max(safeStart, Math.min(span.end, source.length));
	const edits: CanonicalCaseEdit[] = [];
	for (const token of tokenize(source)) {
		if (token.start < safeStart || token.end > safeEnd || !isIdentifierToken(token)) {
			continue;
		}
		const edit = resolveCanonicalCaseEdit(source, token.end, ctx);
		if (edit && edit.start === token.start && edit.end === token.end) {
			edits.push(edit);
		}
	}
	return edits;
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

function canonicalFromTypeCompletion(
	source: string,
	offset: number,
	word: string,
	ctx: TypeCompletionContext = {},
): string | undefined {
	return resolveTypeCompletions(source, offset, ctx).find(
		(item) => item.name.toLowerCase() === word.toLowerCase(),
	)?.name;
}
