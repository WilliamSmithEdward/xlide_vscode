import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { identifierSpanEndingAt } from './cursorContext';
import {
	resolveIdentifierCompletions,
	type IdentifierCompletionContext,
} from './identifierCompletion';
import {
	resolveMemberCompletionNamed,
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

export type CanonicalCaseBoundaryKind = 'token' | 'line';

const TOKEN_BOUNDARY_CHARS = new Set([
	'(',
	')',
	'.',
	',',
	'=',
	':',
	'+',
	'-',
	'*',
	'/',
	'\\',
	'&',
	'<',
	'>',
]);

export function canonicalCaseBoundaryKind(text: string): CanonicalCaseBoundaryKind | undefined {
	if (/^\r?\n[ \t]*$/.test(text)) {
		return 'line';
	}
	if (/^[ \t]$/.test(text) || TOKEN_BOUNDARY_CHARS.has(text)) {
		return 'token';
	}
	return undefined;
}

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

	const canonical = canonicalTextForWord(source, span.end, word, token, ctx);
	if (!canonical) {
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
	const window = physicalLineWindow(source, safeStart, safeEnd);
	const edits: CanonicalCaseEdit[] = [];
	for (const token of tokenize(source.slice(window.start, window.end))) {
		const start = window.start + token.start;
		const end = window.start + token.end;
		if (start < safeStart || end > safeEnd || !isIdentifierToken(token)) {
			continue;
		}
		// Match the single-position path: the char-level span ending at the
		// token end must agree with the token, so e.g. an identifier glued to
		// a numeric literal never produces an edit.
		const charSpan = identifierSpanEndingAt(source, end);
		if (!charSpan || charSpan.start !== start || charSpan.end !== end) {
			continue;
		}
		const word = source.slice(start, end);
		const canonical = canonicalTextForWord(source, end, word, token, ctx);
		if (canonical) {
			edits.push({ start, end, text: canonical });
		}
	}
	return edits;
}

/** Canonical replacement text for `word` ending at `offset`, if it differs. */
function canonicalTextForWord(
	source: string,
	offset: number,
	word: string,
	token: VbaToken,
	ctx: CanonicalCaseContext,
): string | undefined {
	const canonical =
		token.canonicalText ??
		canonicalFromTypeCompletion(source, offset, word, ctx.type) ??
		canonicalFromMemberCompletion(source, offset, word, ctx.member) ??
		canonicalFromIdentifierCompletion(source, offset, word, ctx.identifier);
	return canonical && canonical !== word ? canonical : undefined;
}

/**
 * Finds the token covering exactly [start, end) by tokenizing only the
 * physical line(s) of the span. VBA comments and string literals never span
 * physical lines, so the line-local classification matches the whole-module
 * tokenization without rescanning the entire document per keystroke.
 */
function tokenAtSpan(source: string, start: number, end: number): VbaToken | undefined {
	const window = physicalLineWindow(source, start, end);
	return tokenize(source.slice(window.start, window.end)).find(
		(token) => window.start + token.start === start && window.start + token.end === end,
	);
}

/** Expands [start, end] to the enclosing physical line boundaries. */
function physicalLineWindow(
	source: string,
	start: number,
	end: number,
): { start: number; end: number } {
	const before = Math.max(
		source.lastIndexOf('\n', Math.max(0, start - 1)),
		source.lastIndexOf('\r', Math.max(0, start - 1)),
	);
	let windowEnd = end;
	while (windowEnd < source.length && source[windowEnd] !== '\n' && source[windowEnd] !== '\r') {
		windowEnd += 1;
	}
	return { start: before + 1, end: windowEnd };
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
	return resolveMemberCompletionNamed(source, offset, word, ctx)?.name;
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
