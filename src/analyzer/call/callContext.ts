// Shared VBA call-site/context helpers.
//
// Completion, signature help, and diagnostics all need the same distinction:
// VBA has expression calls, explicit `Call` statements, and parenless call
// statements, and those contexts have different parenthesis rules. Keep that
// classification here so editor surfaces do not drift.

import { STATEMENT_KEYWORDS as STATEMENT_KEYWORD_LIST } from '../lexer/keywordTable';
import { VbaToken } from '../lexer/tokenKinds';
import { tokenize } from '../lexer/tokenize';

const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;

// Statement-like words that are not in MS-VBAL's statement-keyword list but
// still cannot be treated as bare parenless call targets.
const ADDITIONAL_NON_CALL_STATEMENT_LEADS = [
	'then',
	'property',
	'error',
	'line',
	'name',
	'kill',
	'mkdir',
	'rmdir',
	'chdir',
	'chdrive',
	'load',
	'unload',
];

/** Lowercase statement-leading words that must not be parsed as bare calls. */
export const STATEMENT_KEYWORDS: ReadonlySet<string> = new Set([
	...STATEMENT_KEYWORD_LIST.map((word) => word.toLowerCase()),
	...ADDITIONAL_NON_CALL_STATEMENT_LEADS,
]);

/** A located call whose argument list contains the caret. */
export interface VbaCallSite {
	calleeName: string;
	/** True when the callee is `receiver.Member` or a leading-dot With member. */
	isMember: boolean;
	/** True when the containing statement starts with the explicit `Call` keyword. */
	isExplicitCall: boolean;
	/** Absolute offset just past the callee identifier. */
	calleeEndOffset: number;
	/** Zero-based index of the argument the caret is in. */
	activeParameter: number;
}

export interface VbaTextSpan {
	start: number;
	end: number;
}

export interface BareCallStatementTarget {
	name: string;
	span: VbaTextSpan;
}

export interface ParenthesizedCallStatementTarget {
	name: string;
	span: VbaTextSpan;
	isMember: boolean;
	startsWithLeadingDot: boolean;
	calleeEndOffset: number;
}

export function isIdentLike(token: VbaToken): boolean {
	return (
		(token.kind === 'identifier' || token.kind === 'keyword') &&
		IDENT_RE.test(token.rawText)
	);
}

/**
 * If `span` covers a statement whose callee is a bare identifier, returns that
 * target. Member calls, assignments, labels, statement keywords, and implicit
 * Application index/member forms are intentionally excluded.
 */
export function bareCallStatementTarget(
	source: string,
	span: VbaTextSpan,
): BareCallStatementTarget | undefined {
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	if (toks.length === 0) {
		return undefined;
	}

	let idx = 0;
	const explicitCall = toks[0].rawText.toLowerCase() === 'call';
	if (explicitCall) {
		idx = 1;
	}

	const callee = toks[idx];
	if (!callee || callee.kind !== 'identifier') {
		return undefined;
	}
	if (STATEMENT_KEYWORDS.has(callee.rawText.toLowerCase())) {
		return undefined;
	}

	const result = {
		name: callee.rawText,
		span: { start: span.start + callee.start, end: span.start + callee.end },
	};

	const next = toks[idx + 1];
	if (!next) {
		if (!explicitCall) {
			let j = span.start + callee.end;
			while (j < source.length && (source[j] === ' ' || source[j] === '\t')) {
				j++;
			}
			if (source[j] === ':') {
				return undefined;
			}
		}
		return result;
	}

	const r = next.rawText;
	if (r === '.' || r === ':') {
		return undefined;
	}
	if (!explicitCall && r === '(') {
		return undefined;
	}
	if (!explicitCall) {
		const gap = source.slice(span.start + callee.end, span.start + next.start);
		if (!/\s/.test(gap)) {
			return undefined;
		}
	}
	let depth = 0;
	for (let k = idx + 1; k < toks.length; k += 1) {
		const tr = toks[k].rawText;
		if (tr === '(' || tr === '[') {
			depth += 1;
		} else if (tr === ')' || tr === ']') {
			depth -= 1;
		} else if (depth === 0 && tr === '=') {
			return undefined;
		}
	}
	return result;
}

export function explicitCallStatementTarget(
	source: string,
	span: VbaTextSpan,
): BareCallStatementTarget | undefined {
	const toks = statementTokens(source, span);
	if (toks.length < 2 || toks[0].rawText.toLowerCase() !== 'call') {
		return undefined;
	}
	const name = tokenName(toks[1]);
	if (!name) {
		return undefined;
	}
	return {
		name,
		span: { start: span.start + toks[1].start, end: span.start + toks[1].end },
	};
}

export function explicitCallStatementArgumentWithoutParens(
	source: string,
	span: VbaTextSpan,
): VbaTextSpan | undefined {
	const toks = statementTokens(source, span);
	if (toks.length === 0 || toks[0].rawText.toLowerCase() !== 'call') {
		return undefined;
	}
	const consumed = consumeCallableChain(toks, 1);
	if (!consumed) {
		return undefined;
	}
	const stray = toks[consumed.nextIndex];
	return stray
		? { start: span.start + stray.start, end: span.start + stray.end }
		: undefined;
}

export function standaloneEmptyParenthesizedCallStatement(
	source: string,
	span: VbaTextSpan,
): ParenthesizedCallStatementTarget | undefined {
	const toks = statementTokens(source, span);
	if (
		toks.length < 3 ||
		toks[0]?.rawText.toLowerCase() === 'call' ||
		topLevelTokenIndex(toks, '=') >= 0
	) {
		return undefined;
	}
	for (let i = 0; i < toks.length - 2; i += 1) {
		const name = tokenName(toks[i]);
		if (!name || toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (
			close !== i + 2 ||
			close !== toks.length - 1 ||
			!isCompleteStatementChainThroughEmptyCall(toks, i, close)
		) {
			continue;
		}
		return {
			name,
			isMember: i > 0 && toks[i - 1]?.rawText === '.',
			startsWithLeadingDot: toks[0]?.rawText === '.',
			calleeEndOffset: span.start + toks[i].end,
			span: { start: span.start + toks[i].start, end: span.start + toks[close].end },
		};
	}
	return undefined;
}

/**
 * Locates the active VBA call at `offset`, if the caret is in a parenthesized
 * argument list or a parenless call statement's argument region.
 */
export function findActiveCallSite(source: string, offset: number): VbaCallSite | undefined {
	if (offset < 0) {
		return undefined;
	}
	const prefix = source.slice(0, offset);
	const tokens = tokenize(prefix).filter((t) => t.kind !== 'comment');
	if (tokens.length === 0) {
		return undefined;
	}
	return findParenCall(tokens) ?? findParenlessCall(tokens, source, offset);
}

/**
 * Callable completions can use parentheses in expression and explicit `Call`
 * contexts, but VBA call statements like `mySub()` and `Application.Calculate()`
 * are syntax errors unless prefixed with `Call`.
 */
export function callableCompletionShouldInsertParens(
	source: string,
	offset: number,
): boolean {
	const prefixText = source.slice(0, Math.max(0, offset));
	const tokens = tokenize(prefixText).filter((t) => t.kind !== 'comment');
	if (tokens.length === 0) {
		return false;
	}

	let last = tokens.length - 1;
	if (last >= 0 && isIdentLike(tokens[last])) {
		last -= 1;
	}
	if (last < 0) {
		return false;
	}

	let boundary = last;
	while (boundary >= 0 && !isStatementBoundary(tokens[boundary])) {
		boundary -= 1;
	}
	const statement = tokens.slice(boundary + 1, last + 1);
	if (statement.length === 0) {
		return false;
	}

	const prev = statement[statement.length - 1].rawText.toLowerCase();
	if (prev === 'call' || statement[0].rawText.toLowerCase() === 'call') {
		return true;
	}
	if (statementContainsExpressionIntroducer(statement)) {
		return true;
	}
	return isExpressionContinuationToken(prev);
}

export function isExplicitCallTargetCompletionContext(
	tokens: readonly VbaToken[],
	last: number,
): boolean {
	if (last < 0) {
		return false;
	}
	let boundary = last;
	while (boundary >= 0 && !isStatementBoundary(tokens[boundary])) {
		boundary -= 1;
	}
	const statement = tokens.slice(boundary + 1, last + 1);
	return statement.length === 1 && statement[0].rawText.toLowerCase() === 'call';
}

/**
 * Locates the innermost enclosing *parenthesized* call whose argument list the
 * caret sits in. Returns undefined when the caret is not inside any call paren.
 */
function findParenCall(tokens: readonly VbaToken[]): VbaCallSite | undefined {
	interface Frame {
		isCall: boolean;
		openIndex: number;
		commaCount: number;
	}
	const stack: Frame[] = [];
	for (let k = 0; k < tokens.length; k += 1) {
		const t = tokens[k];
		if (t.kind === 'newline' || t.rawText === ':') {
			stack.length = 0;
			continue;
		}
		const r = t.rawText;
		if (r === '(') {
			const prev = tokens[k - 1];
			const isCall = !!prev && isIdentLike(prev);
			stack.push({ isCall, openIndex: k, commaCount: 0 });
		} else if (r === ')') {
			stack.pop();
		} else if (r === ',' && stack.length > 0) {
			stack[stack.length - 1].commaCount += 1;
		}
	}
	for (let i = stack.length - 1; i >= 0; i -= 1) {
		if (stack[i].isCall) {
			const open = stack[i].openIndex;
			const callee = tokens[open - 1];
			const isMember = open - 2 >= 0 && tokens[open - 2].rawText === '.';
			return {
				calleeName: callee.rawText,
				isMember,
				isExplicitCall: isExplicitCallTarget(tokens, open - 1),
				calleeEndOffset: callee.end,
				activeParameter: stack[i].commaCount,
			};
		}
	}
	return undefined;
}

function findParenlessCall(
	tokens: readonly VbaToken[],
	source: string,
	offset: number,
): VbaCallSite | undefined {
	let start = 0;
	for (let k = tokens.length - 1; k >= 0; k -= 1) {
		if (tokens[k].kind === 'newline' || tokens[k].rawText === ':') {
			start = k + 1;
			break;
		}
	}
	const stmt = tokens.slice(start);
	if (stmt.length === 0) {
		return undefined;
	}
	let idx = 0;
	const isExplicitCall = stmt[0].rawText.toLowerCase() === 'call';
	if (isExplicitCall) {
		idx = 1;
	}
	const calleeSite = parenlessCalleeSite(stmt, idx, source);
	if (!calleeSite) {
		return undefined;
	}
	const callee = stmt[calleeSite.calleeIndex];
	if (!calleeSite.isMember && STATEMENT_KEYWORDS.has(callee.rawText.toLowerCase())) {
		return undefined;
	}
	const afterTokens = stmt.slice(calleeSite.calleeIndex + 1);
	const gap = source.slice(callee.end, offset);
	const argsStarted = afterTokens.length > 0 || /\s/.test(gap);
	if (!argsStarted) {
		return undefined;
	}
	let depth = 0;
	let commaCount = 0;
	for (const t of afterTokens) {
		const r = t.rawText;
		if (r === '(' || r === '[') {
			depth += 1;
		} else if (r === ')' || r === ']') {
			depth -= 1;
		} else if (depth === 0) {
			if (r === '=') {
				return undefined;
			}
			if (r === ',') {
				commaCount += 1;
			}
		}
	}
	return {
		calleeName: callee.rawText,
		isMember: calleeSite.isMember,
		isExplicitCall,
		calleeEndOffset: callee.end,
		activeParameter: commaCount,
	};
}

function parenlessCalleeSite(
	stmt: readonly VbaToken[],
	startIndex: number,
	source: string,
): { calleeIndex: number; isMember: boolean } | undefined {
	if (startIndex >= stmt.length) {
		return undefined;
	}
	let i = startIndex;
	let isMember = false;
	if (stmt[i].rawText === '.') {
		isMember = true;
		i += 1;
	}
	if (!stmt[i] || !isIdentLike(stmt[i])) {
		return undefined;
	}
	let calleeIndex = i;
	for (;;) {
		const paren = stmt[i + 1];
		if (paren?.rawText === '(' && noWhitespaceBetween(source, stmt[i], paren)) {
			const close = matchParenFrom(stmt, i + 1);
			if (close < 0 || stmt[close + 1]?.rawText !== '.') {
				break;
			}
			i = close + 1;
		}
		if (stmt[i + 1]?.rawText !== '.') {
			break;
		}
		const nextNameIndex = i + 2;
		const nextName = stmt[nextNameIndex];
		if (!nextName || !isIdentLike(nextName)) {
			return undefined;
		}
		isMember = true;
		i = nextNameIndex;
		calleeIndex = i;
	}
	return { calleeIndex, isMember };
}

function noWhitespaceBetween(source: string, left: VbaToken, right: VbaToken): boolean {
	return !/\s/.test(source.slice(left.end, right.start));
}

function matchParenFrom(tokens: readonly VbaToken[], open: number): number {
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

function statementTokens(source: string, span: VbaTextSpan): VbaToken[] {
	return tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
}

function tokenName(token: VbaToken | undefined): string | undefined {
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

function consumeCallableChain(
	tokens: readonly VbaToken[],
	start: number,
): { nextIndex: number } | undefined {
	if (!tokenName(tokens[start])) {
		return undefined;
	}
	let i = start + 1;
	for (;;) {
		const t = tokens[i];
		if (!t) {
			return { nextIndex: i };
		}
		if (t.rawText === '.') {
			if (!tokenName(tokens[i + 1])) {
				return { nextIndex: i };
			}
			i += 2;
			continue;
		}
		if (t.rawText === '(') {
			const close = matchParenFrom(tokens, i);
			if (close < 0) {
				return undefined;
			}
			i = close + 1;
			continue;
		}
		return { nextIndex: i };
	}
}

function isCompleteStatementChainThroughEmptyCall(
	toks: readonly VbaToken[],
	calleeIdx: number,
	closeIdx: number,
): boolean {
	if (calleeIdx === 0) {
		return !!tokenName(toks[0]);
	}
	const first = toks[0];
	if (!first) {
		return false;
	}
	let i = 1;
	if (first.rawText === '.') {
		const nameIdx = 1;
		if (!tokenName(toks[nameIdx])) {
			return false;
		}
		if (nameIdx === calleeIdx) {
			return toks[nameIdx + 1]?.rawText === '(' && matchParenFrom(toks, nameIdx + 1) === closeIdx;
		}
		i = nameIdx + 1;
	} else if (!tokenName(first)) {
		return false;
	}
	while (i < toks.length) {
		const raw = toks[i]?.rawText;
		if (raw === '(') {
			const close = matchParenFrom(toks, i);
			if (close < 0 || close >= calleeIdx) {
				return false;
			}
			i = close + 1;
			continue;
		}
		if (raw !== '.') {
			return false;
		}
		const nameIdx = i + 1;
		if (!tokenName(toks[nameIdx])) {
			return false;
		}
		if (nameIdx === calleeIdx) {
			return toks[nameIdx + 1]?.rawText === '(' && matchParenFrom(toks, nameIdx + 1) === closeIdx;
		}
		i = nameIdx + 1;
	}
	return false;
}

function topLevelTokenIndex(tokens: readonly VbaToken[], rawText: string): number {
	let depth = 0;
	for (let i = 0; i < tokens.length; i += 1) {
		const raw = tokens[i].rawText;
		if (raw === '(' || raw === '[') {
			depth += 1;
		} else if (raw === ')' || raw === ']') {
			depth -= 1;
		} else if (depth === 0 && raw === rawText) {
			return i;
		}
	}
	return -1;
}

function isExplicitCallTarget(tokens: readonly VbaToken[], calleeIndex: number): boolean {
	let start = calleeIndex;
	while (start > 0) {
		const prev = tokens[start - 1];
		if (prev.kind === 'newline' || prev.rawText === ':') {
			break;
		}
		start--;
	}
	return tokens[start]?.rawText.toLowerCase() === 'call';
}

function isStatementBoundary(token: VbaToken): boolean {
	return token.kind === 'newline' || token.rawText === ':';
}

function statementContainsExpressionIntroducer(tokens: readonly VbaToken[]): boolean {
	let depth = 0;
	for (const token of tokens) {
		if (token.rawText === '(') {
			depth += 1;
			continue;
		}
		if (token.rawText === ')') {
			depth = Math.max(0, depth - 1);
			continue;
		}
		if (depth > 0) {
			continue;
		}
		const lower = token.rawText.toLowerCase();
		if (token.rawText === '=') {
			return true;
		}
		if (
			lower === 'if' ||
			lower === 'elseif' ||
			lower === 'while' ||
			lower === 'until' ||
			lower === 'case' ||
			lower === 'select' ||
			lower === 'for' ||
			lower === 'to' ||
			lower === 'step'
		) {
			return true;
		}
	}
	return false;
}

function isExpressionContinuationToken(lowerTokenText: string): boolean {
	return (
		lowerTokenText === '(' ||
		lowerTokenText === ',' ||
		lowerTokenText === '=' ||
		lowerTokenText === '+' ||
		lowerTokenText === '-' ||
		lowerTokenText === '*' ||
		lowerTokenText === '/' ||
		lowerTokenText === '\\' ||
		lowerTokenText === '&' ||
		lowerTokenText === '<' ||
		lowerTokenText === '>' ||
		lowerTokenText === '<=' ||
		lowerTokenText === '>=' ||
		lowerTokenText === '<>' ||
		lowerTokenText === '^' ||
		lowerTokenText === 'and' ||
		lowerTokenText === 'or' ||
		lowerTokenText === 'xor' ||
		lowerTokenText === 'eqv' ||
		lowerTokenText === 'imp' ||
		lowerTokenText === 'mod' ||
		lowerTokenText === 'not' ||
		lowerTokenText === 'like' ||
		lowerTokenText === 'is'
	);
}
