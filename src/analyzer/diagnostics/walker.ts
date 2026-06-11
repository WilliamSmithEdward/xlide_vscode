// Shared AST/statement traversal utilities for the diagnostics engine.
//
// These walkers and statement-token helpers were extracted verbatim from
// `analyzeModule.ts` so individual rule modules can share one set of
// traversal primitives instead of each carrying private copies. They are
// pure: no rule logic, no diagnostics, no caching.

import type { VbaToken } from '../lexer/tokenKinds';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type {
	BodyNode,
	ModuleMember,
	ModuleNode,
	ProcedureNode,
	Span,
	StatementNode,
	VariableGroupNode,
} from '../parser/nodes';

// `tokenText` and `matchParenFrom` are byte-identical to the shared lexer
// helpers (`tokenWord`, `matchParenFrom`); re-export them so the diagnostics
// engine keeps one implementation. `statementTokens` comes from the per-pass
// cache in analysisContext.ts (audit #5) so every rule shares one
// tokenization per statement.
export { matchParenFrom, tokenWord as tokenText } from '../lexer/tokenHelpers';
export { statementTokens } from './analysisContext';
import { tokenWord as tokenText } from '../lexer/tokenHelpers';
import { statementTokens } from './analysisContext';
import { trackedLocalsPassedAsCallArguments } from './dataflow';

export function isInactiveNode(
	activity: ConditionalActivityTracker | undefined,
	node: { span: Span },
): boolean {
	return activity?.isInactive(node.span) ?? false;
}

export function activeModuleMembers(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): readonly ModuleMember[] {
	if (!activity) {
		return mod.members;
	}
	return mod.members.filter((member) => !isInactiveNode(activity, member));
}

/** Walks every StatementNode in a body, descending into nested blocks. */
export function forEachStatement(
	body: BodyNode[],
	visit: (stmt: StatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			forEachStatement(node.body, visit, activity);
		}
	}
}

/**
 * One per-procedure visitor of the shared statement walk (audit #0): given a
 * procedure, returns the per-statement callback to run inside it, or
 * undefined to skip the procedure entirely.
 */
export type ProcedureStatementVisitor = (
	proc: ProcedureNode,
) => ((stmt: StatementNode) => void) | undefined;

/**
 * Runs every registered per-statement rule on ONE walk over the module's
 * active procedures and statements (audit #0). Each visitor sees procedures
 * and statements in exactly the order the rules' former private walks used:
 * active members in source order, `forEachStatement` within each body.
 */
export function walkProcedureStatements(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	visitors: readonly ProcedureStatementVisitor[],
): void {
	if (visitors.length === 0) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const callbacks: Array<(stmt: StatementNode) => void> = [];
		for (const visitor of visitors) {
			const callback = visitor(member);
			if (callback) {
				callbacks.push(callback);
			}
		}
		if (callbacks.length === 0) {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			for (const callback of callbacks) {
				callback(stmt);
			}
		}, activity);
	}
}

/** Walks every VariableGroupNode in a body, descending into nested blocks. */
export function forEachVariableGroup(
	body: BodyNode[],
	visit: (group: VariableGroupNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'VariableGroup') {
			visit(node);
		} else if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			forEachVariableGroup((node as { body: BodyNode[] }).body, visit, activity);
		}
	}
}

/** Walks every generic StatementNode in a procedure body, descending into nested blocks. */
export function forEachBodyStatement(
	body: BodyNode[],
	visit: (statement: StatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			visit(node);
		} else if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			forEachBodyStatement((node as { body: BodyNode[] }).body, visit, activity);
		}
	}
}

export function forEachProcedureBodyLine(
	source: string,
	procedure: ProcedureNode,
	visit: (span: Span) => void,
): void {
	const firstBreak = firstLineBreakAtOrAfter(source, procedure.span.start);
	if (firstBreak < 0 || firstBreak >= procedure.span.end) {
		return;
	}
	let lineStart = nextLineStart(source, firstBreak);
	while (lineStart < procedure.span.end) {
		let lineEnd = lineStart;
		while (lineEnd < procedure.span.end && source[lineEnd] !== '\r' && source[lineEnd] !== '\n') {
			lineEnd++;
		}
		visit({ start: lineStart, end: lineEnd });
		lineStart = nextLineStart(source, lineEnd);
	}
}

export function nextLineStart(source: string, lineBreakOffset: number): number {
	if (
		source[lineBreakOffset] === '\r' &&
		lineBreakOffset + 1 < source.length &&
		source[lineBreakOffset + 1] === '\n'
	) {
		return lineBreakOffset + 2;
	}
	return lineBreakOffset + 1;
}

export function firstLineBreakAtOrAfter(source: string, start: number): number {
	for (let i = start; i < source.length; i++) {
		const ch = source[i];
		if (ch === '\n' || ch === '\r') {
			return i;
		}
	}
	return -1;
}

export function statementTokensAfterLeadingLabel(source: string, span: Span): VbaToken[] {
	const toks = statementTokens(source, span);
	const firstExecutable = firstExecutableTokenIndex(toks);
	return firstExecutable > 0 ? toks.slice(firstExecutable) : toks;
}

export function firstExecutableTokenIndex(toks: readonly VbaToken[]): number {
	if (toks.length > 1 && toks[0].kind === 'integerLiteral' && /^\d+$/.test(toks[0].rawText)) {
		return 1;
	}
	if (
		toks.length > 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		return 2;
	}
	return 0;
}

export function topLevelOperatorIndex(toks: readonly VbaToken[], operator: string): number {
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (depth === 0 && toks[i].kind === 'operator' && raw === operator) {
			return i;
		}
	}
	return -1;
}

export function tokenName(tok: VbaToken): string | undefined {
	if (tok.kind === 'identifier' || tok.kind === 'keyword') {
		return tok.rawText;
	}
	if (tok.kind === 'bracketedIdentifier') {
		return stripHeaderBrackets(tok.rawText);
	}
	return undefined;
}

export function stripHeaderBrackets(text: string): string {
	return text.startsWith('[') && text.endsWith(']')
		? text.slice(1, -1)
		: text;
}

export function absoluteSpan(base: Span, token: VbaToken): Span {
	return { start: base.start + token.start, end: base.start + token.end };
}

/**
 * If the statement spanning `span` is a simple assignment to a bare identifier
 * (`name = ...` or `Let name = ...`), returns that identifier and its span;
 * otherwise undefined. `Set` (object) assignments and any left-hand side with a
 * `.` or `(` are excluded so only true scalar-name assignments are considered.
 */
export function bareAssignmentTarget(
	source: string,
	span: Span,
): { name: string; span: Span; valueTokens: VbaToken[] } | undefined {
	const toks = statementTokens(source, span);
	let i = firstExecutableTokenIndex(toks);
	// Skip an explicit `Let`; bail on `Set` (object assignment).
	if (toks[i] && toks[i].kind === 'keyword') {
		const kw = toks[i].rawText.toLowerCase();
		if (kw === 'set') {
			return undefined;
		}
		if (kw === 'let') {
			i++;
		}
	}
	const nameTok = toks[i];
	if (!nameTok || nameTok.kind !== 'identifier') {
		return undefined; // first token must be a plain identifier LHS
	}
	const next = toks[i + 1];
	if (!next || next.kind !== 'operator' || next.rawText !== '=') {
		return undefined; // not `name =` (excludes `.`, `(`, `<=`, `<>`, comparisons)
	}
	return {
		name: nameTok.rawText,
		span: { start: span.start + nameTok.start, end: span.start + nameTok.end },
		valueTokens: toks.slice(i + 2),
	};
}

export function setAssignmentTarget(
	source: string,
	span: Span,
): { name: string; span: Span; valueTokens: VbaToken[] } | undefined {
	const toks = statementTokens(source, span);
	const i = firstExecutableTokenIndex(toks);
	if (tokenText(toks[i]) !== 'set') {
		return undefined;
	}
	const nameTok = toks[i + 1];
	const name = nameTok ? tokenName(nameTok) : undefined;
	if (!nameTok || !name) {
		return undefined;
	}
	const equals = toks[i + 2];
	if (!equals || equals.kind !== 'operator' || equals.rawText !== '=') {
		return undefined;
	}
	return {
		name,
		span: { start: span.start + nameTok.start, end: span.start + nameTok.end },
		valueTokens: toks.slice(i + 3),
	};
}

/** Lowercased tracked locals passed as bare call arguments in one statement. */
export function localsPassedAsCallArguments(
	source: string,
	span: Span,
	tracked: ReadonlyMap<string, unknown>,
): Set<string> {
	return trackedLocalsPassedAsCallArguments(
		statementTokensAfterLeadingLabel(source, span),
		(lower) => tracked.has(lower),
	);
}

export function blockHeaderLineSpan(source: string, span: Span): Span {
	const nl = firstLineBreakAtOrAfter(source, span.start);
	if (nl < 0 || nl > span.end) {
		return span;
	}
	return { start: span.start, end: nl };
}

export function blockFooterLineSpan(source: string, span: Span): Span {
	let start = span.end;
	while (start > span.start && source[start - 1] !== '\n' && source[start - 1] !== '\r') {
		start--;
	}
	return { start, end: span.end };
}

export function declaredNameSpan(source: string, span: Span, name: string): Span {
	const lower = name.toLowerCase();
	for (const tok of statementTokens(source, span)) {
		if (tokenName(tok)?.toLowerCase() === lower) {
			return absoluteSpan(span, tok);
		}
	}
	return span;
}

export function firstTokenSpan(source: string, span: Span): Span {
	const tok = statementTokens(source, span)[0];
	return tok ? absoluteSpan(span, tok) : span;
}

export function pluralizeCount(count: number, singular: string): string {
	return `${count} ${singular}${count === 1 ? '' : 's'}`;
}

export function physicalLineSpanAtOffset(source: string, offset: number): Span {
	const safe = Math.max(0, Math.min(offset, source.length));
	const before = source.lastIndexOf('\n', Math.max(0, safe - 1));
	const start = before < 0 ? 0 : before + 1;
	const after = source.indexOf('\n', safe);
	let end = after < 0 ? source.length : after;
	if (end > start && source[end - 1] === '\r') {
		end--;
	}
	return { start, end };
}
