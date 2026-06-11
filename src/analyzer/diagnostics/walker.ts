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

// `statementTokens` and `tokenText` are byte-identical to the shared lexer
// helpers (`statementTokens`, `tokenWord`); re-export them so the diagnostics
// engine keeps one implementation.
export { statementTokens, tokenWord as tokenText } from '../lexer/tokenHelpers';
import { statementTokens } from '../lexer/tokenHelpers';

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
