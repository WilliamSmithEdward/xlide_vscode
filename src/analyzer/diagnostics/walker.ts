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
	LeafStatementNode,
	ModuleMember,
	ModuleNode,
	ProcedureNode,
	Span,
	StatementNode,
	VariableGroupNode,
} from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';

// `tokenText`, `tokenName`, and `matchParenFrom` are byte-identical to the
// shared lexer helpers (`tokenWord`, `tokenName`, `matchParenFrom`); re-export
// them so the diagnostics engine keeps one implementation. `statementTokens`
// comes from the per-pass cache in analysisContext.ts (audit #5) so every rule
// shares one tokenization per statement.
export { absoluteSpan, matchParenFrom, tokenWord as tokenText, tokenName } from '../lexer/tokenHelpers';
export { statementTokens } from './analysisContext';
export { blockHeaderStatements } from './blockHeaders';
import { blockHeaderStatements } from './blockHeaders';
import { tokenWord as tokenText, tokenName, absoluteSpan, statementTokens as lexStatementTokens } from '../lexer/tokenHelpers';
import { statementTokens } from './analysisContext';
import { trackedLocalsNamedWhole } from './dataflow';
import { calleeKeepsArgument } from './calleeArguments';

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

/** Walks every leaf statement (Assignment/Call/Statement) in a body, descending into nested blocks. */
export function forEachStatement(
	body: BodyNode[],
	visit: (stmt: LeafStatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (isLeafStatement(node)) {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			forEachStatement(node.body, visit, activity);
		}
	}
}

/**
 * {@link forEachStatement}, with each block's header line visited as a
 * statement of its own before the body, and a Do's `Loop While` line after
 * it ({@link blockHeaderStatements}, issue #233). For a rule that judges an
 * expression wherever it stands.
 */
export function forEachStatementWithHeaders(
	source: string,
	body: BodyNode[],
	visit: (stmt: LeafStatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (isLeafStatement(node)) {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			const { before, after } = blockHeaderStatements(source, node);
			if (before) {
				visit(before);
			}
			forEachStatementWithHeaders(source, node.body, visit, activity);
			if (after) {
				visit(after);
			}
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
) => ((stmt: LeafStatementNode) => void) | undefined;

/**
 * Runs every registered per-statement rule on ONE walk over the module's
 * active procedures and statements (audit #0). Each visitor sees procedures
 * and statements in exactly the order the rules' former private walks used:
 * active members in source order, `forEachStatement` within each body.
 */
export interface ProcedureWalkHooks {
	/** Called before a procedure's factories run (incremental attribution). */
	beforeMember?: (member: ProcedureNode) => void;
	/**
	 * When it returns true, the member's body iteration is skipped. Factories
	 * are still invoked so factory-level bookkeeping matches a full pass.
	 */
	skipBody?: (member: ProcedureNode) => boolean;
}

export function walkProcedureStatements(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	visitors: readonly ProcedureStatementVisitor[],
	hooks?: ProcedureWalkHooks,
	/**
	 * The module's text, and for each visitor whether it also takes block
	 * headers ({@link blockHeaderStatements}). Without it no visitor does.
	 */
	headers?: { source: string; takes: readonly boolean[] },
): void {
	const takesHeaders = headers?.takes ?? [];
	if (visitors.length === 0) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		hooks?.beforeMember?.(member);
		// Skipped before its visitors are built: an incremental pass walks the
		// edited procedure only, and asking every rule for a visitor for each
		// of the others cost a large module tens of milliseconds a keystroke.
		// A visitor is per member (cross-member state belongs in a run rule),
		// so a skipped one had nothing to contribute.
		if (hooks?.skipBody?.(member)) {
			continue;
		}
		const callbacks: Array<(stmt: LeafStatementNode) => void> = [];
		const headerCallbacks: Array<(stmt: LeafStatementNode) => void> = [];
		visitors.forEach((visitor, k) => {
			const callback = visitor(member);
			if (callback) {
				callbacks.push(callback);
				if (takesHeaders[k]) {
					headerCallbacks.push(callback);
				}
			}
		});
		if (callbacks.length === 0) {
			continue;
		}
		const header = (stmt: StatementNode | undefined): void => {
			if (stmt) {
				for (const callback of headerCallbacks) {
					callback(stmt);
				}
			}
		};
		const visit = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if (isInactiveNode(activity, node)) {
					continue;
				}
				if (isLeafStatement(node)) {
					for (const callback of callbacks) {
						callback(node);
					}
				} else if ('body' in node && Array.isArray(node.body)) {
					const { before, after } = headers && headerCallbacks.length > 0 ? blockHeaderStatements(headers.source, node) : {};
					header(before);
					// A block If's own line: `If 10 / d > 1 Then` evaluates its
					// condition as a statement does (issue #492). Its ElseIf lines
					// are statements of the body already.
					const opening = node.kind === 'IfBlock' && headers && headerCallbacks.length > 0 ? node.branches[0]?.headerSpan : undefined;
					header(opening ? { kind: 'Statement', span: opening, raw: headers!.source.slice(opening.start, opening.end) } : undefined);
					visit(node.body);
					header(after);
				}
			}
		};
		visit(member.body);
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

/**
 * Significant tokens of an expression the parser carried as its own string
 * (an If condition, a For Each source, an Enum member value, a parameter
 * default, a Const value). That text is not the module source, so it must
 * not go through `statementTokens`: the statement cache is keyed by source
 * string and holds two of them, and each distinct raw string sent there
 * evicted the module, so the next ordinary statement re-lexed the whole
 * module (issue #139, a 17x slowdown on real projects).
 */
export function rawExpressionTokens(text: string): VbaToken[] {
	// A `#` that opens a line is a directive to the lexer, and an expression
	// never opens one: `#12/31/9999#` is a date (issue #255). It is lexed
	// behind an `=` and moved back, so every offset stays the text's.
	if (!/^[ \t]*#/.test(text)) {
		return lexStatementTokens(text, { start: 0, end: text.length });
	}
	return lexStatementTokens(`=${text}`, { start: 0, end: text.length + 1 })
		.slice(1)
		.map((tok) => ({ ...tok, start: tok.start - 1, end: tok.end - 1 }));
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

export function stripHeaderBrackets(text: string): string {
	return text.startsWith('[') && text.endsWith(']')
		? text.slice(1, -1)
		: text;
}

/**
 * If the statement spanning `span` is a simple assignment to a bare identifier
 * (`name = ...` or `Let name = ...`), returns that identifier and its span;
 * otherwise undefined. `Set` (object) assignments and any left-hand side with a
 * `.` or `(` are excluded so only true scalar-name assignments are considered.
 */
/**
 * The spans a rule should read for one statement: the statement itself, plus
 * the statements a single-line `If` executes after `Then` and after `Else`.
 *
 * Opt-in rather than folded into {@link forEachStatement}, because rules
 * disagree about what a statement is. A rule that SCANS the text already sees
 * inside a single-line If and would double-report; a rule that reads it
 * STRUCTURALLY - first token is the assignment target, first token is the
 * callee - sees only `If`, and was blind to `If ok Then x = 1` (issue #46).
 * Only the structural ones call this.
 */
export function statementAndBranchSpans(stmt: LeafStatementNode): Span[] {
	const branches = stmt.kind === 'Statement' ? stmt.singleLineIfBranches : undefined;
	return branches ? [stmt.span, ...branches] : [stmt.span];
}

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
	// A name that SPELLS a keyword is still a name. The lexer classifies `Text`,
	// `Read` and `Type` as keywords, so requiring an `identifier` here hid every
	// assignment to a variable or Function named one of them - `Function text()`
	// assigning `text = 1` read as never assigning its own return (issue #46).
	// The `= ` that follows is what settles it: no VBA statement keyword is
	// followed by a bare `=` at statement start, and `Set`/`Let` are handled above.
	if (!nameTok || (nameTok.kind !== 'identifier' && nameTok.kind !== 'keyword')) {
		return undefined; // first token must be a name-like LHS
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
/**
 * Every tracked local the statement names whole - bare, not the statement's
 * own head, not a member access, not indexed - in an argument position: a
 * call statement's argument, an argument to a function inside an expression,
 * or an argument to a qualified member call. VBA passes by reference by
 * default, so the callee may have assigned or allocated the caller's variable
 * and its state is unknown from that point on (#70). A mention the callee
 * provably only reads is left out: the operand of `Is`, and the argument of
 * an intrinsic in `readOnlyIntrinsics`. The value is the first such mention's
 * absolute offset, so a rule can tell an access before the pass from one
 * after it within the same statement.
 */
export function localsNamedWhole(
	source: string,
	span: Span,
	tracked: ReadonlyMap<string, unknown>,
	readOnlyIntrinsics: ReadonlySet<string>,
): Map<string, number> {
	return trackedLocalsNamedWhole(
		statementTokensAfterLeadingLabel(source, span),
		span.start,
		(lower) => tracked.has(lower),
		readOnlyIntrinsics,
		undefined,
		// A procedure of the module that cannot change the argument keeps
		// what is known about it (issue #449).
		calleeKeepsArgument(source),
	);
}

/** The block's header line, with the lines a ` _` continues it onto. */
export function blockHeaderLineSpan(source: string, span: Span): Span {
	let nl = firstLineBreakAtOrAfter(source, span.start);
	while (nl >= 0 && nl <= span.end && endsInContinuation(source, span.start, nl)) {
		const next = source[nl] === '\r' && source[nl + 1] === '\n' ? nl + 2 : nl + 1;
		nl = firstLineBreakAtOrAfter(source, next);
	}
	if (nl < 0 || nl > span.end) {
		return span;
	}
	return { start: span.start, end: nl };
}

function endsInContinuation(source: string, start: number, nl: number): boolean {
	let i = nl - 1;
	while (i >= start && (source[i] === ' ' || source[i] === '\t')) {
		i--;
	}
	return i > start && source[i] === '_' && (source[i - 1] === ' ' || source[i - 1] === '\t');
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
