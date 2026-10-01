// Rule family: object-variable state (audit #0).
//
// Extracted verbatim from analyzeModule.ts: member access on unset object
// variables (straight-line Set tracking) and member access on known scalars.

import { precedesLeadingMemberDot, type MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	ForBlockNode,
	ModuleNode,
	ProcedureNode,
	Span,
	LeafStatementNode,
} from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import {
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import { walkBranchMergedBody, walkStraightLineBody } from '../dataflow';
import { procedureHasUnstructuredFlow } from '../../flow/procedureUnstructured';
import { statementLabelDeclarations, statementLabelReferences } from '../../flow/procedureLabels';
import { resolveExhaustiveMemberSurface } from '../rules/shared';
import {
	declaredTypeForSourceBinding,
	isKnownObjectAssignmentType,
	isKnownScalarType,
	normalizeType,
	objectLetAssignmentVerdict,
	returnAssignmentTypeFor,
	type SourceDeclaredType,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
	blockHeaderStatements,
	bareAssignmentTarget,
	forEachStatement,
	isInactiveNode,
	localsNamedWhole,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

/** Per-statement rule: rides the shared procedure-statement walk (audit #0). */
export function checkScalarMemberAccess(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			for (const hit of scalarMemberAccesses(
				source,
				stmt.span,
				env,
				(name) => declaredTypeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'memberReceiver',
				),
			)) {
				push(
					'scalarMemberAccess',
					`Member access on '${hit.name}' is invalid because it is declared as ${hit.asType}. This is a VBE compile error: ${hit.vbeError}.`,
					hit.span,
				);
			}
		};
	};
}

function scalarMemberAccesses(
	source: string,
	span: Span,
	env: ReadonlyMap<string, string>,
	resolveDeclaredType?: (name: string) => SourceDeclaredType,
): Array<{ name: string; asType: string; span: Span; vbeError: string }> {
	const toks = statementTokens(source, span);
	const out: Array<{ name: string; asType: string; span: Span; vbeError: string }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (toks[i + 1].rawText !== '.') {
			continue;
		}
		if (toks[i - 1]?.rawText === '.') {
			continue;
		}
		const name = tokenName(toks[i]);
		if (!name) {
			continue;
		}
		const declaredType = resolveDeclaredType?.(name);
		const asType = declaredType?.resolved
			? declaredType.asType
			: env.get(name.toLowerCase());
		const normalized = normalizeType(asType);
		if (!asType || !normalized || !isKnownScalarType(normalized)) {
			continue;
		}
		const memberName = toks[i + 2] ? tokenName(toks[i + 2]) : undefined;
		out.push({
			name,
			asType,
			vbeError: memberName ? 'Invalid qualifier' : 'Syntax error',
			span: { start: span.start + toks[i].start, end: span.start + toks[i + 1].end },
		});
	}
	return out;
}

interface LocalObjectVariable {
	name: string;
	asType: string;
	/**
	 * A Function's own result: Nothing until the function Sets it, so a Let
	 * into it raises 91 (issue #193). Only a Let reads it; inside the function
	 * its name with a dot or in a With is a recursive call.
	 */
	letOnly?: boolean;
}

type ObjectVariableState = 'unset' | 'set' | 'unknown';

export function checkObjectVariableNotSet(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const nothingFunctions = functionsReturningNothing(source, mod, memberCtx, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		if (nothingFunctions.size > 0) {
			forEachStatement(member.body, (stmt) => {
				for (const span of statementAndBranchSpans(stmt)) {
					for (const hit of nothingResultMemberAccess(statementTokens(source, span), nothingFunctions)) {
						push('objectVariableNotSet', hit.message, { start: span.start + hit.start, end: span.start + hit.end });
					}
				}
			}, activity);
		}
		checkGoToIntoWith(source, member, activity, push);
		for (const finding of objectStateWalk(source, member, symbols, memberCtx, activity).findings) {
			push(...finding);
		}
	}
}

/** Words that raise or end before a Function could return: Err.Raise, Error, End, Stop. */
const PREEMPTING_WORDS: ReadonlySet<string> = new Set(['raise', 'error', 'stop']);

/**
 * The Functions of the module, by lowercased name, that return an object
 * and never name their result: each returns Nothing (issue #240, measured
 * in Excel 16.0), so `F().Count` raises 91. A body that may raise or end
 * first is left out.
 */
function functionsReturningNothing(
	source: string,
	mod: ModuleNode,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
): Map<string, ProcedureNode> {
	const out = new Map<string, ProcedureNode>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || member.procKind !== 'Function' || !member.returnType || /\(\s*\)\s*$/.test(member.returnType)) {
			continue;
		}
		if (!isKnownObjectAssignmentType(member.returnType, memberCtx)) {
			continue;
		}
		const lower = member.name.toLowerCase();
		const body = statementTokens(source, { start: member.span.start, end: member.span.end });
		// The header names it once and `End Function` closes it.
		const named = body.filter((tok) => tokenName(tok)?.toLowerCase() === lower).length;
		const preempts = body.some((tok, i) => PREEMPTING_WORDS.has(tokenText(tok)) || (tokenText(tok) === 'end' && i > 0 && !['function', 'if', 'select', 'with', 'sub', 'property'].includes(tokenText(body[i + 1]))));
		if (named === 1 && !preempts) {
			out.set(lower, member);
		}
	}
	return out;
}

/** `F().Count` or `F.Count` on a Function that returns Nothing. Offsets are the statement's. */
function nothingResultMemberAccess(
	toks: readonly VbaToken[],
	functions: ReadonlyMap<string, ProcedureNode>,
): Array<{ start: number; end: number; message: string }> {
	const out: Array<{ start: number; end: number; message: string }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const fn = functions.get(tokenName(toks[i])?.toLowerCase() ?? '');
		if (!fn || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		let end = i;
		if (toks[i + 1].rawText === '(') {
			let depth = 0;
			for (let k = i + 1; k < toks.length; k++) {
				depth += toks[k].rawText === '(' ? 1 : toks[k].rawText === ')' ? -1 : 0;
				if (depth === 0) {
					end = k;
					break;
				}
			}
		} else if (fn.params.length > 0) {
			continue;
		}
		if (toks[end + 1]?.rawText !== '.' || !tokenName(toks[end + 2])) {
			continue;
		}
		out.push({
			start: toks[i].start,
			end: toks[end].end,
			message: `Function '${fn.name}' never sets its result, so it returns Nothing, and '.${toks[end + 2].rawText}' has no object to reach. This will raise Run-time error '91': Object variable or With block variable not set.`,
		});
	}
	return out;
}

/** What one procedure's object-state walk found, and the state at each Let. */
interface ObjectStateWalk {
	findings: Array<Parameters<PushFn>>;
	/** The state of the target at each bare Let into a tracked object, by the target's offset. */
	lets: Map<number, ObjectVariableState>;
}

// Keyed by the procedure node; the source, the activity and the member
// context must match too, since a parse is reused under another host.
const OBJECT_STATE_WALKS = new WeakMap<ProcedureNode, { source: string; activity: ConditionalActivityTracker | undefined; memberCtx: MemberCompletionContext; walk: ObjectStateWalk }>();

/**
 * Whether the object a Let assigns through at `offset` is provably set, or
 * provably Nothing, there (issue #193): set-required names 438 only for one
 * that holds an object, and leaves one still Nothing to object-variable-not-set.
 */
export function objectLetStateAt(
	source: string,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	offset: number,
): 'set' | 'unset' | 'unknown' {
	return objectStateWalk(source, member, symbols, memberCtx, activity).lets.get(offset) ?? 'unknown';
}

function objectStateWalk(
	source: string,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
): ObjectStateWalk {
	const cached = OBJECT_STATE_WALKS.get(member);
	if (cached && cached.source === source && cached.activity === activity && cached.memberCtx === memberCtx) {
		return cached.walk;
	}
	const walk: ObjectStateWalk = { findings: [], lets: new Map() };
	const push: PushFn = (...finding) => {
		walk.findings.push(finding);
	};
	walkObjectState(source, member, symbols, memberCtx, activity, push, walk.lets);
	OBJECT_STATE_WALKS.set(member, { source, activity, memberCtx, walk });
	return walk;
}

function walkObjectState(
	source: string,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	lets: Map<number, ObjectVariableState>,
): void {
	const locals = localObjectVariablesFor(symbols, member, memberCtx);
	if (locals.size === 0) {
		return;
	}
	const state = new Map<string, ObjectVariableState>();
	for (const key of locals.keys()) {
		state.set(key, 'unset');
	}
	// The locals some statement anywhere in the procedure Sets: a `GoSub`
	// may run any of those statements before control comes back (issue
	// #108), so after it none of them is provably still Nothing.
	const setAnywhere = new Set<string>();
	forEachStatement(member.body, (stmt) => {
		for (const span of statementAndBranchSpans(stmt)) {
			const lower = setAssignmentTarget(source, span)?.name.toLowerCase();
			if (lower && locals.has(lower)) {
				setAnywhere.add(lower);
			}
		}
	}, activity);
	const walk = procedureHasUnstructuredFlow(source, member, activity)
		? walkStraightLineBody
		: walkBranchMergedBody;
	walk(source, member.body, (node) => isInactiveNode(activity, node), {
		onStatement: (stmt) =>
			checkObjectVariableNotSetStatement(source, stmt, locals, state, setAnywhere, memberCtx, push, lets),
		onBlock: (node) => {
			// The header runs as the block is entered, with the state as it
			// stands: `For i = 1 To c.Count`, `Select Case c.Count` (issue #233).
			if (node.kind === 'SelectBlock' || node.kind === 'DoBlock' || node.kind === 'WhileBlock' || (node.kind === 'ForBlock' && !node.each)) {
				const { before } = blockHeaderStatements(source, node);
				if (before) {
					checkObjectVariableNotSetStatement(source, before, locals, state, setAnywhere, memberCtx, push, lets);
				}
			}
			// A For Each that runs to its end leaves the control variable
			// Nothing, so an access after the loop is right to report. One
			// the body can leave early - Exit For, or a GoTo out of it -
			// leaves it on the current element, so nothing is proven
			// (issue #108: `Exit For` on the first sheet, then `ws.Name`).
			if (node.kind === 'ForBlock') {
				// `For Each x In c` with c still Nothing raises 424, not 91:
				// the loop asks the collection for its enumerator (issue #121).
				const over = node.each ? node.sourceExpression?.trim().toLowerCase() : undefined;
				if (over && locals.has(over) && !locals.get(over)!.letOnly && state.get(over) === 'unset' && node.sourceExpressionSpan) {
					push(
						'objectVariableNotSet',
						`Object variable '${locals.get(over)!.name}' is Nothing when For Each asks it for its elements. This will raise Run-time error '424': Object required.`,
						node.sourceExpressionSpan,
					);
				}
				const lower = node.controlVariable?.toLowerCase();
				if (node.each && lower && locals.has(lower) && state.get(lower) === 'unset'
					&& bodyCanLeaveLoop(source, node, activity)) {
					state.set(lower, 'unknown');
				}
				return;
			}
			if (node.kind !== 'WithBlock') {
				return;
			}
			const receiver = unsetWithObjectReceiver(source, node.span, locals, state);
			if (receiver) {
				push(
					'objectVariableNotSet',
					`Object variable '${receiver.name}' is Nothing before With member access. This will raise Run-time error '91': Object variable or With block variable not set.`,
					receiver.span,
				);
			}
		},
		touchesInStatement: (stmt) => {
			const touched = new Set(
				localsNamedWhole(source, stmt.span, locals, OBJECT_READ_ONLY_INTRINSICS).keys(),
			);
			// A single-line If's branches Set too.
			for (const span of statementAndBranchSpans(stmt)) {
				const lower = setAssignmentTarget(source, span)?.name.toLowerCase();
				if (lower && locals.has(lower)) {
					touched.add(lower);
				}
			}
			return touched;
		},
		demoteToUnknown: (lower) => {
			if (state.get(lower) === 'unset') {
				state.set(lower, 'unknown');
			}
		},
		snapshotState: () => new Map(state),
		restoreState: (snapshot) => {
			state.clear();
			for (const [key, value] of snapshot) {
				state.set(key, value as ObjectVariableState);
			}
		},
		setState: (key, value) => state.set(key, value as ObjectVariableState),
		lattice: { init: 'unset', good: 'set', unknown: 'unknown' },
	});
}

/**
 * `GoTo L` from outside a With block to a label inside it skips the With
 * statement, so the With has no object: the first leading-dot member after
 * the label raises 91 (issue #184, measured in Excel 16.0). A GoTo inside the
 * same With runs, and so does one into a For loop.
 */
function checkGoToIntoWith(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const hasWith = (list: readonly BodyNode[]): boolean => list.some((node) =>
		node.kind === 'WithBlock' || ('body' in node && Array.isArray(node.body) && hasWith(node.body as BodyNode[])));
	if (!hasWith(proc.body)) {
		return;
	}
	const labels = new Map<string, { withs: readonly BodyNode[]; access: string | undefined }>();
	const jumps: Array<{ key: string; text: string; span: Span; withs: readonly BodyNode[] }> = [];
	const visit = (list: readonly BodyNode[], withs: readonly BodyNode[]): void => {
		for (let i = 0; i < list.length; i++) {
			const node = list[i];
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (isLeafStatement(node)) {
				for (const label of withs.length > 0 ? statementLabelDeclarations(source, node.span) : []) {
					if (!labels.has(label.key)) {
						labels.set(label.key, { withs, access: firstLeadingDotMember(source, list, i, activity) });
					}
				}
				for (const ref of statementLabelReferences(source, node.span)) {
					if (ref.statementKind === 'goto') {
						jumps.push({ key: ref.key, text: ref.text, span: ref.span, withs });
					}
				}
				continue;
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[], node.kind === 'WithBlock' ? [...withs, node] : withs);
			}
		}
	};
	visit(proc.body, []);
	for (const jump of jumps) {
		const target = labels.get(jump.key);
		if (target?.access && target.withs.some((block) => !jump.withs.includes(block))) {
			push(
				'objectVariableNotSet',
				`GoTo ${jump.text} jumps into a With block past its With statement, so '${target.access}' after the label has no object. This will raise Run-time error '91': Object variable or With block variable not set.`,
				jump.span,
			);
		}
	}
}

/**
 * The first leading-dot member (`.Add`) that runs from `list[from]` on, in
 * the statements that follow in a straight line. A block may not run, and an
 * Exit, GoTo or Return leaves, so either ends the search.
 */
function firstLeadingDotMember(
	source: string,
	list: readonly BodyNode[],
	from: number,
	activity: ConditionalActivityTracker | undefined,
): string | undefined {
	for (let j = from; j < list.length; j++) {
		const node = list[j];
		if (isInactiveNode(activity, node) || node.kind === 'VariableGroup') {
			continue;
		}
		if (!isLeafStatement(node)) {
			return undefined;
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span);
		const head = tokenText(toks[0]);
		if (j > from && (head === 'elseif' || head === 'else' || head === 'case')) {
			return undefined;
		}
		// A one-line If always runs its condition, and its branches maybe.
		const then = node.kind === 'Statement' && node.singleLineIfBranches ? toks.findIndex((tok) => tokenText(tok) === 'then') : -1;
		const limit = then >= 0 ? then : toks.length;
		for (let k = 0; k < limit; k++) {
			if (toks[k].rawText === '.' && tokenName(toks[k + 1]) && (k === 0 || precedesLeadingMemberDot(toks[k - 1]))) {
				return `.${toks[k + 1].rawText}`;
			}
		}
		if (then >= 0 || head === 'exit' || head === 'goto' || head === 'return' || head === 'resume' || head === 'end') {
			return undefined;
		}
	}
	return undefined;
}

/**
 * Whether the loop body can leave the loop before it ends: an `Exit For` at
 * its own depth (one inside a nested For leaves that one), or any `GoTo`.
 */
function bodyCanLeaveLoop(
	source: string,
	loop: ForBlockNode,
	activity: ConditionalActivityTracker | undefined,
): boolean {
	const visit = (body: readonly BodyNode[]): boolean => {
		for (const node of body) {
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (node.kind === 'ForBlock') {
				continue; // its Exit For is its own
			}
			if ('body' in node && Array.isArray(node.body)) {
				if (visit(node.body as BodyNode[])) {
					return true;
				}
				continue;
			}
			for (const span of statementAndBranchSpans(node as LeafStatementNode)) {
				const toks = statementTokensAfterLeadingLabel(source, span);
				const head = tokenText(toks[0]);
				if ((head === 'exit' && tokenText(toks[1]) === 'for') || head === 'goto') {
					return true;
				}
			}
		}
		return false;
	};
	return visit(loop.body);
}

/**
 * The tracked names a single-line If's condition guards: `Not d Is Nothing`
 * guards the Then arm, `d Is Nothing` the Else arm (issue #108: the block
 * form already read the guard, the one-line form did not).
 */
function nothingGuardNames(condition: readonly VbaToken[]): { thenArm: Set<string>; elseArm: Set<string> } {
	const thenArm = new Set<string>();
	const elseArm = new Set<string>();
	for (let i = 0; i + 2 < condition.length; i++) {
		if (tokenText(condition[i + 1]) !== 'is' || tokenText(condition[i + 2]) !== 'nothing') {
			continue;
		}
		const name = tokenName(condition[i])?.toLowerCase();
		if (!name) {
			continue;
		}
		if (tokenText(condition[i - 1]) === 'not') {
			thenArm.add(name);
		} else {
			elseArm.add(name);
		}
	}
	return { thenArm, elseArm };
}

function checkObjectVariableNotSetStatement(
	source: string,
	stmt: LeafStatementNode,
	locals: ReadonlyMap<string, LocalObjectVariable>,
	state: Map<string, ObjectVariableState>,
	setAnywhere: ReadonlySet<string>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
	lets: Map<number, ObjectVariableState>,
): void {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	const head = tokenText(toks[0]);
	// `GoSub Label` runs the subroutine, which may Set any of the locals,
	// before the statement after it (issue #108).
	if (head === 'gosub' || (head === 'on' && toks.some((tok) => tokenText(tok) === 'gosub'))) {
		for (const lower of setAnywhere) {
			if (state.get(lower) === 'unset') {
				state.set(lower, 'unknown');
			}
		}
		return;
	}
	// The arms of a single-line If and what its condition proves about them.
	const branches = statementAndBranchSpans(stmt);
	let guards = { thenArm: new Set<string>(), elseArm: new Set<string>() };
	if (head === 'if' && branches.length > 1) {
		const thenIndex = toks.findIndex((tok, index) => index > 0 && tokenText(tok) === 'then');
		if (thenIndex > 0) {
			guards = nothingGuardNames(toks.slice(1, thenIndex));
		}
	}
	const guardedAt = (name: string, offset: number): boolean => {
		const within = (span: Span | undefined): boolean =>
			span !== undefined && offset >= span.start && offset < span.end;
		return (guards.thenArm.has(name) && within(branches[1]))
			|| (guards.elseArm.has(name) && within(branches[2]));
	};
	// A bare `obj = value` is a Let through the object's default member
	// (issue #107), which needs an object to reach: on a variable still
	// Nothing it raises 91, the same as a member access would.
	for (const span of branches) {
		const let_ = bareAssignmentTarget(source, span);
		const lower = let_?.name.toLowerCase();
		if (!let_ || !lower || !locals.has(lower)) {
			continue;
		}
		const letState = guardedAt(lower, let_.span.start) ? 'unknown' : state.get(lower) ?? 'unknown';
		lets.set(let_.span.start, letState);
		// A type with no default member for the Let, or one that needs an
		// argument, is set-required's to report, with the 91 when it is still
		// Nothing (issue #193): the fix there is the Set.
		const verdict = objectLetAssignmentVerdict(locals.get(lower)!.asType, memberCtx);
		if (letState === 'unset' && verdict !== 'noDefault' && verdict !== 'argument') {
			const what = locals.get(lower)!.letOnly ? `The result '${let_.name}'` : `Object variable '${let_.name}'`;
			push(
				'objectVariableNotSet',
				`${what} is Nothing before the default-member assignment. This will raise Run-time error '91': Object variable or With block variable not set.`,
				let_.span,
			);
		}
	}
	const passedWhole = localsNamedWhole(source, stmt.span, locals, OBJECT_READ_ONLY_INTRINSICS);
	for (const hit of unsetObjectMemberAccesses(source, stmt.span, locals, state, memberCtx)) {
		// An access after a whole pass in the same statement, as in
		// `If TryGet(obj) Then obj.Name`, runs after the callee had its chance
		// to Set it. One before the pass, as in `Load(obj.Name)`, does not.
		const passAt = passedWhole.get(hit.name.toLowerCase());
		if (passAt !== undefined && hit.span.start > passAt) {
			continue;
		}
		if (guardedAt(hit.name.toLowerCase(), hit.span.start)) {
			continue;
		}
		push(
			'objectVariableNotSet',
			`Object variable '${hit.name}' is Nothing before member access. This will raise Run-time error '91': Object variable or With block variable not set.`,
			hit.span,
		);
	}
	const target = setAssignmentTarget(source, stmt.span);
	if (target) {
		const lower = target.name.toLowerCase();
		if (locals.has(lower)) {
			state.set(lower, setAssignmentValueIsNothing(target) ? 'unset' : 'set');
			return;
		}
	}
	for (const lower of passedWhole.keys()) {
		if (state.get(lower) === 'unset') {
			state.set(lower, 'unknown');
		}
	}
	// A Set in a single-line If's branch runs on one path only, so it moves an
	// unset object to 'unknown' the way a block If without Else does, not to
	// 'set' - as unallocated-dynamic-array-access reads a conditional ReDim.
	for (const branch of statementAndBranchSpans(stmt).slice(1)) {
		const lower = setAssignmentTarget(source, branch)?.name.toLowerCase();
		if (lower && locals.has(lower) && state.get(lower) === 'unset') {
			state.set(lower, 'unknown');
		}
	}
}

/** Intrinsics that read an object argument and never Set it. */
const OBJECT_READ_ONLY_INTRINSICS: ReadonlySet<string> = new Set([
	'typename', 'vartype', 'isobject', 'isnull', 'isempty', 'ismissing', 'objptr',
]);

function localObjectVariablesFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	memberCtx: MemberCompletionContext,
): Map<string, LocalObjectVariable> {
	const out = new Map<string, LocalObjectVariable>();
	const procSym = procedureSymbolFor(symbols, proc);
	for (const child of procSym?.children ?? []) {
		if (
			child.kind !== 'localVariable' ||
			child.visibility === 'Static' ||
			child.isArray === true ||
			// `Dim x As New Invoice` is instantiated on ANY access, including
			// the first one and including after `Set x = Nothing`, so it can
			// never be Nothing when a member is touched. Tracking it produced
			// error 91 warnings on code that runs.
			child.isAutoInstantiated === true ||
			!isKnownObjectAssignmentType(child.asType, memberCtx) ||
			!child.asType
		) {
			continue;
		}
		out.set(child.name.toLowerCase(), { name: child.name, asType: child.asType });
	}
	const result = returnAssignmentTypeFor(proc);
	if (result && isKnownObjectAssignmentType(result, memberCtx) && !out.has(proc.name.toLowerCase())) {
		out.set(proc.name.toLowerCase(), { name: proc.name, asType: result, letOnly: true });
	}
	return out;
}

function unsetObjectMemberAccesses(
	source: string,
	span: Span,
	locals: ReadonlyMap<string, LocalObjectVariable>,
	state: ReadonlyMap<string, ObjectVariableState>,
	memberCtx: MemberCompletionContext,
): Array<{ name: string; span: Span }> {
	const toks = statementTokens(source, span);
	const out: Array<{ name: string; span: Span }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (toks[i + 1].rawText !== '.' || toks[i - 1]?.rawText === '.') {
			continue;
		}
		const name = tokenName(toks[i]);
		if (!name) {
			continue;
		}
		const lower = name.toLowerCase();
		if (!locals.has(lower) || locals.get(lower)!.letOnly || state.get(lower) !== 'unset') {
			continue;
		}
		const member = toks[i + 2] ? tokenName(toks[i + 2]) : undefined;
		if (
			member &&
			hasDefiniteMissingMember(source, span.start + toks[i + 1].end, member, memberCtx)
		) {
			continue;
		}
		out.push({
			name,
			span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
		});
	}
	return out;
}

function hasDefiniteMissingMember(
	source: string,
	dotEndOffset: number,
	memberName: string,
	memberCtx: MemberCompletionContext,
): boolean {
	const surface = resolveExhaustiveMemberSurface(source, dotEndOffset, memberCtx);
	return surface !== undefined && !surface.hasMember(memberName);
}

function unsetWithObjectReceiver(
	source: string,
	span: Span,
	locals: ReadonlyMap<string, LocalObjectVariable>,
	state: ReadonlyMap<string, ObjectVariableState>,
): { name: string; span: Span } | undefined {
	const header = blockHeaderLineSpan(source, span);
	const toks = statementTokensAfterLeadingLabel(source, header);
	if (tokenText(toks[0]) !== 'with' || toks.length !== 2) {
		return undefined;
	}
	const name = tokenName(toks[1]);
	if (!name) {
		return undefined;
	}
	const lower = name.toLowerCase();
	if (!locals.has(lower) || locals.get(lower)!.letOnly || state.get(lower) !== 'unset') {
		return undefined;
	}
	return {
		name,
		span: { start: header.start + toks[1].start, end: header.start + toks[1].end },
	};
}

function setAssignmentValueIsNothing(
	target: { valueTokens: readonly VbaToken[] },
): boolean {
	const toks = target.valueTokens.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
	return toks.length === 1 && tokenText(toks[0]) === 'nothing';
}
