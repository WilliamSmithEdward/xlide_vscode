// Rule family: object-variable state (audit #0).
//
// Extracted verbatim from analyzeModule.ts: member access on unset object
// variables (straight-line Set tracking) and member access on known scalars.

import type { MemberCompletionContext } from '../../completion/memberAccess';
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
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import {
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import { walkBranchMergedBody, walkStraightLineBody } from '../dataflow';
import { procedureHasUnstructuredFlow } from '../../flow/procedureUnstructured';
import { resolveExhaustiveMemberSurface } from '../rules/shared';
import {
	declaredTypeForSourceBinding,
	isKnownObjectAssignmentType,
	isKnownScalarType,
	normalizeType,
	type SourceDeclaredType,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
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
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const locals = localObjectVariablesFor(symbols, member, memberCtx);
		if (locals.size === 0) {
			continue;
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
		walk(member.body, (node) => isInactiveNode(activity, node), {
			onStatement: (stmt) =>
				checkObjectVariableNotSetStatement(source, stmt, locals, state, setAnywhere, memberCtx, push),
			onBlock: (node) => {
				// A For Each that runs to its end leaves the control variable
				// Nothing, so an access after the loop is right to report. One
				// the body can leave early - Exit For, or a GoTo out of it -
				// leaves it on the current element, so nothing is proven
				// (issue #108: `Exit For` on the first sheet, then `ws.Name`).
				if (node.kind === 'ForBlock') {
					// `For Each x In c` with c still Nothing raises 424, not 91:
					// the loop asks the collection for its enumerator (issue #121).
					const over = node.each ? node.sourceExpression?.trim().toLowerCase() : undefined;
					if (over && locals.has(over) && state.get(over) === 'unset' && node.sourceExpressionSpan) {
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
		if (let_ && lower && locals.has(lower) && state.get(lower) === 'unset'
			&& !guardedAt(lower, let_.span.start)) {
			push(
				'objectVariableNotSet',
				`Object variable '${let_.name}' is Nothing before the default-member assignment. This will raise Run-time error '91': Object variable or With block variable not set.`,
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
		if (!locals.has(lower) || state.get(lower) !== 'unset') {
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
	return (
		surface !== undefined &&
		!surface.members.some(
			(candidate) => candidate.name.toLowerCase() === memberName.toLowerCase(),
		)
	);
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
	if (!locals.has(lower) || state.get(lower) !== 'unset') {
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
