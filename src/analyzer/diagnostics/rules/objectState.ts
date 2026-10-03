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
import { conditionOperands } from '../conditionOperands';
import { walkBranchMergedBody, walkStraightLineBody } from '../dataflow';
import { untouchedModuleVariablesIn } from '../moduleState';
import { procedureHasUnstructuredFlow } from '../../flow/procedureUnstructured';
import { statementLabelDeclarations, statementLabelReferences } from '../../flow/procedureLabels';
import { builtinNameBefore, resolveExhaustiveMemberSurface, ONE_VALUE_BUILTINS } from '../rules/shared';
import {
	declaredTypeForSourceBinding,
	defTypeOf,
	isKnownObjectAssignmentType,
	sourceIdentifierBinding,
	isKnownScalarType,
	normalizeType,
	objectHoldingDefault,
	objectLetAssignmentVerdict,
	returnAssignmentTypeFor,
	type SourceDeclaredType,
	typeEnvironmentFor,
	unreachableStatementsIn,
} from '../typeInference';
import { conditionValue } from '../conditionValue';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import {
	activeModuleMembers,
	blockHeaderLineSpan,
	blockHeaderStatements,
	bareAssignmentTarget,
	forEachStatement,
	isInactiveNode,
	localsNamedWhole,
	matchParenFrom,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	type ProcedureStatementVisitor,
} from '../walker';

/** The operators that read an object's default member as an operand. */
const OPERAND_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>', '+', '-', '*', '/', '\\', '&', '^']);

/** Per-statement rule: rides the shared procedure-statement walk (audit #0). */
export function checkScalarMemberAccess(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
	memberCtx: MemberCompletionContext = {},
): ProcedureStatementVisitor {
	// The project's other standard modules: `Foo.Foo()` names module Foo
	// before its Function Foo, so the Function's Long is no receiver (issue #403).
	const otherModules = new Set((memberCtx.projectClassMembers ?? [])
		.filter((type) => type.kind === 'standardModule' && type.name.toLowerCase() !== symbols.moduleName.toLowerCase())
		.map((type) => type.name.toLowerCase()));
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			for (const hit of scalarMemberAccesses(
				source,
				stmt.span,
				env,
				(name) => otherModules.has(name.toLowerCase()) && sourceIdentifierBinding(symbols, procSym, projectVisibleSymbols, name, 'memberReceiver').scope === 'project'
					? { resolved: true }
					: declaredTypeForSourceBinding(
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
	/** A Variant: Nothing only once `Set v = Nothing` (issue #343). Only member reads are judged. */
	variant?: boolean;
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
					for (const hit of nothingResultMemberAccess(source, statementTokens(source, span), nothingFunctions)) {
						push('objectVariableNotSet', hit.message, { start: span.start + hit.start, end: span.start + hit.end });
					}
				}
			}, activity);
		}
		checkGoToIntoWith(source, member, activity, push);
		checkForEachOverEmptyObjectArray(source, member, symbols, memberCtx, activity, push);
		// A module variable nothing ever sets is Nothing in every procedure (issue #241).
		const unset = [...untouchedModuleVariablesIn(source, symbols, member)].filter(([, variable]) =>
			variable.asType !== undefined && isKnownObjectAssignmentType(variable.asType, memberCtx));
		if (unset.length > 0) {
			const objects = new Map(unset);
			forEachStatement(member.body, (stmt) => {
				for (const span of statementAndBranchSpans(stmt)) {
					const toks = statementTokens(source, span);
					for (let i = 0; i + 2 < toks.length; i++) {
						const variable = objects.get(tokenName(toks[i])?.toLowerCase() ?? '');
						if (!variable || toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!' || (toks[i + 1].rawText !== '.' && toks[i + 1].rawText !== '!') || !tokenName(toks[i + 2])) {
							continue;
						}
						const scope = variable.visibility === 'Public' || variable.visibility === 'Global' ? 'the project' : 'this module';
						push(
							'objectVariableNotSet',
							`Object variable '${toks[i].rawText}' is never set anywhere in ${scope}, so it is Nothing here. This will raise Run-time error '91': Object variable or With block variable not set.`,
							{ start: span.start + toks[i].start, end: span.start + toks[i].end },
						);
					}
				}
			}, activity);
		}
		for (const finding of objectStateWalk(source, mod, member, symbols, memberCtx, activity).findings) {
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
		// The header names it once and `End Function` closes it. `Set F =
		// Nothing` names it and still returns Nothing (issue #343).
		const named = body.filter((tok, i) => tokenName(tok)?.toLowerCase() === lower
			&& !(tokenText(body[i - 1]) === 'set' && body[i + 1]?.rawText === '=' && tokenText(body[i + 2]) === 'nothing')).length;
		const preempts = body.some((tok, i) => PREEMPTING_WORDS.has(tokenText(tok)) || (tokenText(tok) === 'end' && i > 0 && !['function', 'if', 'select', 'with', 'sub', 'property'].includes(tokenText(body[i + 1]))));
		if (named === 1 && !preempts) {
			out.set(lower, member);
		}
	}
	return out;
}

/** `F().Count` or `F.Count` on a Function that returns Nothing. Offsets are the statement's. */
function nothingResultMemberAccess(
	source: string,
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
		const setsNothing = fn.body.length > 0 && new RegExp(`\\bset\\s+${fn.name}\\s*=\\s*nothing\\b`, 'i').test(source.slice(fn.span.start, fn.span.end));
		out.push({
			start: toks[i].start,
			end: toks[end].end,
			message: `Function '${fn.name}' ${setsNothing ? 'sets its result to Nothing' : 'never sets its result, so it returns Nothing'}, and '.${toks[end + 2].rawText}' has no object to reach. This will raise Run-time error '91': Object variable or With block variable not set.`,
		});
	}
	return out;
}

/** What the walk needs to know about the module's other procedures (issue #343). */
interface ModuleObjectFacts {
	/** The Functions that return Nothing, by lowercased name. */
	nothingFunctions: ReadonlyMap<string, ProcedureNode>;
	/**
	 * The procedures whose object parameter's first use is a member read, by
	 * lowercased name: for each such parameter's position, the read as written,
	 * `c.Count`. Passed Nothing, the procedure raises 91 there.
	 */
	memberFirst: ReadonlyMap<string, ReadonlyMap<number, string>>;
}

const MODULE_OBJECT_FACTS = new WeakMap<ModuleNode, { source: string; activity: ConditionalActivityTracker | undefined; memberCtx: MemberCompletionContext; facts: ModuleObjectFacts }>();

function moduleObjectFacts(
	source: string,
	mod: ModuleNode,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
): ModuleObjectFacts {
	const cached = MODULE_OBJECT_FACTS.get(mod);
	if (cached && cached.source === source && cached.activity === activity && cached.memberCtx === memberCtx) {
		return cached.facts;
	}
	const memberFirst = new Map<string, Map<number, string>>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || member.procKind === 'PropertyLet' || member.procKind === 'PropertySet') {
			continue;
		}
		const reads = new Map<number, string>();
		member.params.forEach((param, k) => {
			if (!param.paramArray && !param.isArray && param.asType && isKnownObjectAssignmentType(param.asType, memberCtx)) {
				const read = firstUseMemberRead(source, member, param.name.toLowerCase(), activity);
				if (read) {
					reads.set(k, read);
				}
			}
		});
		if (reads.size > 0) {
			memberFirst.set(member.name.toLowerCase(), reads);
		}
	}
	const facts = { nothingFunctions: functionsReturningNothing(source, mod, memberCtx, activity), memberFirst };
	MODULE_OBJECT_FACTS.set(mod, { source, activity, memberCtx, facts });
	return facts;
}

/** Statement heads that leave, jump, raise or change error handling. */
const STRAIGHT_LINE_ENDS: ReadonlySet<string> = new Set(['on', 'resume', 'gosub', 'goto', 'exit', 'end', 'stop', 'return', 'error', 'err']);

/**
 * The member read, `c.Count`, when the first statement of the procedure to
 * name `lower` reads a member of it, and every statement before that runs
 * in a straight line: no block, label, On Error or single-line If.
 */
function firstUseMemberRead(
	source: string,
	member: ProcedureNode,
	lower: string,
	activity: ConditionalActivityTracker | undefined,
): string | undefined {
	for (const node of member.body) {
		if (isInactiveNode(activity, node) || node.kind === 'VariableGroup') {
			continue;
		}
		if (!isLeafStatement(node) || statementAndBranchSpans(node).length > 1 || statementLabelDeclarations(source, node.span).length > 0) {
			return undefined;
		}
		const toks = statementTokensAfterLeadingLabel(source, node.span).filter((tok) => tok.kind !== 'comment');
		const head = tokenText(toks[0]);
		if (STRAIGHT_LINE_ENDS.has(head)) {
			return undefined;
		}
		const at = toks.findIndex((tok, i) => tokenName(tok)?.toLowerCase() === lower && toks[i - 1]?.rawText !== '.' && toks[i - 1]?.rawText !== '!');
		if (at < 0) {
			continue;
		}
		const name = tokenName(toks[at + 2]);
		return head !== 'set' && toks[at + 1]?.rawText === '.' && name ? `${toks[at].rawText}.${name}` : undefined;
	}
	return undefined;
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
	mod: ModuleNode,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	offset: number,
): 'set' | 'unset' | 'unknown' {
	return objectStateWalk(source, mod, member, symbols, memberCtx, activity).lets.get(offset) ?? 'unknown';
}

function objectStateWalk(
	source: string,
	mod: ModuleNode,
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
	walkObjectState(source, moduleObjectFacts(source, mod, memberCtx, activity), member, symbols, memberCtx, activity, push, walk.lets);
	OBJECT_STATE_WALKS.set(member, { source, activity, memberCtx, walk });
	return walk;
}

function walkObjectState(
	source: string,
	facts: ModuleObjectFacts,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	lets: Map<number, ObjectVariableState>,
): void {
	const locals = localObjectVariablesFor(source, symbols, member, memberCtx);
	const elements = objectArrayElements(source, symbols, member, memberCtx, activity);
	if (locals.size === 0 && elements.keys.size === 0) {
		return;
	}
	const state = new Map<string, ObjectVariableState>();
	for (const key of locals.keys()) {
		// A Variant starts Empty, which is no object and not Nothing.
		state.set(key, locals.get(key)!.variant ? 'unknown' : 'unset');
	}
	// Each element of a fixed array of objects is Nothing until Set (issue
	// #489); a dynamic one's are, once ReDim allocates them.
	for (const [key, element] of elements.keys) {
		state.set(key, elements.arrays.get(element.array)!.fixed ? 'unset' : 'unknown');
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
	// The GoTo-following walk runs the body until its labels settle, and
	// reports on its last run (issue #271).
	let silent = false;
	const report: PushFn = (...finding) => {
		if (!silent) {
			push(...finding);
		}
	};
	const walk = procedureHasUnstructuredFlow(source, member, activity)
		? walkStraightLineBody
		: walkBranchMergedBody;
	// The For Each loops with no way out but the end.
	const nothingAfter = new Set<BodyNode>();
	// A statement a known guard keeps from running (issue #273).
	const unreachable = unreachableStatementsIn(source, member, symbols, activity);
	walk(source, member.body, (node) => isInactiveNode(activity, node) || unreachable.has(node), {
		onStatement: (stmt) =>
			checkObjectVariableNotSetStatement(source, stmt, locals, state, setAnywhere, memberCtx, report, lets, facts, elements),
		onBlock: (node) => {
			// The header runs as the block is entered, with the state as it
			// stands: `For i = 1 To c.Count`, `Select Case c.Count` (issue #233).
			if (node.kind === 'SelectBlock' || node.kind === 'DoBlock' || node.kind === 'WhileBlock' || (node.kind === 'ForBlock' && !node.each)) {
				const { before, after } = blockHeaderStatements(source, node);
				if (before) {
					checkObjectVariableNotSetStatement(source, before, locals, state, setAnywhere, memberCtx, report, lets, facts, elements);
				}
				// `Loop Until x` reads x after the body, with what it entered with
				// when the body never names x (issue #424). Any local the line
				// reads counts, `c` of `Loop While c.Count < 1` too (issue #560).
				if (after && node.kind === 'DoBlock') {
					const inBody = source.slice(before?.span.end ?? node.span.start, after.span.start).toLowerCase();
					const afterToks = statementTokensAfterLeadingLabel(source, after.span);
					const named = afterToks.some((tok) => {
						const lower = tokenName(tok)?.toLowerCase();
						return lower !== undefined && locals.has(lower) && new RegExp(`\\b${lower}\\b`).test(inBody);
					});
					if (!named) {
						checkObjectVariableNotSetStatement(source, after, locals, state, setAnywhere, memberCtx, report, lets, facts);
					}
				}
			}
			// A block If's own line and its ElseIf lines read their conditions
			// as the block is entered (issue #424).
			if (node.kind === 'IfBlock') {
				for (const branch of node.branches) {
					if (branch.branchKind !== 'else') {
						checkObjectVariableNotSetStatement(source, { kind: 'Statement', span: branch.headerSpan, raw: source.slice(branch.headerSpan.start, branch.headerSpan.end) }, locals, state, setAnywhere, memberCtx, report, lets, facts);
					}
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
				if (over && locals.has(over) && !locals.get(over)!.letOnly && !locals.get(over)!.variant && state.get(over) === 'unset' && node.sourceExpressionSpan) {
					report(
						'objectVariableNotSet',
						`Object variable '${locals.get(over)!.name}' is Nothing when For Each asks it for its elements. This will raise Run-time error '424': Object required.`,
						node.sourceExpressionSpan,
					);
				}
				const lower = node.controlVariable?.toLowerCase();
				if (node.each && lower && locals.has(lower) && !locals.get(lower)!.letOnly) {
					if (!bodyCanLeaveLoop(source, node, activity) && !locals.get(lower)!.variant) {
						nothingAfter.add(node);
					} else if (state.get(lower) === 'unset') {
						state.set(lower, 'unknown');
					}
				}
				return;
			}
			if (node.kind !== 'WithBlock') {
				return;
			}
			const receiver = unsetWithObjectReceiver(source, node.span, locals, state);
			if (receiver) {
				report(
					'objectVariableNotSet',
					`Object variable '${receiver.name}' is Nothing before With member access. This will raise Run-time error '91': Object variable or With block variable not set.`,
					receiver.span,
				);
			}
		},
		// A For Each that ends leaves its control variable Nothing, over an
		// empty collection too (issue #336, measured in Excel 16.0).
		afterBlock: (node) => {
			const lower = node.kind === 'ForBlock' ? node.controlVariable?.toLowerCase() : undefined;
			if (lower && nothingAfter.has(node)) {
				state.set(lower, 'unset');
			}
		},
		touchesInStatement: (stmt) => {
			const touched = new Set(
				localsNamedWhole(source, stmt.span, locals, OBJECT_READ_ONLY_INTRINSICS).keys(),
			);
			for (const span of statementAndBranchSpans(stmt)) {
				for (const key of elementTouches(statementTokensAfterLeadingLabel(source, span), elements)) {
					touched.add(key);
				}
			}
			// A single-line If's branches Set too. A Let gives a Variant a value.
			for (const span of statementAndBranchSpans(stmt)) {
				const lower = setAssignmentTarget(source, span)?.name.toLowerCase();
				if (lower && locals.has(lower)) {
					touched.add(lower);
				}
				const let_ = bareAssignmentTarget(source, span)?.name.toLowerCase();
				if (let_ && locals.get(let_)?.variant) {
					touched.add(let_);
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
		// A local never Set is Nothing: `If c Is Nothing Then Exit Function`
		// always leaves (issue #273). 'set' proves nothing, since a Set from
		// a call may store Nothing.
		knownCondition: (condition) => conditionValue(condition, {
			value: () => undefined,
			isNothing: (lower) => (locals.has(lower) && !locals.get(lower)!.letOnly && state.get(lower) === 'unset' ? true : undefined),
		}),
		setSilent: (quiet) => {
			silent = quiet;
		},
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
 * The Then and Else arms of a single-line If, as offsets. A nested one-line If
 * sits inside the outer Then arm, and an Else belongs to the innermost If still
 * open: in `If A Then If B Then X Else Y`, Y is B's (issue #575).
 */
function singleLineIfArms(toks: readonly VbaToken[], thenIndex: number, span: Span): { thenArm: Span; elseArm?: Span } {
	const start = span.start + toks[thenIndex].end;
	let open = 0;
	for (let i = thenIndex + 1; i < toks.length; i++) {
		const word = tokenText(toks[i]);
		if (word === 'if') {
			open++;
		} else if (word === 'else') {
			if (open === 0) {
				return { thenArm: { start, end: span.start + toks[i].start }, elseArm: { start: span.start + toks[i].end, end: span.end } };
			}
			open--;
		}
	}
	return { thenArm: { start, end: span.end } };
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
	facts: ModuleObjectFacts,
	elements: ObjectArrayElements = { arrays: new Map(), keys: new Map() },
): void {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	const head = tokenText(toks[0]);
	// `GoSub Label` runs the subroutine, which may Set any of the locals,
	// before the statement after it (issue #108).
	if (head === 'gosub' || (head === 'on' && toks.some((tok) => tokenText(tok) === 'gosub'))) {
		for (const lower of [...setAnywhere, ...elements.keys.keys()]) {
			if (state.get(lower) === 'unset') {
				state.set(lower, 'unknown');
			}
		}
		return;
	}
	// `a(0).Count` on an element never set (issue #489, measured in Excel 16.0).
	if (elements.keys.size > 0) {
		for (const hit of unsetElementAccesses(toks, elements, state)) {
			push(
				'objectVariableNotSet',
				`Element ${hit.text} of '${elements.arrays.get(hit.text.slice(0, hit.text.indexOf('(')).toLowerCase())?.name ?? hit.text}' is Nothing before member access. This will raise Run-time error '91': Object variable or With block variable not set.`,
				{ start: stmt.span.start + hit.start, end: stmt.span.start + hit.end },
			);
		}
		updateElements(toks, elements, state);
	}
	// The arms of a single-line If and what its condition proves about them.
	const branches = statementAndBranchSpans(stmt);
	let guards = { thenArm: new Set<string>(), elseArm: new Set<string>() };
	let arms: { thenArm: Span; elseArm?: Span } | undefined;
	if (head === 'if' && branches.length > 1) {
		const thenIndex = toks.findIndex((tok, index) => index > 0 && tokenText(tok) === 'then');
		if (thenIndex > 0) {
			guards = nothingGuardNames(toks.slice(1, thenIndex));
			arms = singleLineIfArms(toks, thenIndex, stmt.span);
		}
	}
	const guardedAt = (name: string, offset: number): boolean => {
		const within = (span: Span | undefined): boolean =>
			span !== undefined && offset >= span.start && offset < span.end;
		return (guards.thenArm.has(name) && within(arms?.thenArm))
			|| (guards.elseArm.has(name) && within(arms?.elseArm));
	};
	// A bare `obj = value` is a Let through the object's default member
	// (issue #107), which needs an object to reach: on a variable still
	// Nothing it raises 91, the same as a member access would.
	for (const span of branches) {
		const let_ = bareAssignmentTarget(source, span);
		const lower = let_?.name.toLowerCase();
		if (!let_ || !lower || !locals.has(lower) || locals.get(lower)!.variant) {
			continue;
		}
		const letState = guardedAt(lower, let_.span.start) ? 'unknown' : state.get(lower) ?? 'unknown';
		lets.set(let_.span.start, letState);
		// A type with no default member for the Let, or one that needs an
		// argument, is set-required's to report, with the 91 when it is still
		// Nothing (issue #193): the fix there is the Set.
		const verdict = objectLetAssignmentVerdict(locals.get(lower)!.asType, memberCtx);
		if (letState === 'unset' && verdict !== 'noDefault' && verdict !== 'argument' && !objectHoldingDefault(locals.get(lower)!.asType, memberCtx)) {
			const what = locals.get(lower)!.letOnly ? `The result '${let_.name}'` : `Object variable '${let_.name}'`;
			push(
				'objectVariableNotSet',
				`${what} is Nothing before the default-member assignment. This will raise Run-time error '91': Object variable or With block variable not set.`,
				let_.span,
			);
		}
	}
	// `If c Then` reads c's value for the condition: on c still Nothing that
	// raises 91 whatever its type's default member (issue #268, measured in
	// Excel 16.0 on a Collection). A type with no default member is
	// object-default-value's, 438 or 91 (issue #415).
	// So do a loop's condition, Select Case, IIf, Not and And (issue #424).
	for (const { index, form } of conditionOperands(toks)) {
		const lower = toks[index].rawText.toLowerCase();
		const local = locals.get(lower);
		if (local && !local.letOnly && !local.variant && state.get(lower) === 'unset' && !guardedAt(lower, stmt.span.start + toks[index].start)
			&& objectLetAssignmentVerdict(local.asType, memberCtx) !== 'noDefault') {
			const reads = form === 'condition' ? 'the condition reads' : form === 'select' ? 'Select Case reads' : form === 'iif' ? 'IIf reads' : `'${form === 'not' ? 'Not' : 'the Boolean operator'}' reads`;
			push(
				'objectVariableNotSet',
				`Object variable '${toks[index].rawText}' is Nothing when ${reads} its value. This will raise Run-time error '91': Object variable or With block variable not set.`,
				{ start: stmt.span.start + toks[index].start, end: stmt.span.start + toks[index].end },
			);
		}
	}
	// `x = c` reads c's default member, which needs an object: on c still
	// Nothing it raises 91 (issue #256, measured in Excel 16.0). A type with
	// no default member, or one that needs an argument, is object-default-value's.
	for (const span of branches) {
		const target = bareAssignmentTarget(source, span);
		const value = target?.valueTokens.filter((tok) => tok.kind !== 'comment') ?? [];
		const lower = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
		const local = lower ? locals.get(lower) : undefined;
		if (!target || !local || local.letOnly || local.variant || locals.has(target.name.toLowerCase()) || state.get(lower!) !== 'unset'
			|| guardedAt(lower!, span.start + value[0].start) || objectLetAssignmentVerdict(local.asType, memberCtx) !== 'lets') {
			continue;
		}
		push(
			'objectVariableNotSet',
			`Object variable '${value[0].rawText}' is Nothing when its default member is read. This will raise Run-time error '91': Object variable or With block variable not set.`,
			{ start: span.start + value[0].start, end: span.start + value[0].end },
		);
	}
	// `x + 1`, `x & "a"` read x's default member as an operand: on x still
	// Nothing that raises 91 (issue #462, measured in Word and PowerPoint
	// 16.0 on a Range, the Selection and a TextRange).
	for (const span of branches) {
		const operandToks = statementTokens(source, span).filter((tok) => tok.kind !== 'comment');
		// A Set's `=` is no operator: `Set x = y` reads neither value. Its
		// value indexed, `Set p = o(1)`, calls o's default member, which
		// needs o (issue #296, measured in Excel 16.0: 91).
		if (tokenText(operandToks[0]) === 'set' || setAssignmentTarget(source, span)) {
			const eq = operandToks.findIndex((tok) => tok.rawText === '=');
			const value = operandToks[eq + 1];
			const lower = tokenName(value)?.toLowerCase();
			const local = lower ? locals.get(lower) : undefined;
			if (eq > 0 && local && !local.letOnly && !local.variant && operandToks[eq + 2]?.rawText === '(' && state.get(lower!) === 'unset'
				&& !guardedAt(lower!, span.start + value.start) && matchParenFrom(operandToks, eq + 2) === operandToks.length - 1
				&& objectLetAssignmentVerdict(local.asType, memberCtx) !== 'noDefault') {
				push(
					'objectVariableNotSet',
					`Object variable '${value.rawText}' is Nothing when its default member is indexed. This will raise Run-time error '91': Object variable or With block variable not set.`,
					{ start: span.start + value.start, end: span.start + value.end },
				);
			}
			continue;
		}
		const target = bareAssignmentTarget(source, span);
		const eq = target ? operandToks.findIndex((tok) => tok.rawText === '=') : -1;
		const then = span === stmt.span && branches.length > 1 ? operandToks.findIndex((tok) => tokenText(tok) === 'then') : -1;
		const limit = then > 0 ? then : operandToks.length;
		const isOperator = (index: number): boolean => index !== eq && operandToks[index] !== undefined
			&& ((operandToks[index].kind === 'operator' && OPERAND_OPERATORS.has(operandToks[index].rawText)) || tokenText(operandToks[index]) === 'mod');
		for (let i = 0; i < limit; i++) {
			const lower = tokenName(operandToks[i])?.toLowerCase();
			const local = lower ? locals.get(lower) : undefined;
			if (!local || local.letOnly || local.variant || i === eq - 1 || operandToks[i - 1]?.rawText === '.' || operandToks[i + 1]?.rawText === '.'
				|| state.get(lower!) !== 'unset' || guardedAt(lower!, span.start + operandToks[i].start) || objectHoldingDefault(local.asType, memberCtx)) {
				continue;
			}
			const verdict = objectLetAssignmentVerdict(local.asType, memberCtx);
			// `CStr(x)`, `Len(x)`: a whole argument of a built-in that reads one
			// value (issue #415, measured in Excel 16.0). A Collection or Names
			// there does not compile, which is collection-operand's; a type with
			// no default member is object-default-value's, 438 or 91.
			const argument = ['(', ','].includes(operandToks[i - 1]?.rawText ?? '') && [')', ','].includes(operandToks[i + 1]?.rawText ?? '')
				&& ONE_VALUE_BUILTINS.has(tokenText(operandToks[builtinNameBefore(operandToks, i)])) && verdict === 'lets';
			// `x(1)` passes the index to the default member: an Excel Range's
			// takes one, an Object is late bound, and a Collection's needs one.
			// A default that takes none, as Application's Name, does not
			// compile; a type with none is object-default-value's.
			const type = normalizeType(local.asType);
			const close = operandToks[i + 1]?.rawText === '(' ? matchParenFrom(operandToks, i + 1) : -1;
			const indexed = close > i + 2 && operandToks[close + 1]?.rawText !== '.' && (verdict === 'argument' || type === 'object' || (type === 'range' && memberCtx.model?.hostName !== 'Word' && memberCtx.model?.hostName !== 'PowerPoint'));
			const operand = operandToks[i + 1]?.rawText !== '(' && (isOperator(i - 1) || isOperator(i + 1)) && verdict === 'lets';
			if (!argument && !indexed && !operand) {
				continue;
			}
			push(
				'objectVariableNotSet',
				`Object variable '${operandToks[i].rawText}' is Nothing when its ${indexed ? 'default member is indexed' : argument ? 'value is read' : 'default member is read as an operand'}. This will raise Run-time error '91': Object variable or With block variable not set.`,
				{ start: span.start + operandToks[i].start, end: span.start + operandToks[i].end },
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
	// Nothing passed to a procedure of the module that reads a member of the
	// parameter first raises 91 there (issue #343, measured in Excel 16.0).
	for (const span of branches) {
		for (const pass of nothingPassedToMemberRead(statementTokens(source, span), locals, state, facts.memberFirst)) {
			if (guardedAt(pass.name.toLowerCase(), span.start + pass.start)) {
				continue;
			}
			push(
				'objectVariableNotSet',
				`Object variable '${pass.name}' is Nothing, and ${pass.callee} reads '${pass.read}' from it first. This will raise Run-time error '91': Object variable or With block variable not set.`,
				{ start: span.start + pass.start, end: span.start + pass.end },
			);
		}
	}
	const target = setAssignmentTarget(source, stmt.span);
	if (target) {
		const lower = target.name.toLowerCase();
		if (locals.has(lower)) {
			state.set(lower, setValueState(target, locals.get(lower)!, locals, state, facts.nothingFunctions));
			return;
		}
	}
	// A Let gives a Variant a value that is no object.
	const let_ = bareAssignmentTarget(source, stmt.span)?.name.toLowerCase();
	if (let_ && locals.get(let_)?.variant) {
		state.set(let_, 'unknown');
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
		const letTarget = bareAssignmentTarget(source, branch)?.name.toLowerCase();
		if (letTarget && locals.get(letTarget)?.variant && state.get(letTarget) === 'unset') {
			state.set(letTarget, 'unknown');
		}
	}
}

/**
 * What a Set leaves in its target: Nothing from `Nothing`, from a local
 * still Nothing, or from a Function of the module that returns Nothing
 * (issue #343); a local's own state from a local; otherwise an object.
 */
function setValueState(
	target: { valueTokens: readonly VbaToken[] },
	into: LocalObjectVariable,
	locals: ReadonlyMap<string, LocalObjectVariable>,
	state: ReadonlyMap<string, ObjectVariableState>,
	nothingFunctions: ReadonlyMap<string, ProcedureNode>,
): ObjectVariableState {
	if (setAssignmentValueIsNothing(target)) {
		return 'unset';
	}
	const toks = target.valueTokens.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
	const lower = tokenName(toks[0])?.toLowerCase() ?? '';
	const local = locals.get(lower);
	if (toks.length === 1 && local && !local.letOnly) {
		// A late-bound copy of a typed variable is runtime-member-not-found's,
		// which names the 91 with the 438 its members raise.
		const lateBound = (type: string): boolean => ['object', 'variant'].includes(normalizeType(type) ?? 'variant');
		const copied = state.get(lower) ?? 'unknown';
		return copied === 'unset' && lateBound(into.asType) && !lateBound(local.asType) ? 'unknown' : copied;
	}
	const fn = nothingFunctions.get(lower);
	const called = toks.length === 1 ? fn?.params.length === 0 : toks[1]?.rawText === '(' && matchParenFrom(toks, 1) === toks.length - 1;
	return fn && called ? 'unset' : 'set';
}

/**
 * The locals still Nothing that a statement passes whole to a procedure of
 * the module whose parameter there is read with a member first:
 * `TakeC(o)`, `TakeC o`, `Call TakeC(o)`. Offsets are the statement's.
 */
function nothingPassedToMemberRead(
	toks: readonly VbaToken[],
	locals: ReadonlyMap<string, LocalObjectVariable>,
	state: ReadonlyMap<string, ObjectVariableState>,
	memberFirst: ReadonlyMap<string, ReadonlyMap<number, string>>,
): Array<{ name: string; callee: string; read: string; start: number; end: number }> {
	const out: Array<{ name: string; callee: string; read: string; start: number; end: number }> = [];
	const code = toks.filter((tok) => tok.kind !== 'comment');
	for (let i = 0; i < code.length; i++) {
		const reads = memberFirst.get(tokenName(code[i])?.toLowerCase() ?? '');
		if (!reads || code[i - 1]?.rawText === '.' || code[i - 1]?.rawText === '!') {
			continue;
		}
		// Parenthesized anywhere, or bare as the statement's first word.
		const paren = code[i + 1]?.rawText === '(' ? matchParenFrom(code, i + 1) : -1;
		const bare = paren < 0 && (i === 0 || (i === 1 && tokenText(code[0]) === 'call'));
		if (paren < 0 && !bare) {
			continue;
		}
		const args = paren > 0
			? splitTopLevelTokenGroups(code, i + 2, ',', paren)
			: splitTopLevelTokenGroups(code, i + 1, ',', code.length);
		for (const [k, read] of reads) {
			const arg = args[k];
			const lower = arg?.length === 1 ? tokenName(arg[0])?.toLowerCase() : undefined;
			const local = lower ? locals.get(lower) : undefined;
			if (!local || local.letOnly || state.get(lower!) !== 'unset') {
				continue;
			}
			out.push({ name: arg![0].rawText, callee: code[i].rawText, read, start: arg![0].start, end: arg![0].end });
		}
	}
	return out;
}

/** The object arrays of a procedure and the elements its code names by a literal index. */
interface ObjectArrayElements {
	/** By lowercased name: the array as declared, and whether its bounds are fixed. */
	arrays: Map<string, { name: string; fixed: boolean }>;
	/** By state key, `a(0)`: the array and the index. */
	keys: Map<string, { array: string; index: number }>;
}

/**
 * The local arrays of an object type, not `As New`, and every `a(n)` the
 * procedure writes with a whole-number literal n (issue #489, measured in
 * Excel 16.0). Only those elements are followed.
 */
function objectArrayElements(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
): ObjectArrayElements {
	const arrays = new Map<string, { name: string; fixed: boolean }>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		const elementType = child.asType?.replace(/\(\s*\)\s*$/, '');
		if (child.kind === 'localVariable' && child.isArray && child.visibility !== 'Static' && !child.isAutoInstantiated
			&& elementType && isKnownObjectAssignmentType(elementType, memberCtx)) {
			arrays.set(child.name.toLowerCase(), { name: child.name, fixed: child.arrayBounds !== undefined });
		}
	}
	const keys = new Map<string, { array: string; index: number }>();
	if (arrays.size === 0) {
		return { arrays, keys };
	}
	forEachStatement(proc.body, (stmt) => {
		const toks = statementTokens(source, stmt.span);
		for (let i = 0; i + 3 < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			if (lower && arrays.has(lower) && toks[i - 1]?.rawText !== '.' && toks[i + 1].rawText === '(' && toks[i + 2].kind === 'integerLiteral' && toks[i + 3].rawText === ')') {
				const index = Number(toks[i + 2].rawText.replace(/[%&^]$/, ''));
				if (Number.isInteger(index)) {
					keys.set(`${lower}(${index})`, { array: lower, index });
				}
			}
		}
	}, activity);
	return { arrays, keys };
}

/**
 * `For Each x In a` over a fixed array of objects the procedure never fills:
 * x is Nothing on the first pass, so the body's first use of it through a
 * member raises 91 (issue #489, measured in Excel 16.0).
 */
function checkForEachOverEmptyObjectArray(
	source: string,
	member: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const { arrays } = objectArrayElements(source, symbols, member, memberCtx, activity);
	const empty = new Set([...arrays].filter(([, array]) => array.fixed).map(([lower]) => lower));
	if (empty.size === 0) {
		return;
	}
	forEachStatement(member.body, (stmt) => {
		for (const span of statementAndBranchSpans(stmt)) {
			const toks = statementTokensAfterLeadingLabel(source, span);
			for (const lower of [...empty]) {
				// Any write of an element, or a whole mention, may fill it.
				const named = toks.some((tok, i) => tokenName(tok)?.toLowerCase() === lower && toks[i - 1]?.rawText !== '.'
					&& (toks[i + 1]?.rawText !== '(' || tokenText(toks[i - 1]) === 'set' || ['redim', 'erase'].includes(tokenText(toks[0]))));
				if (named) {
					empty.delete(lower);
				}
			}
		}
	}, activity);
	const visit = (body: readonly BodyNode[]): void => {
		for (const node of body) {
			if (isInactiveNode(activity, node) || !('body' in node) || !Array.isArray(node.body)) {
				continue;
			}
			const over = node.kind === 'ForBlock' && node.each ? node.sourceExpression?.trim().toLowerCase() : undefined;
			const control = node.kind === 'ForBlock' ? node.controlVariable?.toLowerCase() : undefined;
			if (over && control && empty.has(over)) {
				for (const stmt of node.body as BodyNode[]) {
					if (isInactiveNode(activity, stmt)) {
						continue;
					}
					if (!isLeafStatement(stmt) || statementAndBranchSpans(stmt).length > 1) {
						break;
					}
					const toks = statementTokensAfterLeadingLabel(source, stmt.span);
					const at = toks.findIndex((tok, i) => tokenName(tok)?.toLowerCase() === control && toks[i - 1]?.rawText !== '.');
					if (at < 0) {
						continue;
					}
					if (tokenText(toks[0]) !== 'set' && toks[at + 1]?.rawText === '.' && tokenName(toks[at + 2])) {
						push(
							'objectVariableNotSet',
							`'${toks[at].rawText}' takes each element of '${arrays.get(over)!.name}', which the code never sets, so it is Nothing on the first pass. This will raise Run-time error '91': Object variable or With block variable not set.`,
							{ start: stmt.span.start + toks[at].start, end: stmt.span.start + toks[at].end },
						);
					}
					break;
				}
			}
			visit(node.body as BodyNode[]);
		}
	};
	visit(member.body);
}

/** The element keys of an array, `a(0)` and the rest. */
function elementKeysOf(elements: ObjectArrayElements, array: string): string[] {
	return [...elements.keys].filter(([, element]) => element.array === array).map(([key]) => key);
}

/**
 * What a statement does to the followed elements, after its reads: `Set
 * a(0) = ...` sets one, `Set a(i) = ...` may set any, a plain ReDim and
 * Erase leave every one Nothing, ReDim Preserve keeps them, and any other
 * whole mention of the array may change them.
 */
function updateElements(toks: readonly VbaToken[], elements: ObjectArrayElements, state: Map<string, ObjectVariableState>): void {
	const head = tokenText(toks[0]);
	for (const array of elements.arrays.keys()) {
		const keys = elementKeysOf(elements, array);
		if (keys.length === 0 || !toks.some((tok) => tokenName(tok)?.toLowerCase() === array)) {
			continue;
		}
		if (head === 'set' && tokenName(toks[1])?.toLowerCase() === array && toks[2]?.rawText === '(') {
			const close = matchParenFrom(toks, 2);
			const literal = close === 4 && toks[3].kind === 'integerLiteral' ? `${array}(${Number(toks[3].rawText.replace(/[%&^]$/, ''))})` : undefined;
			const value = toks.slice(close + 2).filter((tok) => tok.kind !== 'comment');
			const nothing = value.length === 1 && tokenText(value[0]) === 'nothing';
			if (literal && state.has(literal)) {
				state.set(literal, nothing ? 'unset' : 'set');
			} else if (!literal) {
				for (const key of keys) {
					state.set(key, 'unknown');
				}
			}
			continue;
		}
		// A Set in a one-line If's branch may run or not.
		if (head !== 'set' && toks.some((tok, i) => tokenText(tok) === 'set' && tokenName(toks[i + 1])?.toLowerCase() === array)) {
			for (const key of keys) {
				if (state.get(key) === 'unset') {
					state.set(key, 'unknown');
				}
			}
			continue;
		}
		if (head === 'redim' || head === 'erase') {
			if (!(head === 'redim' && tokenText(toks[1]) === 'preserve')) {
				for (const key of keys) {
					state.set(key, 'unset');
				}
			}
			continue;
		}
		// `Fill a`, `b = a`: the whole array, which the callee or a copy may change.
		const whole = toks.some((tok, i) => tokenName(tok)?.toLowerCase() === array && toks[i - 1]?.rawText !== '.' && toks[i + 1]?.rawText !== '(');
		if (whole) {
			for (const key of keys) {
				state.set(key, 'unknown');
			}
		}
	}
}

/** The followed elements a statement may change, for a block that runs it or not. */
function elementTouches(toks: readonly VbaToken[], elements: ObjectArrayElements): string[] {
	const out: string[] = [];
	for (const array of elements.arrays.keys()) {
		if (toks.some((tok, i) => tokenName(tok)?.toLowerCase() === array && toks[i - 1]?.rawText !== '.' && (toks[i + 1]?.rawText !== '(' || tokenText(toks[0]) === 'set'))
			|| (['redim', 'erase'].includes(tokenText(toks[0])) && toks.some((tok) => tokenName(tok)?.toLowerCase() === array))) {
			out.push(...elementKeysOf(elements, array));
		}
	}
	return out;
}

/** `a(0).Count` with a(0) still Nothing. Offsets are the statement's. */
function unsetElementAccesses(toks: readonly VbaToken[], elements: ObjectArrayElements, state: ReadonlyMap<string, ObjectVariableState>): Array<{ text: string; start: number; end: number }> {
	const out: Array<{ text: string; start: number; end: number }> = [];
	for (let i = 0; i + 5 < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		if (!lower || !elements.arrays.has(lower) || toks[i - 1]?.rawText === '.' || toks[i + 1].rawText !== '(' || toks[i + 2].kind !== 'integerLiteral' || toks[i + 3].rawText !== ')'
			|| toks[i + 4].rawText !== '.' || !tokenName(toks[i + 5])) {
			continue;
		}
		const key = `${lower}(${Number(toks[i + 2].rawText.replace(/[%&^]$/, ''))})`;
		if (state.get(key) === 'unset') {
			out.push({ text: toks.slice(i, i + 4).map((tok) => tok.rawText).join(''), start: toks[i].start, end: toks[i + 3].end });
		}
	}
	return out;
}

/** Intrinsics that read an object argument and never Set it. */
const OBJECT_READ_ONLY_INTRINSICS: ReadonlySet<string> = new Set([
	'typename', 'vartype', 'isobject', 'isnull', 'isempty', 'ismissing', 'objptr',
]);

function localObjectVariablesFor(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	memberCtx: MemberCompletionContext,
): Map<string, LocalObjectVariable> {
	const out = new Map<string, LocalObjectVariable>();
	const procSym = procedureSymbolFor(symbols, proc);
	for (const child of procSym?.children ?? []) {
		// `DefObj O` then `Dim o` is an Object (issue #285).
		const asType = child.asType ?? (child.kind === 'localVariable' ? defTypeOf(symbols, child.name) : undefined);
		if (
			child.kind !== 'localVariable' ||
			child.visibility === 'Static' ||
			child.isArray === true ||
			// `Dim x As New Invoice` is instantiated on ANY access, including
			// the first one and including after `Set x = Nothing`, so it can
			// never be Nothing when a member is touched. Tracking it produced
			// error 91 warnings on code that runs.
			child.isAutoInstantiated === true ||
			!isKnownObjectAssignmentType(asType, memberCtx) ||
			!asType
		) {
			continue;
		}
		out.set(child.name.toLowerCase(), { name: child.name, asType });
	}
	// A Variant is followed only where the procedure sets it to Nothing.
	const text = source.slice(proc.span.start, proc.span.end);
	for (const child of procSym?.children ?? []) {
		const type = normalizeType(child.asType);
		if (child.kind === 'localVariable' && child.visibility !== 'Static' && !child.isArray && (type === undefined || type === 'variant')
			&& new RegExp(`\\bset\\s+${child.name}\\s*=\\s*nothing\\b`, 'i').test(text)) {
			out.set(child.name.toLowerCase(), { name: child.name, asType: 'Variant', variant: true });
		}
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
