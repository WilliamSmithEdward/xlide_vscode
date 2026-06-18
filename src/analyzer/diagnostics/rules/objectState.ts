// Rule family: object-variable state (audit #0).
//
// Extracted verbatim from analyzeModule.ts: member access on unset object
// variables (straight-line Set tracking) and member access on known scalars.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
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
	isInactiveNode,
	localsPassedAsCallArguments,
	setAssignmentTarget,
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
		const walk = procedureHasUnstructuredFlow(source, member, activity)
			? walkStraightLineBody
			: walkBranchMergedBody;
		walk(member.body, (node) => isInactiveNode(activity, node), {
			onStatement: (stmt) =>
				checkObjectVariableNotSetStatement(source, stmt, locals, state, memberCtx, push),
			onBlock: (node) => {
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
				const lower = setAssignmentTarget(source, stmt.span)?.name.toLowerCase();
				return lower && locals.has(lower) ? [lower] : [];
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

function checkObjectVariableNotSetStatement(
	source: string,
	stmt: LeafStatementNode,
	locals: ReadonlyMap<string, LocalObjectVariable>,
	state: Map<string, ObjectVariableState>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	for (const hit of unsetObjectMemberAccesses(source, stmt.span, locals, state, memberCtx)) {
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
	for (const lower of localsPassedAsCallArguments(source, stmt.span, locals)) {
		if (state.get(lower) === 'unset') {
			state.set(lower, 'unknown');
		}
	}
}

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
