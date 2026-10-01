// Rule family: the members of a user-defined type (issue #253). Every case
// was measured in Excel 16.0 (build 20326, 2026-10-01).
//
//  - object-variable-not-set: an object member of a Type local that nothing
//    has Set, or that was Set to Nothing, given a member or a call:
//    `t.o.Add 1`, `t.o.Count`, `t.o(1)`, `.o.Add 1` inside `With t`, and
//    `.Add 1` inside `With t.o`. Each raises 91.
//  - scalar-member-access: a member of a member that holds a number or a
//    string, `t.a.Value`, `t.s.Length`, `.a.Value` in `With t`: "Invalid
//    qualifier".
//  - fixed-array-redim: ReDim of a fixed array member, `ReDim t.f(3)`:
//    "Array already dimensioned".
//  - erase-requires-array: Erase of a whole Type value, `Erase t`, or of a
//    member that is no array, `Erase t.a`: "Expected array".
//
// What a member holds comes from typeMemberState.ts.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { LeafStatementNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import { fieldChain, isFixedArrayField, isLeadingDot, moduleTypes, typeKey, typeRootAt, variableSymbolIn, walkWithSubjects, type FieldStep, type ModuleTypes, type WithSubject } from '../typeFields';
import { typeMemberStatesAt, type MemberStatesAt } from '../typeMemberState';
import { isKnownObjectAssignmentType, isKnownScalarType, normalizeType } from '../typeInference';
import { activeModuleMembers, statementTokensAfterLeadingLabel, tokenName, tokenText } from '../walker';
import { moduleOptionBase } from './arrays';

const NOT_SET = `This will raise Run-time error '91': Object variable or With block variable not set.`;

export function checkTypeMembers(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const types = moduleTypes(source, mod, activity);
	if (types.size === 0) {
		return;
	}
	const optionBase = moduleOptionBase(mod, activity);
	const isObjectType = (type: string): boolean => !types.has(type) && type !== 'variant' && isKnownObjectAssignmentType(type, memberCtx);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const statesAt = typeMemberStatesAt(source, symbols, member, types, activity, optionBase, isObjectType);
		walkWithSubjects(source, member.body, activity, symbols, member, types, undefined, (stmt, subject) => {
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			const head = tokenText(toks[0]);
			if (head === 'redim' || head === 'erase') {
				checkResizes(stmt.span, toks, head, symbols, member, types, subject, push);
				return;
			}
			checkWithObject(stmt, toks, subject, statesAt, push);
			for (let i = 0; i < toks.length; i++) {
				const root = typeRootAt(toks, i, symbols, member, types, subject);
				if (!root) {
					continue;
				}
				const steps = fieldChain(toks, root, types);
				const hit = scalarQualifier(stmt.span, toks, i, steps) ?? nothingAccess(stmt, toks, i, steps, statesAt, isObjectType);
				if (hit) {
					push(hit.rule, hit.message, hit.span);
				}
			}
		});
	}
}

/** `t.a.Value`: a member of a field that holds a number or a string. */
function scalarQualifier(
	span: Span,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
): { rule: 'scalarMemberAccess'; message: string; span: Span } | undefined {
	for (const step of steps) {
		const end = step.close ?? step.at;
		const value = !step.field.isArray || step.open !== undefined;
		const type = normalizeType(step.field.type);
		if (value && type && isKnownScalarType(type) && toks[end + 1]?.rawText === '.') {
			return {
				rule: 'scalarMemberAccess',
				message: `Member access on '${step.display}' is invalid because it is declared as ${step.field.typeName}. This is a VBE compile error: Invalid qualifier.`,
				span: { start: span.start + toks[start].start, end: span.start + toks[end + 1].end },
			};
		}
	}
	return undefined;
}

/** `t.o.Add 1` and `t.o(1)` with t.o still Nothing. */
function nothingAccess(
	stmt: LeafStatementNode,
	toks: readonly VbaToken[],
	start: number,
	steps: readonly FieldStep[],
	statesAt: MemberStatesAt,
	isObjectType: (type: string) => boolean,
): { rule: 'objectVariableNotSet'; message: string; span: Span } | undefined {
	for (const step of steps) {
		const next = toks[step.at + 1]?.rawText;
		if (step.field.isArray || !step.path || !isObjectType(step.field.type ?? '') || (next !== '.' && next !== '!' && next !== '(')) {
			continue;
		}
		if (statesAt(stmt, stmt.span.start + toks[start].start).get(step.path) === 'nothing') {
			return {
				rule: 'objectVariableNotSet',
				message: `'${step.display}' is an object member that nothing has Set here, so it is Nothing. ${NOT_SET}`,
				span: { start: stmt.span.start + toks[start].start, end: stmt.span.start + toks[step.at].end },
			};
		}
		return undefined;
	}
	return undefined;
}

/** `.Add 1` inside `With t.o`, while t.o is Nothing: reported once, at the first member it reaches. */
function checkWithObject(
	stmt: LeafStatementNode,
	toks: readonly VbaToken[],
	subject: WithSubject | undefined,
	statesAt: MemberStatesAt,
	push: PushFn,
): void {
	if (!subject?.path || subject.type || !subject.field || subject.field.isArray) {
		return;
	}
	for (let i = 0; i + 1 < toks.length; i++) {
		if (toks[i].rawText !== '.' || !isLeadingDot(toks, i) || !tokenName(toks[i + 1])) {
			continue;
		}
		if (statesAt(stmt, stmt.span.start + toks[i].start).get(subject.path) === 'nothing') {
			push(
				'objectVariableNotSet',
				`The With object '${subject.display}' is an object member that nothing has Set, so it is Nothing here. ${NOT_SET}`,
				{ start: stmt.span.start + toks[i].start, end: stmt.span.start + toks[i + 1].end },
			);
		}
		return;
	}
}

/** `ReDim t.f(3)` of a fixed member, and `Erase t` or `Erase t.a` of what is no array. */
function checkResizes(
	span: Span,
	toks: readonly VbaToken[],
	head: 'redim' | 'erase',
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	types: ModuleTypes,
	subject: WithSubject | undefined,
	push: PushFn,
): void {
	const first = head === 'redim' && tokenText(toks[1]) === 'preserve' ? 2 : 1;
	let from = first;
	for (let i = first; i <= toks.length; i++) {
		if (i < toks.length && (toks[i].rawText !== ',' || depthAt(toks, from, i) > 0)) {
			continue;
		}
		const group = toks.slice(from, i).filter((tok) => tok.kind !== 'comment');
		from = i + 1;
		if (group.length === 0) {
			continue;
		}
		const at = (last: VbaToken): Span => ({ start: span.start + group[0].start, end: span.start + last.end });
		if (head === 'erase' && group.length === 1) {
			const variable = variableSymbolIn(symbols, proc, tokenName(group[0])?.toLowerCase() ?? '');
			const type = variable && !variable.isArray ? typeKey(variable.asType) : undefined;
			if (type && types.has(type)) {
				push('eraseRequiresArray', `Erase target '${group[0].rawText}' is a user-defined type, not an array. This is a VBE compile error: Expected array.`, at(group[0]));
			}
			continue;
		}
		const root = typeRootAt(group, 0, symbols, proc, types, subject);
		const step = root ? fieldChain(group, root, types).at(-1) : undefined;
		if (!step) {
			continue;
		}
		if (head === 'redim' && isFixedArrayField(step.field) && step.open !== undefined) {
			push('fixedArrayRedim', `'${step.display}' is a fixed-size array member, which ReDim cannot resize. This is a VBE compile error: Array already dimensioned.`, at(group[step.at]));
		} else if (head === 'erase' && !step.field.isArray && step.at === group.length - 1 && step.field.type !== undefined
			&& (isKnownScalarType(normalizeType(step.field.type) ?? '') || types.has(step.field.type))) {
			push('eraseRequiresArray', `Erase target '${step.display}' must be an array or Variant, but it is declared As ${step.field.typeName}. This is a VBE compile error: Expected array.`, at(group[step.at]));
		}
	}
}

/** The parenthesis depth at `index`, counting from `from`. */
function depthAt(toks: readonly VbaToken[], from: number, index: number): number {
	let depth = 0;
	for (let i = from; i < index; i++) {
		depth += toks[i].rawText === '(' ? 1 : toks[i].rawText === ')' ? -1 : 0;
	}
	return depth;
}
