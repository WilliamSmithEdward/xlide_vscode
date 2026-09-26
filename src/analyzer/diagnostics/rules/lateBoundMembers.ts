// Rule: a member the VBE binds at run time, on an object whose class the code
// makes plain and whose member list is complete (issue #121). Measured in
// Excel 16.0 (build 20326, 2026-09-26); each compiles and raises 438, "Object
// doesn't support this property or method", every time it runs.
//
//  - `Application.Zzq`: Application is extensible, so the VBE compiles any
//    name on it (worksheet functions such as Application.Match are ordinary
//    VBA there), and a name that is neither an Application member nor a
//    WorksheetFunction raises when it runs. Excel only: its model lists every
//    member, hidden ones included.
//  - `Dim o As Object: Set o = New Collection: o.Foo`: a late-bound variable
//    holding a class with a known member list. Collection has Add, Count,
//    Item and Remove; a project class module has its public members.

import type { HostObjectModel } from '../../host/excelObjectModel';
import { getHostMembers, getHostType } from '../../host/hostModel';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import { resolveReceiverTypeAt } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { ModuleNode } from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { PushFn } from '../analysisContext';
import { normalizeType, typeEnvironmentFor } from '../typeInference';
import {
	activeModuleMembers,
	setAssignmentTarget,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

const COLLECTION_MEMBERS: ReadonlySet<string> = new Set(['add', 'count', 'item', 'remove']);

/** Names a late-bound local is known to hold: the class display name and its members. */
interface KnownClass {
	display: string;
	members: ReadonlySet<string>;
}

export function checkRuntimeMemberNotFound(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const model = memberCtx.model;
	const applicationSurface = excelApplicationSurface(model);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const lateBound = new Set<string>();
		for (const [lower, type] of env) {
			const normalized = normalizeType(type);
			if (normalized === 'object' || normalized === 'variant' || normalized === undefined) {
				lateBound.add(lower);
			}
		}
		const held = new Map<string, KnownClass>();
		for (const node of member.body) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (node.kind === 'VariableGroup') {
				continue; // a Dim inside the body declares, and runs nothing
			}
			if (!isLeafStatement(node)) {
				held.clear();
				continue;
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				held.clear();
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forgetMentioned(toks, held);
				continue;
			}
			checkStatement(source, node.span.start, toks, held, applicationSurface, memberCtx, push);
			const set = setAssignmentTarget(source, node.span);
			if (set && lateBound.has(set.name.toLowerCase())) {
				const lower = set.name.toLowerCase();
				const value = toks.slice(toks.findIndex((tok) => tok.rawText === '=') + 1);
				const known = value.length === 2 && tokenText(value[0]) === 'new' ? knownClassNamed(tokenName(value[1]), memberCtx) : undefined;
				if (known) {
					held.set(lower, known);
				} else {
					held.delete(lower);
				}
				continue;
			}
			forgetOtherUses(toks, held);
		}
	}
}

function knownClassNamed(name: string | undefined, memberCtx: MemberCompletionContext): KnownClass | undefined {
	if (!name) {
		return undefined;
	}
	if (name.toLowerCase() === 'collection') {
		return { display: 'Collection', members: COLLECTION_MEMBERS };
	}
	const projectType = (memberCtx.projectClassMembers ?? []).find(
		(type) => type.kind === 'class' && type.exhaustive === true && type.name.toLowerCase() === name.toLowerCase(),
	);
	if (!projectType) {
		return undefined;
	}
	return { display: projectType.name, members: new Set(projectType.members.map((m) => m.name.toLowerCase())) };
}

/** Excel's Application members plus the worksheet functions it also answers to. */
function excelApplicationSurface(model: HostObjectModel | undefined): ReadonlySet<string> | undefined {
	if (model && model.hostName !== undefined && model.hostName !== 'Excel') {
		return undefined;
	}
	if (getHostType('Excel.Application', model)?.exhaustive !== true) {
		return undefined;
	}
	const names = new Set<string>();
	for (const member of getHostMembers('Excel.Application', model)) {
		names.add(member.name.toLowerCase());
	}
	for (const member of getHostMembers('Excel.WorksheetFunction', model)) {
		names.add(member.name.toLowerCase());
	}
	return names;
}

function checkStatement(
	source: string,
	base: number,
	toks: readonly VbaToken[],
	held: ReadonlyMap<string, KnownClass>,
	applicationSurface: ReadonlySet<string> | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	for (let i = 0; i + 2 < toks.length; i++) {
		if (toks[i + 1].rawText !== '.' || toks[i - 1]?.rawText === '.') {
			continue;
		}
		const receiver = tokenName(toks[i]);
		const memberName = tokenName(toks[i + 2]);
		if (!receiver || !memberName) {
			continue;
		}
		const at = { start: base + toks[i + 2].start, end: base + toks[i + 2].end };
		const known = held.get(receiver.toLowerCase());
		if (known) {
			if (!known.members.has(memberName.toLowerCase())) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, which has no member '${memberName}'. This will raise Run-time error '438': Object doesn't support this property or method.`, at);
			}
			continue;
		}
		if (applicationSurface && receiver.toLowerCase() === 'application' && !applicationSurface.has(memberName.toLowerCase())) {
			const type = resolveReceiverTypeAt(source, base + toks[i + 1].end, memberCtx);
			if (type === 'Excel.Application') {
				push('runtimeMemberNotFound', `Application has no member '${memberName}', and it is not a worksheet function either. The VBE compiles the name because Application is extensible; this will raise Run-time error '438': Object doesn't support this property or method.`, at);
			}
		}
	}
}

/** A tracked variable named in any position other than `name.Member` is no longer followed. */
function forgetOtherUses(toks: readonly VbaToken[], held: Map<string, KnownClass>): void {
	for (let i = 0; i < toks.length; i++) {
		const lower = tokenName(toks[i])?.toLowerCase();
		if (lower && held.has(lower) && toks[i - 1]?.rawText !== '.' && toks[i + 1]?.rawText !== '.') {
			held.delete(lower);
		}
	}
}

function forgetMentioned(toks: readonly VbaToken[], held: Map<string, KnownClass>): void {
	for (const tok of toks) {
		const lower = tokenName(tok)?.toLowerCase();
		if (lower && held.has(lower)) {
			held.delete(lower);
		}
	}
}
