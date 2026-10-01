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
//
// Issue #224 (measured in Excel 16.0): the class also reaches the variable
// from one declared as it, `Set o = c`, where c raises 91 instead while it is
// Nothing. A Private member is not on the list (438). A property with a Get
// and no Let raises 451 when assigned, and one with a Let and no Get 450 when
// read.

import type { HostObjectModel } from '../../host/excelObjectModel';
import { getHostMembers, getHostType } from '../../host/hostModel';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import { projectTypeAt, resolveReceiverTypeAt } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { statementLabelDeclaration } from '../../flow/procedureLabels';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ForBlockNode, ModuleNode, ProcedureNode } from '../../parser/nodes';
import { heldObjectsAt } from '../heldObjects';
import { isLeafStatement } from '../../parser/nodes';
import { walkEnteringBlocks } from '../dataflow';
import { bodyMayLeaveLoop, namesIn } from './shared';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { normalizeType, objectAssignmentIncompatibilityReason, stringLiteralValue, typeEnvironmentFor } from '../typeInference';
import {
	activeModuleMembers,
	forEachStatement,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

const COLLECTION_MEMBERS: ReadonlySet<string> = new Set(['add', 'count', 'item', 'remove']);

/** Names a late-bound local is known to hold: the class display name and its members. */
interface KnownClass {
	display: string;
	members: ReadonlySet<string>;
	/** Properties with a Get and no Let or Set: assigning one raises 451. */
	readOnly?: ReadonlySet<string>;
	/** Properties with a Let and no Get: reading one raises 450. */
	writeOnly?: ReadonlySet<string>;
	/** Set from a variable that may still be Nothing: 91 before 438. */
	mayBeNothing?: boolean;
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
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkFormControlNames(source, span.start, statementTokens(source, span), memberCtx, push);
			}
		}, activity);
		checkCollectionItems(source, member, symbols, env, memberCtx, activity, push);
		const autoInstanced = new Set<string>();
		for (const child of procedureSymbolFor(symbols, member)?.children ?? []) {
			if (child.isAutoInstantiated) {
				autoInstanced.add(child.name.toLowerCase());
			}
		}
		// Asked only for the target of a Set: walking the whole environment
		// for every procedure was 5% of a large module's pass (issue #139).
		const isLateBound = (lower: string): boolean => {
			if (!env.has(lower)) {
				return false;
			}
			const normalized = normalizeType(env.get(lower));
			return normalized === 'object' || normalized === 'variant' || normalized === undefined;
		};
		const held = new Map<string, KnownClass>();
		// Blocks are entered with the state they start with (issue #237).
		const visit = (node: BodyNode): void => {
			if (!isLeafStatement(node)) {
				return; // a Dim inside the body declares, and runs nothing
			}
			const toks = statementTokensAfterLeadingLabel(source, node.span);
			if (statementLabelDeclaration(source, node.span) || tokenText(toks[0]) === 'gosub') {
				held.clear();
			}
			if (node.kind === 'Statement' && node.singleLineIfBranches) {
				forgetMentioned(toks, held);
				return;
			}
			checkStatement(source, node.span.start, toks, held, applicationSurface, memberCtx, push);
			const set = setAssignmentTarget(source, node.span);
			if (set && isLateBound(set.name.toLowerCase())) {
				const lower = set.name.toLowerCase();
				const value = toks.slice(toks.findIndex((tok) => tok.rawText === '=') + 1);
				const source1 = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
				const fromVariable = source1 !== undefined && !isLateBound(source1) ? knownClassNamed(env.get(source1), memberCtx) : undefined;
				const known = value.length === 2 && tokenText(value[0]) === 'new'
					? knownClassNamed(tokenName(value[1]), memberCtx)
					: fromVariable && { ...fromVariable, mayBeNothing: !autoInstanced.has(source1!) };
				if (known) {
					held.set(lower, known);
				} else {
					held.delete(lower);
				}
				return;
			}
			forgetOtherUses(toks, held);
		};
		walkEnteringBlocks(source, member.body, (node) => activity?.isInactive(node.span) === true, visit, {
			snapshot: () => new Map(held),
			restore: (saved) => {
				held.clear();
				for (const [lower, known] of saved) {
					held.set(lower, known);
				}
			},
			forget: (names) => {
				for (const lower of names) {
					held.delete(lower);
				}
			},
			touches: (stmt) => namesIn(source, stmt.span),
		});
	}
}

/**
 * The items a Collection local holds, by class (issue #246, measured in Excel
 * 16.0): `c.Add New Flat1` then `c(1).Radius()` asks a Flat1 for a member it
 * lacks, 438; and `For Each x In c` with x a Round1 Sets each item into x,
 * raising 13 at an item of another class.
 */
function checkCollectionItems(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	env: ReadonlyMap<string, string>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const heldAt = heldObjectsAt(source, proc, symbols, activity);
	forEachStatement(proc.body, (stmt) => {
		const items = heldAt(stmt).items;
		if (items.size === 0 || (stmt.kind === 'Statement' && stmt.singleLineIfBranches)) {
			return;
		}
		const toks = statementTokens(source, stmt.span);
		for (let i = 0; i < toks.length; i++) {
			const lower = tokenName(toks[i])?.toLowerCase();
			const held = lower ? items.get(lower) : undefined;
			if (!held || held.length === 0 || toks[i - 1]?.rawText === '.') {
				continue;
			}
			// `c(1).Member` or `c.Item(1).Member`.
			const open = toks[i + 1]?.rawText === '(' ? i + 1 : toks[i + 1]?.rawText === '.' && tokenText(toks[i + 2]) === 'item' && toks[i + 3]?.rawText === '(' ? i + 3 : -1;
			const close = open >= 0 ? matchParen(toks, open) : -1;
			if (close < 0 || toks[close + 1]?.rawText !== '.' || !tokenName(toks[close + 2])) {
				continue;
			}
			const index = close === open + 2 && toks[open + 1].kind === 'integerLiteral' ? Number(toks[open + 1].rawText) : undefined;
			const className = index !== undefined && index >= 1 && index <= held.length ? held[index - 1]
				: held.every((name) => name.toLowerCase() === held[0].toLowerCase()) ? held[0] : undefined;
			const known = knownClassNamed(className, memberCtx);
			const memberName = tokenName(toks[close + 2])!;
			if (!known || known.members.has(memberName.toLowerCase())) {
				continue;
			}
			const shown = toks.slice(i, close + 1).map((tok) => tok.rawText).join('');
			push('runtimeMemberNotFound', `'${shown}' holds a ${known.display} here, which has no member '${memberName}'. This will raise Run-time error '438': Object doesn't support this property or method.`, { start: stmt.span.start + toks[i].start, end: stmt.span.start + toks[close].end });
		}
	}, activity);
	forEachLoopIn(proc.body, activity, (loop) => {
		const source1 = loop.sourceExpression?.trim().toLowerCase();
		const control = loop.controlVariable?.toLowerCase();
		const held = source1 ? heldAt(loop).items.get(source1) : undefined;
		const expected = control ? env.get(control) : undefined;
		if (!held || !expected || !loop.sourceExpressionSpan) {
			return;
		}
		// A body that may leave the loop may stop before any later item: only
		// the first is certainly Set (issue #356, measured in Excel 16.0).
		const reached = bodyMayLeaveLoop(source, loop.body) ? held.slice(0, 1) : held;
		const position = reached.findIndex((name) => objectAssignmentIncompatibilityReason(expected, { type: name, label: name, span: loop.sourceExpressionSpan! }, memberCtx));
		if (position >= 0) {
			push('assignmentObjectTypeMismatch', `For Each Sets each item of '${loop.sourceExpression!.trim()}' into '${loop.controlVariable}', a ${expected}, and item ${position + 1} is a ${held[position]}. This will raise Run-time error '13': Type mismatch.`, loop.sourceExpressionSpan);
		}
	});
}

function matchParen(toks: readonly VbaToken[], open: number): number {
	let depth = 0;
	for (let i = open; i < toks.length; i++) {
		depth += toks[i].rawText === '(' ? 1 : toks[i].rawText === ')' ? -1 : 0;
		if (depth === 0) {
			return i;
		}
	}
	return -1;
}

/** Every For Each loop in a body, nested ones included. */
function forEachLoopIn(body: readonly BodyNode[], activity: ConditionalActivityTracker | undefined, visit: (loop: ForBlockNode) => void): void {
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (node.kind === 'ForBlock' && node.each) {
			visit(node);
		}
		if ('body' in node && Array.isArray(node.body)) {
			forEachLoopIn(node.body as BodyNode[], activity, visit);
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
	const properties = projectType.members.filter((m) => m.kind === 'property' && m.signature !== undefined);
	return {
		display: projectType.name,
		members: new Set(projectType.members.map((m) => m.name.toLowerCase())),
		readOnly: new Set(properties.filter((m) => !m.letAccessor && !m.setAccessor).map((m) => m.name.toLowerCase())),
		writeOnly: new Set(projectType.members.filter((m) => m.kind === 'property' && m.letAccessor && m.signature === undefined).map((m) => m.name.toLowerCase())),
	};
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
			const lower = memberName.toLowerCase();
			const nothing = known.mayBeNothing ? `, or '91' while '${receiver}' is Nothing` : '';
			if (!known.members.has(lower)) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, which has no member '${memberName}'. This will raise Run-time error '438': Object doesn't support this property or method${nothing}.`, at);
				continue;
			}
			// `o.RO = 5` as the statement, a Let into a Get-only property.
			const assigned = i === 0 && toks[i + 3]?.rawText === '=';
			if (assigned && known.readOnly?.has(lower)) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' has a Property Get and no Property Let. This will raise Run-time error '451': Property let procedure not defined and property get procedure did not return an object${nothing}.`, at);
			} else if (!assigned && known.writeOnly?.has(lower)) {
				push('runtimeMemberNotFound', `'${receiver}' holds a ${known.display} here, whose '${memberName}' has a Property Let and no Property Get, so it has no value to read. This will raise Run-time error '450': Wrong number of arguments or invalid property assignment${nothing}.`, at);
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

/**
 * `f.Controls("Nope")` on a form whose controls are known, with no control of
 * that name (case-insensitive, those inside a Frame included), raises
 * -2147024809, "Could not find the specified object" (issue #226, measured in
 * Excel 16.0). `Me.Controls(...)` inside the form does the same.
 */
function checkFormControlNames(
	source: string,
	base: number,
	toks: readonly VbaToken[],
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	for (let i = 1; i + 3 < toks.length; i++) {
		if (tokenText(toks[i]) !== 'controls' || toks[i - 1].rawText !== '.' || toks[i + 1].rawText !== '('
			|| toks[i + 2].kind !== 'stringLiteral' || toks[i + 3].rawText !== ')') {
			continue;
		}
		const form = projectTypeAt(source, base + toks[i - 1].end, memberCtx);
		if (form?.kind !== 'userform' || form.exhaustive !== true) {
			continue;
		}
		const name = stringLiteralValue(toks[i + 2].rawText);
		const controls = form.members.filter((member) => /^MSForms\./i.test(member.returns ?? ''));
		if (!controls.some((control) => control.name.toLowerCase() === name.toLowerCase())) {
			push(
				'runtimeMemberNotFound',
				`The form ${form.name} has no control named "${name}". This will raise Run-time error '-2147024809': Could not find the specified object.`,
				{ start: base + toks[i + 2].start, end: base + toks[i + 2].end },
			);
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
