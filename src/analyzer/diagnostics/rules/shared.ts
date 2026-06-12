// Helpers shared by more than one diagnostics rule family.
//
// Extracted verbatim from analyzeModule.ts (audit #0). Everything here is
// used by at least two of the rules/<family>.ts modules: the module-level
// declaration-statement classifier, the conservative read-reference scanner,
// declaration name-token hits, the conditional-compilation branch-order
// scanner, and the exhaustive member-surface resolver.

import { bareCallStatementTarget as callStatementTarget } from '../../call/callContext';
import {
	type MemberCompletion,
	type MemberCompletionContext,
	resolveMemberSurfaceAt,
} from '../../completion/memberAccess';
import {
	collectConditionalDirectives,
	type ConditionalActivityTracker,
} from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	ModuleNode,
	Span,
} from '../../parser/nodes';
import {
	qualifiedProcedureKey,
	type VbaProjectClassMembers,
} from '../../symbols/symbolModel';
import type { CallableTypeSignature } from '../callExtraction';
import {
	absoluteSpan,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	firstExecutableTokenIndex,
	isInactiveNode,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
} from '../walker';

export function resolveExhaustiveMemberSurface(
	source: string,
	dotEndOffset: number,
	memberCtx: MemberCompletionContext,
): { owner: string; members: MemberCompletion[] } | undefined {
	const surface = resolveMemberSurfaceAt(source, dotEndOffset, memberCtx);
	if (!surface?.exhaustive) {
		return undefined;
	}
	return { owner: surface.owner, members: surface.members };
}

const PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS = new Set(['public', 'private', 'friend', 'global']);

export const DEFTYPE_KEYWORDS = new Set([
	'defbool',
	'defbyte',
	'defcur',
	'defdate',
	'defdbl',
	'defdec',
	'defint',
	'deflng',
	'deflnglng',
	'deflngptr',
	'defobj',
	'defsng',
	'defstr',
	'defvar',
]);

export function leadingDeclarationModifierCount(toks: readonly VbaToken[]): number {
	let i = 0;
	while (PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS.has(tokenText(toks[i]))) {
		i++;
	}
	return i;
}

export function moduleDeclarationStatementInProcedure(
	source: string,
	span: Span,
): { label: string; span: Span } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const first = toks[0];
	const head = tokenText(first);
	if (!first) {
		return undefined;
	}
	if (head === 'option') {
		return { label: 'Option statements', span: absoluteSpan(span, first) };
	}
	if (head === 'attribute') {
		return { label: 'Attribute statements', span: absoluteSpan(span, first) };
	}
	if (DEFTYPE_KEYWORDS.has(head)) {
		return {
			label: `${first.canonicalText ?? first.rawText} statements`,
			span: absoluteSpan(span, first),
		};
	}
	const modifierCount = leadingDeclarationModifierCount(toks);
	const declarationHead = tokenText(toks[modifierCount]);
	if (declarationHead === 'type' || declarationHead === 'enum') {
		const tok = toks[modifierCount];
		return {
			label: declarationHead === 'type' ? 'Type blocks' : 'Enum blocks',
			span: tok ? absoluteSpan(span, tok) : absoluteSpan(span, first),
		};
	}
	if (PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS.has(head)) {
		return {
			label: `${first.canonicalText ?? first.rawText} declarations`,
			span: absoluteSpan(span, first),
		};
	}
	return undefined;
}

export interface NameTokenHit {
	name: string;
	span: Span;
	bracketed: boolean;
}

export function declarationNameHit(
	source: string,
	span: Span,
	name: string,
): NameTokenHit | undefined {
	const lower = name.toLowerCase();
	for (const tok of statementTokens(source, span)) {
		const found = tokenName(tok);
		if (found?.toLowerCase() === lower) {
			return nameTokenHit(span, tok, found);
		}
	}
	return undefined;
}

export function nameTokenHit(base: Span, tok: VbaToken, name: string): NameTokenHit {
	return {
		name,
		span: absoluteSpan(base, tok),
		bracketed: tok.kind === 'bracketedIdentifier',
	};
}

export function isBareOrVbaQualifiedIntrinsicCall(toks: readonly VbaToken[], nameIndex: number): boolean {
	if (toks[nameIndex - 1]?.rawText !== '.') {
		return true;
	}
	const qualifier = nameIndex >= 2 ? tokenName(toks[nameIndex - 2]) : undefined;
	return qualifier?.toLowerCase() === 'vba' && toks[nameIndex - 3]?.rawText !== '.';
}

export function forEachUndeclaredReferenceSpan(
	source: string,
	body: BodyNode[],
	visit: (span: Span) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			visit(node.span);
		} else if ('body' in node && Array.isArray(node.body)) {
			visit(blockHeaderLineSpan(source, node.span));
			if (node.kind === 'DoBlock') {
				const footer = blockFooterLineSpan(source, node.span);
				if (footer.start > node.span.start) {
					visit(footer);
				}
			}
			forEachUndeclaredReferenceSpan(source, node.body, visit, activity);
		}
	}
}

export function valueReadReferences(
	source: string,
	span: Span,
	isKnownForSkip: (name: string) => boolean,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): Array<{ name: string; span: Span }> {
	const toks = statementTokens(source, span);
	const out: Array<{ name: string; span: Span }> = [];
	const skip = undeclaredReferenceSkipIndexes(
		source,
		span,
		toks,
		isKnownForSkip,
		moduleSignatures,
		projectMembers,
	);
	for (let i = 0; i < toks.length; i++) {
		if (skip.has(i) || !isPotentialVariableReferenceToken(toks[i])) {
			continue;
		}
		if (toks[i - 1]?.rawText === '.') {
			continue;
		}
		const name = tokenName(toks[i]);
		if (!name) {
			continue;
		}
		out.push({
			name,
			span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
		});
	}
	return out;
}

function undeclaredReferenceSkipIndexes(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
	isKnown: (name: string) => boolean,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): Set<number> {
	const skip = new Set<number>();
	if (toks.length === 0) {
		return skip;
	}
	if (toks[1]?.rawText === ':' || isLineLabelOnlyStatement(source, span, toks)) {
		skip.add(0); // line label declaration
	}
	const firstExecutable = firstExecutableTokenIndex(toks);
	if (moduleDeclarationStatementInProcedure(source, span)) {
		for (let i = firstExecutable; i < toks.length; i++) {
			skip.add(i);
		}
		return skip;
	}
	if (tokenText(toks[firstExecutable]) === 'implements') {
		for (let i = firstExecutable + 1; i < toks.length; i++) {
			skip.add(i);
		}
		return skip;
	}

	const call = callStatementTarget(source, span);
	if (call) {
		const callIdx = toks.findIndex((tok) => span.start + tok.start === call.span.start);
		if (callIdx >= 0) {
			skip.add(callIdx);
			if (!isKnown(call.name)) {
				// Unknown call targets may be external procedures or unresolved call
				// errors; do not also guess about their argument identifiers.
				for (let i = callIdx + 1; i < toks.length; i++) {
					skip.add(i);
				}
			}
		}
	}

	const assignment = simpleAssignmentLhsIdentifierIndex(toks);
	if (assignment >= 0) {
		skip.add(assignment);
	}

	for (let i = 0; i < toks.length; i++) {
		const word = tokenText(toks[i]);
		if (
			isQualifiedProjectCallableQualifier(toks, i, moduleSignatures) ||
			isQualifiedProjectMemberQualifier(toks, i, projectMembers)
		) {
			skip.add(i);
		}
		if (word === 'new' && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (word === 'is' && hasEarlierTypeOf(toks, i) && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (isLabelReferenceKeyword(word) && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (word === 'raiseevent' && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (word === 'addressof' && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (isNamedArgumentLabel(toks, i)) {
			skip.add(i);
		}
	}

	return skip;
}

function isQualifiedProjectCallableQualifier(
	toks: readonly VbaToken[],
	index: number,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	if (!isPotentialVariableReferenceToken(toks[index]) || toks[index + 1]?.rawText !== '.') {
		return false;
	}
	if (!isPotentialVariableReferenceToken(toks[index + 2])) {
		return false;
	}
	const qualifier = tokenName(toks[index]);
	const member = tokenName(toks[index + 2]);
	if (!qualifier || !member) {
		return false;
	}
	return moduleSignatures.has(qualifiedProcedureKey(qualifier, member));
}

function isQualifiedProjectMemberQualifier(
	toks: readonly VbaToken[],
	index: number,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): boolean {
	if (
		!projectMembers ||
		!isPotentialVariableReferenceToken(toks[index]) ||
		toks[index + 1]?.rawText !== '.'
	) {
		return false;
	}
	if (!isPotentialVariableReferenceToken(toks[index + 2])) {
		return false;
	}
	const qualifier = tokenName(toks[index]);
	const member = tokenName(toks[index + 2]);
	if (!qualifier || !member) {
		return false;
	}
	const qualifierLower = qualifier.toLowerCase();
	const memberLower = member.toLowerCase();
	let surface: VbaProjectClassMembers | undefined;
	for (const candidate of projectMembers) {
		if (candidate.name.toLowerCase() !== qualifierLower) {
			continue;
		}
		if (surface) {
			return false;
		}
		surface = candidate;
	}
	if (!surface) {
		return false;
	}
	if (surface.kind === 'standardModule') {
		return true;
	}
	return surface.members.some((candidate) => candidate.name.toLowerCase() === memberLower);
}

function simpleAssignmentLhsIdentifierIndex(toks: readonly VbaToken[]): number {
	let start = firstExecutableTokenIndex(toks);
	if (tokenText(toks[start]) === 'let' || tokenText(toks[start]) === 'set') {
		start++;
	}
	const eq = topLevelOperatorIndex(toks.slice(start), '=');
	if (eq !== 1) {
		return -1;
	}
	const nameTok = toks[start];
	return nameTok && nameTok.kind === 'identifier' ? start : -1;
}

function isLineLabelOnlyStatement(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
): boolean {
	if (toks.length !== 1 || !isPotentialVariableReferenceToken(toks[0])) {
		return false;
	}
	let j = span.start + toks[0].end;
	while (j < source.length && (source[j] === ' ' || source[j] === '\t')) {
		j++;
	}
	return source[j] === ':';
}

function isPotentialVariableReferenceToken(tok: VbaToken | undefined): boolean {
	return tok?.kind === 'identifier' || tok?.kind === 'bracketedIdentifier';
}

function hasEarlierTypeOf(toks: readonly VbaToken[], before: number): boolean {
	for (let i = 0; i < before; i++) {
		if (tokenText(toks[i]) === 'typeof') {
			return true;
		}
	}
	return false;
}

function isLabelReferenceKeyword(word: string): boolean {
	return word === 'goto' || word === 'gosub' || word === 'resume';
}

function isNamedArgumentLabel(toks: readonly VbaToken[], index: number): boolean {
	if (!isPotentialVariableReferenceToken(toks[index])) {
		return false;
	}
	if (toks[index + 1]?.rawText === ':=') {
		return true;
	}
	return toks[index + 1]?.rawText === ':' && toks[index + 2]?.rawText === '=';
}

interface ElseBranchFrame {
	seenElse: boolean;
	start: Span;
	malformed: boolean;
}

type ConditionalBranchOrderIssueKind = 'elseifAfterElse' | 'duplicateElse';

interface ConditionalBranchOrderIssue {
	kind: ConditionalBranchOrderIssueKind;
	directive: { span: Span };
}

interface ConditionalBranchOrderScan {
	issues: ConditionalBranchOrderIssue[];
	malformedBlockSpans: Span[];
}

export function scanConditionalCompilationBranchOrder(mod: ModuleNode): ConditionalBranchOrderScan {
	const stack: ElseBranchFrame[] = [];
	const issues: ConditionalBranchOrderIssue[] = [];
	const malformedBlockSpans: Span[] = [];
	for (const { directive } of collectConditionalDirectives(mod)) {
		switch (directive.directiveKind) {
			case 'If':
				stack.push({ seenElse: false, start: directive.span, malformed: false });
				break;
			case 'ElseIf': {
				const frame = stack[stack.length - 1];
				if (frame?.seenElse) {
					frame.malformed = true;
					issues.push({ kind: 'elseifAfterElse', directive });
				}
				break;
			}
			case 'Else': {
				const frame = stack[stack.length - 1];
				if (frame?.seenElse) {
					frame.malformed = true;
					issues.push({ kind: 'duplicateElse', directive });
				}
				if (frame) {
					frame.seenElse = true;
				}
				break;
			}
			case 'EndIf': {
				const frame = stack.pop();
				if (frame?.malformed) {
					malformedBlockSpans.push({
						start: frame.start.start,
						end: directive.span.end,
					});
				}
				break;
			}
			case 'Const':
			case 'Unknown':
				break;
		}
	}
	return { issues, malformedBlockSpans };
}
