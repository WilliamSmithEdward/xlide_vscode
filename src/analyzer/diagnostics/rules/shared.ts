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
import { isReservedIdentifier } from '../../lexer/keywordTable';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	ModuleNode,
	Span,
} from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
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

/**
 * What one "this may appear at most once" rule is looking for. `keyOf` returns
 * undefined for an entry the rule does not govern.
 */
export interface RepeatedKeyRule<T> {
	keyOf(entry: T): string | undefined;
	spanOf(entry: T): Span;
	/**
	 * Whether an earlier entry and a later one may not coexist. Defaults to
	 * "any repeat collides"; Property accessors are the exception that needs it.
	 */
	collides?(earlier: T, later: T): boolean;
	report(repeat: T, earlier: T): void;
}

/**
 * Reports every entry that repeats a key an earlier entry already took, once
 * per repeat, in source order.
 *
 * Entries in different arms of one `#If` chain never reach the compiler
 * together, so they are alternatives rather than repeats however the
 * conditional constants evaluate. Asking `activity` for the arm rather than for
 * the activity is what lets a rule keep looking inside a chain it cannot
 * decide, instead of skipping the whole chain and going blind
 * (github.com/WilliamSmithEdward/xlide_vscode/issues/58).
 *
 * Almost every key is taken once, so a key holds its lone entry directly and
 * grows a list only when a second one claims it.
 */
export function reportRepeatedKeys<T>(
	entries: Iterable<T>,
	activity: ConditionalActivityTracker | undefined,
	rule: RepeatedKeyRule<T>,
): void {
	const taken = new Map<string, T | T[]>();
	for (const entry of entries) {
		const key = rule.keyOf(entry);
		if (key === undefined) {
			continue;
		}
		const holder = taken.get(key);
		if (holder === undefined) {
			taken.set(key, entry);
			continue;
		}
		const earlier = Array.isArray(holder) ? holder : [holder];
		const hit = earlier.find(
			(prior) =>
				(rule.collides?.(prior, entry) ?? true) &&
				!activity?.mutuallyExclusive(rule.spanOf(prior), rule.spanOf(entry)),
		);
		earlier.push(entry);
		taken.set(key, earlier);
		if (hit !== undefined) {
			rule.report(entry, hit);
		}
	}
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
		if (isLeafStatement(node)) {
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
): Array<{ name: string; span: Span; bracketed: boolean }> {
	const toks = statementTokens(source, span);
	const out: Array<{ name: string; span: Span; bracketed: boolean }> = [];
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
		// `!` is VBA's default-member accessor, so `rs!CustomerName` and
		// `Forms!frmMain!txtName` are MEMBER ACCESS exactly as a dot is - the
		// name after it belongs to the receiver and was never a variable. The
		// same character is the Single type suffix (`Dim x!`), but a suffix is
		// only ever followed by an operator or the end of the statement, never
		// by a name, so one test covers both.
		if (toks[i - 1]?.rawText === '.' || toks[i - 1]?.rawText === '!') {
			continue;
		}
		const name = tokenName(toks[i]);
		if (!name) {
			continue;
		}
		out.push({
			name,
			span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
			bracketed: toks[i].kind === 'bracketedIdentifier',
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

	const statementHead = tokenText(toks[firstExecutable]);
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
		// A CONTEXTUAL keyword is not a reserved identifier (MS-VBAL 3.3.5.2), so
		// `Dim Text As String` is legal VBA and a reference to `Text` is a real
		// variable read. The scanner therefore accepts one as a name - but each
		// of these words also has a grammar position where it is SYNTAX, and
		// reading the word there would report the statement's own keyword as an
		// undefined variable. Those positions are skipped here.
		if (isContextualGrammarWord(toks, i, statementHead, firstExecutable)) {
			skip.add(i);
		}
		// Open-statement access-clause (MS-VBAL 5.4.5.1.1): in `Open path For
		// mode [Access access] [lock] As #f`, `Access` is a grammar word, not a
		// variable reference. It lexes as an identifier because it is not a
		// reserved identifier (`Dim Access As Long` is legal), so skip it only
		// in its grammar position: inside an Open statement and immediately
		// followed by the access mode `Read` or `Write`. Every other Open
		// grammar word (For/Binary/Input/Output/Append/Random/Read/Write/Lock/
		// Shared/As/Len) already lexes as a keyword and is never scanned.
		if (
			statementHead === 'open' &&
			word === 'access' &&
			(tokenText(toks[i + 1]) === 'read' || tokenText(toks[i + 1]) === 'write')
		) {
			skip.add(i);
		}
		// `ReDim [Preserve] name(bounds) As TypeName`: the token after `As` is a
		// type reference, not a variable read. Bounds identifiers remain uses.
		// Scoped to ReDim because in other leaf statements the token after `As`
		// IS a genuine value read (`Open path For Output As fnum`, `Name a As b`).
		if (statementHead === 'redim' && word === 'as' && isPotentialVariableReferenceToken(toks[i + 1])) {
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
	// A bare module name in front of a dot is only a legal receiver when the
	// module IS a value: a standard module is a namespace, and documents,
	// forms, and classes marked `VB_PredeclaredId = True` have a default
	// instance. A plain class module is a TYPE, so `Ticket.ChangeTest` is the
	// same `Variable not defined` as any other undeclared name - the VBE
	// refuses to compile it (issue #47). Only a VOUCHED-FOR false reports:
	// the attribute is invisible in the code pane, so a host that never read
	// the header leaves this undefined and the name stays skipped.
	if (surface.kind === 'class' && surface.predeclaredId === false) {
		return false;
	}
	return surface.members.some((candidate) => candidate.name.toLowerCase() === memberLower);
}

/** Words `Open ... For <mode>` accepts that are not reserved identifiers. */
const OPEN_MODE_WORDS: ReadonlySet<string> = new Set([
	'binary', 'output', 'append', 'random', 'read',
]);

/** Clause keywords an `Open` mode word may follow. */
const OPEN_CLAUSE_HEADS: ReadonlySet<string> = new Set(['for', 'access', 'lock']);

/** Tokens after which a new statement begins on the same line. */
const STATEMENT_OPENERS: ReadonlySet<string> = new Set(['then', 'else', ':']);

/**
 * True when a contextual keyword sits in the grammar position that gives it its
 * keyword meaning, rather than naming a variable.
 *
 * Only the positions reachable inside a PROCEDURE BODY are listed. `Option`,
 * `Declare`, and `Property` headers never reach the reference scanner - the
 * first two are module-level and an in-procedure declaration is skipped whole -
 * so guarding them here would be dead code.
 */
function isContextualGrammarWord(
	toks: readonly VbaToken[],
	index: number,
	statementHead: string,
	firstExecutable: number,
): boolean {
	const word = tokenText(toks[index]);
	const prev = tokenText(toks[index - 1]);
	// `Exit Property` and the `End Property` footer. Sub, Function, For and Do
	// are reserved, so Property is the only one of these words that reaches the
	// scanner at all.
	if (word === 'property' && (prev === 'exit' || prev === 'end')) {
		return true;
	}
	// `On Error GoTo/Resume ...` and the bare `Error <number>` statement. The
	// latter starts a statement, which is not always token 0: `If x Then Error
	// 5 Else Exit Sub` puts it after `Then`.
	if (word === 'error' && (prev === 'on' || index === firstExecutable || STATEMENT_OPENERS.has(prev))) {
		return true;
	}
	// `Open path For Binary|Output|Append|Random [Access Read] [Lock Read] As #n`.
	// The mode word follows the clause keyword that introduces it; the same
	// words elsewhere in the statement stay readable.
	if (statementHead === 'open' && OPEN_MODE_WORDS.has(word) && OPEN_CLAUSE_HEADS.has(prev)) {
		return true;
	}
	// `For i = 1 To 10 Step 2`.
	if (word === 'step' && statementHead === 'for') {
		return true;
	}
	// `As Object` in any statement that reaches here, such as the type clause of
	// a `ReDim ... As Object`.
	if (word === 'object' && prev === 'as') {
		return true;
	}
	return false;
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
	// A name that SPELLS a contextual keyword is still a name, and this skip is
	// what stops the assignment TARGET also being counted as a read (#46).
	// Bracketed names are deliberately NOT included: `[If] = x` is Excel's
	// Evaluate shorthand rather than a variable, a separate question this skip
	// must not answer as a side effect.
	const nameTok = toks[start];
	return nameTok && (nameTok.kind === 'identifier' || isNonReservedKeyword(nameTok))
		? start
		: -1;
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
	if (!tok) {
		return false;
	}
	if (tok.kind === 'identifier' || tok.kind === 'bracketedIdentifier') {
		return true;
	}
	return isNonReservedKeyword(tok);
}

/**
 * A word the VBE capitalizes in its statement context but MS-VBAL 3.3.5.2 does
 * not reserve, so `Dim Text As String` and `Dim Error As Long` are both legal
 * and a reference to one is a real variable read.
 */
function isNonReservedKeyword(tok: VbaToken | undefined): boolean {
	if (tok?.kind !== 'keyword' || isReservedIdentifier(tok.rawText)) {
		return false;
	}
	// `Property` opens a declaration, and the parser already refuses `Dim
	// Property As Long` with invalid-identifier-start, so the word can never
	// reach a body as a variable. Scanning it only adds a second, confusing
	// diagnostic to source that is already reported as malformed.
	return tokenText(tok) !== 'property';
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
