// Rule family: assignment validation (audit #0).
//
// Extracted verbatim from analyzeModule.ts: Const reassignment, scalar and
// member assignment type compatibility, Set assignment validation, and
// missing Function/Property Get return assignments.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	matchParenFrom,
	splitTopLevelTokenGroups,
} from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	ModuleNode,
	ProcedureNode,
	Span,
} from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
	VbaSymbol,
} from '../../symbols/symbolModel';
import {
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import {
	type CallableParamType,
	type CallableTypeSignature,
	type CallArguments,
	extractCall,
	extractQualifiedCall,
} from '../callExtraction';
import {
	buildModuleTypeSignatures,
	callableSignatureForCall,
	callableTypeSignaturesFor,
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	declaredTypeForSourceBinding,
	type DeclaredValueShape,
	incompatibilityReason,
	objectLetAssignmentVerdict,
	inferArgumentType,
	isKnownObjectAssignmentType,
	isKnownScalarType,
	isMemberStatementChainThrough,
	namedArgumentSlot,
	nonnumericStringArithmeticOperand,
	normalizeType,
	objectAssignmentIncompatibilityReason,
	resolveExactMemberCompletion,
	sourceBindingTypeResolvers,
	type SourceDeclaredShape,
	type SourceDeclaredTypeResolver,
	sourceIdentifierBinding,
	type SourceNameScope,
	sourceNameScopeFor,
	type SourceQualifiedDeclaredTypeResolver,
	typeEnvironmentFor,
} from '../typeInference';
import {
	activeModuleMembers,
	bareAssignmentTarget,
	declaredNameSpan,
	firstExecutableTokenIndex,
	forEachStatement,
	setAssignmentTarget,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	stripHeaderBrackets,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
	type ProcedureStatementVisitor,
} from '../walker';

/**
 * Rule: assigning to a constant is illegal. High-confidence form only - the
 * left-hand side must be a bare identifier (no member access, no index) that
 * resolves to a Const declared at module level or in the enclosing procedure.
 */
/** A literal that can never be an object reference: a number, string, date, True or False. */
function isScalarLiteralToken(tok: VbaToken): boolean {
	if (tok.kind === 'integerLiteral' || tok.kind === 'floatLiteral' || tok.kind === 'stringLiteral' || tok.kind === 'dateLiteral') {
		return true;
	}
	const word = tokenText(tok);
	return word === 'true' || word === 'false';
}

export function checkConstAssignment(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkSpan(span);
			}
		};

		function checkSpan(span: Span): void {
			const hit = bareAssignmentTarget(source, span);
			if (!hit) {
				return;
			}
			const binding = sourceIdentifierBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				hit.name,
				'assignmentTarget',
			);
			if (
				binding.scope !== 'ambiguous' &&
				binding.definitions.some((definition) => definition.kind === 'constant')
			) {
				push(
					'constAssignment',
					`Cannot assign to constant '${hit.name}'.`,
					hit.span,
				);
			}
		}
	};
}

function memberAssignmentTarget(
	source: string,
	span: Span,
): {
	member: string;
	label: string;
	memberSpan: Span;
	valueTokens: VbaToken[];
	usesSet: boolean;
} | undefined {
	const toks = statementTokens(source, span);
	let i = firstExecutableTokenIndex(toks);
	const usesSet = tokenText(toks[i]) === 'set';
	if (usesSet || tokenText(toks[i]) === 'let') {
		i++;
	}
	const eq = topLevelOperatorIndex(toks.slice(i), '=');
	if (eq < 0) {
		return undefined;
	}
	const equalsIndex = i + eq;
	const lhs = toks.slice(i, equalsIndex);
	if (lhs.length < 2) {
		return undefined;
	}
	const memberTok = lhs[lhs.length - 1];
	if (!tokenName(memberTok) || lhs[lhs.length - 2]?.rawText !== '.') {
		return undefined;
	}
	// A target is one receiver chain ending in the member. Anything else
	// before the `=` is another statement comparing the member: an ElseIf or
	// Case header, a single-line If's condition, a call given the comparison
	// (`Debug.Print w.Part = "a"`). ReDim's `ElseIf ReDimUI.SenderPart =
	// "plus" Then` compiles, and was reported as assigning to 'ElseIf
	// ReDimUI.SenderPart'.
	if (!isMemberStatementChainThrough(lhs, 0, lhs.length - 1)) {
		return undefined;
	}
	if (lhs.some((tok) => tok.kind === 'operator' && tok.rawText === '=')) {
		return undefined;
	}
	return {
		member: tokenName(memberTok)!,
		label: source
			.slice(span.start + lhs[0].start, span.start + memberTok.end)
			.trim(),
		memberSpan: {
			start: span.start + memberTok.start,
			end: span.start + memberTok.end,
		},
		valueTokens: toks.slice(equalsIndex + 1),
		usesSet,
	};
}

export function checkAssignmentTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const { resolveExpressionType, resolveQualifiedExpressionType } =
			sourceBindingTypeResolvers(symbols, procSym, projectVisibleSymbols);
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkAssignmentSpan(span);
			}
		}, activity);

		function checkAssignmentSpan(span: Span): void {
			const assignment = bareAssignmentTarget(source, span);
			if (!assignment) {
				return;
			}
			const targetType = declaredTypeForSourceBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				assignment.name,
				'assignmentTarget',
			);
			const expected = targetType.resolved
				? targetType.asType
				: env.get(assignment.name.toLowerCase());
			if (!expected) {
				return;
			}
			if (isKnownObjectAssignmentType(expected, memberCtx)) {
				// The VBE compiles a bare `=` to an object variable as a Let
				// through the type's default member (issue #107): `r = 5`
				// writes the Range's Value. What is reported is what the
				// default member makes of it.
				const verdict = objectLetAssignmentVerdict(expected, memberCtx);
				if (verdict === 'argument') {
					push(
						'setRequired',
						`Assignment to '${assignment.name}' requires Set: the default member of ${expected} takes an argument, so a Let cannot reach it. This is a VBE compile error: Argument not optional.`,
						assignment.span,
					);
				} else if (verdict === 'noDefault') {
					push(
						'setRequired',
						`Assignment to '${assignment.name}' requires Set: ${expected} has no default member for a Let to reach. This will raise Run-time error '438': Object doesn't support this property or method.`,
						assignment.span,
					);
				}
				return;
			}
			const arraySource = arrayAssignmentToScalarSource(
				assignment,
				span.start,
				expected,
				shapes,
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'assignmentTarget',
				),
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'expression',
				),
			);
			if (arraySource) {
				push(
					'arrayAssignmentToScalar',
					`Array variable '${arraySource.name}' cannot be assigned to scalar '${assignment.name}'. Assign an array element or use a Variant/array target.`,
					arraySource.span,
				);
				return;
			}
			// A dynamic Byte array takes a String whole - `b = "abc"` copies the
			// string's bytes, and a String takes the array back (issue #105,
			// measured in Excel 16.0). The element type is not what the value
			// is checked against there.
			const resolvedTargetShape = declaredShapeForSourceBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				assignment.name,
				'assignmentTarget',
			);
			const targetShape = resolvedTargetShape.resolved
				? resolvedTargetShape.shape
				: shapes.get(assignment.name.toLowerCase());
			if (targetShape?.isArray && normalizeType(targetShape.asType) === 'byte') {
				return;
			}
			const stringArithmetic = nonnumericStringArithmeticOperand(
				expected,
				assignment.valueTokens,
				span.start,
			);
			if (stringArithmetic) {
				push(
					'stringArithmeticCoercion',
					`Assignment to '${assignment.name}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
					stringArithmetic.span,
				);
				return;
			}
			const actual = inferArgumentType(
				assignment.valueTokens,
				span.start,
				env,
				moduleSignatures,
				sourceNames,
				source,
				memberCtx,
				resolveExpressionType,
				resolveQualifiedExpressionType,
			);
			if (!actual) {
				return;
			}
			const reason = incompatibilityReason(expected, actual);
			if (!reason) {
				return;
			}
			push(
				'assignmentTypeMismatch',
				`Assignment to '${assignment.name}' expects ${expected}, but got ${actual.label}. ${reason}`,
				actual.span,
			);
		}
		checkMemberAssignmentTypes(
			source,
			member,
			env,
			moduleSignatures,
			sourceNames,
			memberCtx,
			activity,
			push,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
	}
}

function arrayAssignmentToScalarSource(
	assignment: { name: string; valueTokens: VbaToken[] },
	baseOffset: number,
	expectedType: string,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveTargetShape?: (name: string) => SourceDeclaredShape,
	resolveSourceShape?: (name: string) => SourceDeclaredShape,
): { name: string; span: Span } | undefined {
	const targetScalarType = knownScalarAssignmentTargetType(
		assignment.name,
		expectedType,
		shapes,
		resolveTargetShape,
	);
	if (!targetScalarType) {
		return undefined;
	}
	if (assignment.valueTokens.length !== 1) {
		return undefined;
	}
	const tok = assignment.valueTokens[0];
	const sourceName = tokenName(tok);
	if (!sourceName) {
		return undefined;
	}
	const resolvedSourceShape = resolveSourceShape?.(sourceName);
	const sourceShape = resolvedSourceShape?.resolved
		? resolvedSourceShape.shape
		: shapes.get(sourceName.toLowerCase());
	if (!sourceShape?.isArray) {
		return undefined;
	}
	// VBA special case (MS-VBAL Let-statement rules): a Byte array is directly
	// assignable to a String scalar - the idiomatic encoding-conversion pattern
	// (`s = bytes`). Only Byte element types are exempt; every other element
	// type remains a compile error.
	if (targetScalarType === 'string' && normalizeType(sourceShape.asType) === 'byte') {
		return undefined;
	}
	return {
		name: sourceName,
		span: { start: baseOffset + tok.start, end: baseOffset + tok.end },
	};
}

function knownScalarAssignmentTargetType(
	name: string,
	expectedType: string,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveShape?: (name: string) => SourceDeclaredShape,
): string | undefined {
	const resolvedTargetShape = resolveShape?.(name);
	const targetShape = resolvedTargetShape?.resolved
		? resolvedTargetShape.shape
		: shapes.get(name.toLowerCase());
	if (targetShape?.isArray) {
		return undefined;
	}
	const normalized = normalizeType(targetShape?.asType ?? expectedType);
	return normalized && isKnownScalarType(normalized) ? normalized : undefined;
}

/**
 * Rule: a Function/Property Get returns through its hidden return variable.
 * Falling through without assigning that variable is legal VBA, but it silently
 * returns the default value, and the VBE says nothing at all.
 *
 * Typed returns were skipped for a long time, on the reasoning that a typed
 * default might be intentional. The case that matters says otherwise: a
 * `Function ... As Double` that never names itself hands 0 to every caller, and
 * that surfaces as a wrong number rather than an error. Widening it was measured
 * over 67 modules of third-party code - 40 findings, almost all false - so it
 * arrived with the three things that made those false: a return whose FIELDS are
 * assigned counts, a class's empty body is an unimplemented interface member,
 * and a body whose work is to raise owes no return. That leaves 2.
 */
export function checkMissingReturnAssignments(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	activity: ConditionalActivityTracker | undefined,
	moduleName: string | undefined,
	implementedInterfaces: ReadonlySet<string> | undefined,
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const isInterface = moduleName !== undefined
		&& implementedInterfaces?.has(moduleName.toLowerCase()) === true;
	for (const member of activeModuleMembers(mod, activity)) {
		if (
			member.kind !== 'Procedure' ||
			(member.procKind !== 'Function' && member.procKind !== 'PropertyGet')
		) {
			continue;
		}
		if (!member.closed) {
			continue;
		}
		if (procedureHasReturnAssignment(source, member, activity, moduleSignatures)) {
			continue;
		}
		if (returnIsNotExpected(source, member, activity, isInterface)) {
			continue;
		}
		const procLabel = member.procKind === 'PropertyGet' ? 'Property Get' : 'Function';
		push(
			'missingReturnAssignment',
			`${procLabel} '${member.name}' has no return assignment; VBA will return the default value. Assign to '${member.name}' before exit if a value is intended.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function procedureHasReturnAssignment(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	const lower = proc.name.toLowerCase();
	let found = false;
	const assignsIn = (span: { start: number; end: number }): boolean => {
		const bare = bareAssignmentTarget(source, span);
		if (bare?.name.toLowerCase() === lower) {
			return true;
		}
		const set = setAssignmentTarget(source, span);
		if (set?.name.toLowerCase() === lower) {
			return true;
		}
		if (returnAssignedByStatementForm(source, span, lower)) {
			return true;
		}
		const call = extractCall(source, span);
		const qualifiedCall = call
			? undefined
			: extractQualifiedCall(source, span, moduleSignatures);
		const effectiveCall = call ?? qualifiedCall;
		return Boolean(
			effectiveCall && callPassesNameToByRefParam(effectiveCall, lower, moduleSignatures),
		);
	};
	forEachStatement(proc.body, (stmt) => {
		if (found) {
			return;
		}
		// The branch spans cover a single-line If, whose statements the walk
		// itself does not reach (issue #46).
		for (const span of statementAndBranchSpans(stmt)) {
			if (assignsIn(span) || assignsOwnField(source, span, lower)) {
				found = true;
				return;
			}
		}
	}, activity);
	// `For Count3 = 1 To 3` assigns the return variable as its counter (issue
	// #115): the loop is a block, not a statement the walk above visits.
	return found || forLoopAssigns(proc.body, lower, activity);
}

function forLoopAssigns(
	body: readonly BodyNode[],
	lower: string,
	activity: ConditionalActivityTracker | undefined,
): boolean {
	for (const node of body) {
		if (activity?.isInactive(node.span)) {
			continue;
		}
		if (node.kind === 'ForBlock' && node.controlVariable?.toLowerCase() === lower) {
			return true;
		}
		if ('body' in node && Array.isArray(node.body) && forLoopAssigns(node.body, lower, activity)) {
			return true;
		}
	}
	return false;
}

/**
 * The statement forms besides `Name = value` that assign a Function's return
 * variable (issue #115, each measured in Excel 16.0): `For Name = 1 To 3`
 * leaves the counter's final value, `ReDim Name(2)` sizes an array return,
 * `Line Input #f, Name`, `Input #f, Name` and `Get #f, 1, Name` read into it,
 * and `Name$ = "hi"` names it with its type-declaration character.
 */
function returnAssignedByStatementForm(source: string, span: Span, lower: string): boolean {
	const toks = statementTokens(source, span);
	const i = firstExecutableTokenIndex(toks);
	const head = tokenText(toks[i]);
	const isName = (tok: VbaToken | undefined): boolean => tokenName(tok)?.toLowerCase() === lower;
	if (head === 'for') {
		return tokenText(toks[i + 1]) === 'each' ? isName(toks[i + 2]) : isName(toks[i + 1]);
	}
	if (head === 'redim') {
		let k = i + 1;
		if (tokenText(toks[k]) === 'preserve') {
			k++;
		}
		let depth = 0;
		for (; k < toks.length; k++) {
			const raw = toks[k].rawText;
			if (raw === '(') {
				depth++;
			} else if (raw === ')') {
				depth--;
			} else if (depth === 0 && isName(toks[k]) && (k === i + 1 || tokenText(toks[k - 1]) === 'preserve' || toks[k - 1].rawText === ',')) {
				return true;
			}
		}
		return false;
	}
	if (head === 'line' || head === 'input' || head === 'get') {
		// Everything after the file number is a target (Line Input / Input) or
		// the third slot is (Get #f, rec, var); a plain name in one of them
		// is the assignment.
		return toks.slice(i + 1).some((tok, index, rest) => isName(tok) && (rest[index - 1]?.rawText === ',' ));
	}
	// `Name$ = value`: the suffix is glued to the name and the `=` follows.
	if (isName(toks[i]) && toks[i + 1] && toks[i + 1].start === toks[i].end
		&& /^[$%&!#@]$/.test(toks[i + 1].rawText) && toks[i + 2]?.rawText === '=') {
		return true;
	}
	return false;
}

/**
 * True for `Name.Field = value`, which fills in a UDT or object return field by
 * field. `utc_DateToSystemTime.utc_wYear = ...` IS the return assignment, and
 * reading only bare `Name =` counted 16 such functions in the test corpus as
 * never assigning anything.
 */
function assignsOwnField(source: string, span: Span, lower: string): boolean {
	const toks = statementTokens(source, span);
	let i = firstExecutableTokenIndex(toks);
	if (toks[i] && tokenText(toks[i]).toLowerCase() === 'let') {
		i += 1;
	}
	const name = toks[i];
	if (!name || name.rawText.toLowerCase() !== lower || toks[i + 1]?.rawText !== '.') {
		return false;
	}
	return toks.slice(i + 2).some((token) => token.kind === 'operator' && token.rawText === '=');
}

/**
 * True for a procedure that is not expected to assign a return: an empty body in
 * a CLASS (a member of an interface the class declares and does not implement)
 * or one whose work is to raise. Measured on the test corpus these account for
 * 16 of the 40 typed functions that never name themselves, every one deliberate.
 */
function returnIsNotExpected(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	isInterface: boolean,
): boolean {
	let executable = 0;
	let raises = false;
	forEachStatement(proc.body, (stmt) => {
		const text = source.slice(stmt.span.start, stmt.span.end).trim();
		if (!text || text.startsWith("'") || /^(Dim|Const|Static|ReDim)\b/i.test(text)) {
			return;
		}
		executable += 1;
		if (/\bErr\s*\.\s*Raise\b/i.test(text) || /^Error\s/i.test(text)) {
			raises = true;
		}
	}, activity);
	// An empty body is a stub where the module is a contract: a class that
	// another module declares with `Implements` states its members for the
	// implementer to fill in, so every one of them is empty on purpose
	// (stdICallable and stdEnumerator in the test corpus are exactly that).
	// An empty Function anywhere else is unfinished code and still reports.
	//
	// This used to ask the PARSED module kind, which only says `class` when the
	// source carries `Attribute VB_Exposed` and friends. Module text read out of
	// a project carries no attribute lines, so it never said `class` for a real
	// class module and the carve-out never fired
	// (github.com/WilliamSmithEdward/xlide_vscode/issues/60).
	return (executable === 0 && isInterface) || raises;
}

function callPassesNameToByRefParam(
	call: CallArguments,
	lowerName: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	const sig = callableSignatureForCall(call, moduleSignatures);
	if (!sig) {
		return false;
	}
	let positionalIndex = 0;
	for (const slot of call.slots) {
		const named = namedArgumentSlot(slot);
		let param: CallableParamType | undefined;
		let valueSlot = slot;
		if (named) {
			param = sig.params.find((p) => stripHeaderBrackets(p.name).toLowerCase() === named.name.toLowerCase());
			valueSlot = named.value;
		} else {
			param = sig.params[Math.min(positionalIndex, sig.params.length - 1)];
			positionalIndex++;
		}
		if (!param?.byRef || !singleSlotNameEquals(valueSlot, lowerName)) {
			continue;
		}
		return true;
	}
	return false;
}

function singleSlotNameEquals(slot: readonly VbaToken[], lowerName: string): boolean {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return toks.length === 1 && tokenName(toks[0])?.toLowerCase() === lowerName;
}

function checkMemberAssignmentTypes(
	source: string,
	member: ProcedureNode,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): void {
	if (!memberCtx.projectClassMembers || memberCtx.projectClassMembers.length === 0) {
		return;
	}
	const checkStatement = (span: Span): void => {
		const assignment = memberAssignmentTarget(source, span);
		if (!assignment) {
			return;
		}
		const target = resolveExactMemberCompletion(
			source,
			assignment.member,
			assignment.memberSpan.end,
			memberCtx,
		);
		if (!target || target.writable === undefined) {
			return;
		}
		if (target.writable === false) {
			push(
				'readonlyMemberAssignment',
				`Cannot assign to read-only property '${assignment.label}'.`,
				assignment.memberSpan,
			);
			return;
		}
		const expected = target.writeType ?? target.returns;
		if (assignment.usesSet) {
			if (expected && isKnownScalarType(normalizeType(expected) ?? '')) {
				push(
					'setRequiresObject',
					`Set assignment requires an object-valued target, but '${assignment.label}' expects ${expected}.`,
					assignment.memberSpan,
				);
				return;
			}
			// `Set h.Item = x` needs a Property Set; with only a Property Let
			// the VBE refuses it, "Invalid use of property" (issue #107).
			if (target.letAccessor && !target.setAccessor) {
				push(
					'setRequiresObject',
					`Set assignment to '${assignment.label}' needs a Property Set, but the property declares only a Property Let. This is a VBE compile error: Invalid use of property.`,
					assignment.memberSpan,
				);
				return;
			}
			const actual = inferArgumentType(
				assignment.valueTokens,
				span.start,
				env,
				moduleSignatures,
				sourceNames,
				source,
				memberCtx,
				resolveExpressionType,
				resolveQualifiedExpressionType,
			);
			const reason = objectAssignmentIncompatibilityReason(
				expected,
				actual,
				memberCtx,
			);
			if (reason) {
				push(
					'assignmentObjectTypeMismatch',
					`Object assignment to '${assignment.label}' expects ${expected}, but got ${actual?.label}. ${reason}`,
					actual?.span ?? assignment.memberSpan,
				);
			}
			return;
		}
		// A bare `=` to a project property calls its Property Let, whatever the
		// value's type: `h.Item = New Collection` compiles with `Property Let
		// Item(ByVal v As Object)` (issue #107). Only a property with a Set
		// and no Let refuses it: "Invalid use of property".
		if (target.setAccessor && !target.letAccessor) {
			push(
				'setRequired',
				`Assignment to '${assignment.label}' requires Set: the property declares a Property Set and no Property Let. This is a VBE compile error: Invalid use of property.`,
				assignment.memberSpan,
			);
			return;
		}
		if (!target.letAccessor && isKnownObjectAssignmentType(expected, memberCtx)) {
			push(
				'setRequired',
				`Object assignment to '${assignment.label}' requires Set because it expects ${expected}.`,
				assignment.memberSpan,
			);
			return;
		}
		if (!expected || !isKnownScalarType(normalizeType(expected) ?? '')) {
			return; // a Let of an object or unknown type: nothing provable about the value
		}
		const stringArithmetic = nonnumericStringArithmeticOperand(
			expected,
			assignment.valueTokens,
			span.start,
		);
		if (stringArithmetic) {
			push(
				'stringArithmeticCoercion',
				`Assignment to '${assignment.label}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
				stringArithmetic.span,
			);
			return;
		}
		const actual = inferArgumentType(
			assignment.valueTokens,
			span.start,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
		if (!actual) {
			return;
		}
		const reason = incompatibilityReason(expected, actual);
		if (!reason) {
			return;
		}
		push(
			'assignmentTypeMismatch',
			`Assignment to '${assignment.label}' expects ${expected}, but got ${actual.label}. ${reason}`,
			actual.span,
		);
	};
	// This rule reads a statement structurally - what precedes its first `=`
	// is the target - so it takes a single-line If's branches as statements of
	// their own. Read whole, `If ok Then w.Part = 1` had the target
	// `If ok Then w.Part`, and `If w.Part = 1 Then Exit Sub`, which assigns
	// nothing, had `If w.Part`.
	forEachStatement(member.body, (stmt) => {
		for (const span of statementAndBranchSpans(stmt)) {
			checkStatement(span);
		}
	}, activity);
}

export function checkSetAssignments(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const { resolveExpressionType, resolveQualifiedExpressionType } =
			sourceBindingTypeResolvers(symbols, procSym, projectVisibleSymbols);
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkSetSpan(span);
			}
		};

		function checkSetSpan(span: Span): void {
			const target = setAssignmentTarget(source, span);
			if (!target) {
				return;
			}
			const targetDeclaredType = declaredTypeForSourceBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				target.name,
				'assignmentTarget',
			);
			const expected = targetDeclaredType.resolved
				? targetDeclaredType.asType
				: env.get(target.name.toLowerCase());
			const targetType = normalizeType(expected);
			// `Set v = 5` is refused whatever v is: a literal is never an object
			// reference ("Object required", issue #125, measured in Excel 16.0).
			const literal = target.valueTokens.filter((tok) => tok.kind !== 'comment');
			if ((!targetType || targetType === 'variant') && literal.length === 1 && isScalarLiteralToken(literal[0])) {
				push(
					'setRequiresObject',
					`Set assigns an object reference, but ${literal[0].rawText} is a literal value. This is a VBE compile error: Object required.`,
					{ start: span.start + literal[0].start, end: span.start + literal[0].end },
				);
				return;
			}
			if (!targetType || !isKnownScalarType(targetType)) {
				if (!isKnownObjectAssignmentType(expected, memberCtx)) {
					return;
				}
				const actual = inferArgumentType(
					target.valueTokens,
					span.start,
					env,
					moduleSignatures,
					sourceNames,
					source,
					memberCtx,
					resolveExpressionType,
					resolveQualifiedExpressionType,
				);
				const reason = objectAssignmentIncompatibilityReason(
					expected,
					actual,
					memberCtx,
				);
				if (reason) {
					push(
						'assignmentObjectTypeMismatch',
						`Object assignment to '${target.name}' expects ${expected}, but got ${actual?.label}. ${reason}`,
						actual?.span ?? target.span,
					);
				}
				return;
			}
			push(
				'setRequiresObject',
				`Set assignment requires an object variable, but '${target.name}' is declared as ${expected}.`,
				target.span,
			);
		}
	};
}

/**
 * Rule: the target of a `Mid`/`Mid$`/`MidB`/`MidB$` replacement statement
 * (MS-VBAL §5.4.3.4) must be a writable String variable. A string-literal
 * target - `Mid$("abc", 2, 3) = "XY"` - is a compile error (oracle-verified
 * `mid_stmt_literal_target_probe`, `mid_stmt_no_suffix_literal_target_probe`,
 * `midb_stmt_literal_target_probe`; the variable-target control
 * `mid_stmt_variable_target_probe` is accepted).
 *
 * No-FP scope: fires only when a statement's first executable token is
 * `mid`/`midb` (optionally followed by the `$` type-character), immediately
 * followed by `(`, whose matching `)` is followed by a top-level `=`, and whose
 * first argument slot is exactly one string-literal token.
 *
 * Conservative shadowing guard: if the module names `mid`/`midb` in ANY
 * declaration form - a symbol-table entry (Dim/Static/Const/parameter/Function/
 * Property) or an implicit `ReDim` array declaration - the rule stays silent for
 * the whole module. The oracle confirms no shadow form makes a literal-target
 * Mid valid (`mid_shadow_dim_array_probe`, `mid_shadow_redim_implicit_probe`,
 * `mid_shadow_property_let_probe` all compile-error), so this guard is belt-and-
 * suspenders: it cannot prevent a real false positive, but it keeps the rule from
 * touching any module that even mentions a user `Mid` without a binder.
 */
export function checkMidStatementLiteralTarget(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (moduleShadowsMidIntrinsic(symbols) || moduleRedimDeclaresMidIntrinsic(source, mod, activity)) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			const hit = midStatementLiteralTargetViolation(source, stmt.span);
			if (hit) {
				push('midStatementLiteralTarget', hit.message, hit.span);
			}
		}, activity);
	}
}

/** Suffix-stripped, lower-cased word for a token (keyword or identifier). */
function midBaseWord(tok: VbaToken | undefined): string {
	if (!tok) {
		return '';
	}
	return (tokenName(tok) ?? tok.rawText).toLowerCase().replace(/[$%&!#@]$/, '');
}

/** True when a module declares any symbol that shadows the Mid/MidB intrinsic. */
function moduleShadowsMidIntrinsic(symbols: ReturnType<typeof buildModuleSymbols>): boolean {
	return symbols.all.some((sym) => {
		const base = sym.name.toLowerCase().replace(/[$%&!#@]$/, '');
		return base === 'mid' || base === 'midb';
	});
}

/**
 * True when a module implicitly declares an array named `mid`/`midb` via a
 * `ReDim` with no prior `Dim` (the symbol builder models declared variable groups
 * only, so a ReDim-only name is absent from `symbols.all`).
 */
function moduleRedimDeclaresMidIntrinsic(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): boolean {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		let found = false;
		forEachStatement(member.body, (stmt) => {
			if (found) {
				return;
			}
			const toks = statementTokensAfterLeadingLabel(source, stmt.span);
			if (midBaseWord(toks[0]) !== 'redim') {
				return;
			}
			const start = midBaseWord(toks[1]) === 'preserve' ? 2 : 1;
			for (const group of splitTopLevelTokenGroups(toks, start, ',')) {
				const base = midBaseWord(group[0]);
				if (base === 'mid' || base === 'midb') {
					found = true;
					return;
				}
			}
		}, activity);
		if (found) {
			return true;
		}
	}
	return false;
}

function midStatementLiteralTargetViolation(
	source: string,
	span: Span,
): { span: Span; message: string } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (toks.length === 0) {
		return undefined;
	}
	// Strip a trailing type-character so both lexings of `Mid$` are handled: a
	// single `Mid$` token, or `Mid` followed by a separate `$` token (below).
	const head = midBaseWord(toks[0]);
	if (head !== 'mid' && head !== 'midb') {
		return undefined;
	}
	let parenIndex = 1;
	if (toks[parenIndex]?.rawText === '$') {
		parenIndex = 2;
	}
	if (toks[parenIndex]?.rawText !== '(') {
		return undefined;
	}
	const close = matchParenFrom(toks, parenIndex);
	if (close <= parenIndex + 1) {
		return undefined; // empty or unbalanced argument list
	}
	// The Mid replacement-statement form: the matching `)` is followed by `=`.
	if (toks[close + 1]?.rawText !== '=') {
		return undefined;
	}
	const argToks = toks.slice(parenIndex + 1, close).filter((tok) => tok.kind !== 'comment');
	const slots = splitTopLevelTokenGroups(argToks, 0, ',');
	const target = slots[0];
	if (!target || target.length !== 1 || target[0].kind !== 'stringLiteral') {
		return undefined; // target is not exactly one string literal
	}
	return {
		span: { start: span.start + target[0].start, end: span.start + target[0].end },
		message:
			"The target of a Mid statement must be a writable String variable, not a " +
			'string literal. Assigning into a literal is a compile error.',
	};
}
