// Rule family: assignment validation (audit #0).
//
// Extracted verbatim from analyzeModule.ts: Const reassignment, scalar and
// member assignment type compatibility, Set assignment validation, and
// missing Function/Property Get return assignments.

import {
	isLateBoundTypeKey,
	resolveReceiverTypeAt,
	signatureDeclaresParameters,
	type MemberCompletion,
	type MemberCompletionContext,
} from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { isDispatchOnlyHostType, resolveHostEnum } from '../../host/hostModel';
import { hostPropertyValueProblem } from './hostPropertyValues';
import {
	matchParenFrom,
	splitTopLevelTokenGroups,
} from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	LeafStatementNode,
	ModuleNode,
	ProcedureNode,
	Span,
} from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { objectLetStateAt } from './objectState';
import { elementOperandStartingAt, elementsWrittenIn, knownArrayShapesAt, moduleOptionBase, type FixedArrayBound } from './arrays';
import { functionResultAt, knownFunctionResults } from '../functionResults';
import { straightLineAssignments } from '../straightLineValues';
import { heldObjectsAt } from '../heldObjects';
import { resolveRuntimeFunction } from '../../runtime/vbaRuntime';
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
	type InferredArgumentType,
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
	objectHoldingDefault,
	objectLetAssignmentVerdict,
	inferArgumentType,
	isKnownObjectAssignmentType,
	isKnownScalarType,
	isMemberStatementChainThrough,
	knownLocalLiteralValuesAt,
	type KnownLocalValue,
	namedArgumentSlot,
	nonnumericStringArithmeticOperand,
	numericLiteralBounds,
	normalizeType,
	objectAssignmentIncompatibilityReason,
	SCALAR_OBJECT_ASSIGNMENT_REASON,
	resolveExactMemberCompletion,
	runtimeCallableSourceShadowed,
	sourceBindingTypeResolvers,
	type SourceDeclaredShape,
	type SourceDeclaredTypeResolver,
	sourceIdentifierBinding,
	type SourceNameScope,
	sheetsFromCollectionProperty,
	sourceNameScopeFor,
	type SourceQualifiedDeclaredTypeResolver,
	stringLiteralValue,
	typeEnvironmentFor,
	unreachableStatementsIn,
	unwrapOuterParens,
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
import { nameMentions } from './shared';

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
			// `Set K = Nothing` assigns to the constant too (issue #255).
			const hit = bareAssignmentTarget(source, span) ?? setAssignmentTarget(source, span);
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
			if (binding.scope === 'ambiguous') {
				return;
			}
			// An Enum member is a constant too (issue #213): `eA = 2` is
			// "Assignment to constant not permitted".
			if (binding.definitions.some((definition) => definition.kind === 'constant' || definition.kind === 'enumMember')) {
				push(
					'constAssignment',
					`Cannot assign to constant '${hit.name}'.`,
					hit.span,
				);
				return;
			}
			const target = procedureAssignmentTarget(binding.definitions, procSym);
			if (target) {
				push('assignmentToProcedureName', target(hit.name), hit.span);
			}
		}
	};
}

/**
 * Why a procedure's name cannot be assigned to from outside it, measured in
 * Excel 16.0 (issue #213): a Sub's is "Expected Function or variable", and a
 * Function's is "Function call on left-hand side of assignment must return
 * Variant or Object" when it returns a type of VBA's own. Inside the Function
 * the name is its return value and binds locally, so it never reaches here.
 */
/** The default member of a project class when it is a Property Get with no Property Let. */
function readOnlyProjectDefault(type: string, memberCtx: MemberCompletionContext): string | undefined {
	const lower = type.trim().split('.').pop()?.toLowerCase();
	const cls = (memberCtx.projectClassMembers ?? []).find((candidate) => candidate.kind === 'class' && candidate.name.toLowerCase() === lower);
	const member = cls?.exhaustive === true ? cls.members.find((candidate) => candidate.defaultMember) : undefined;
	return member && member.kind === 'property' && !member.letAccessor && member.writable !== true ? member.name : undefined;
}

function procedureAssignmentTarget(
	definitions: readonly VbaSymbol[],
	procSym: VbaSymbol | undefined,
): ((name: string) => string) | undefined {
	if (definitions.length !== 1 || definitions[0] === procSym) {
		return undefined;
	}
	const [definition] = definitions;
	if (definition.kind === 'sub') {
		return (name) => `'${name}' is a Sub, which has no value to assign. This is a VBE compile error: Expected Function or variable.`;
	}
	const returns = normalizeType(definition.asType);
	// A Declare Function too (issue #254): `GetTickCount = 5`.
	const isFunction = definition.kind === 'function' || (definition.kind === 'declare' && definition.declareKind === 'Function');
	if (isFunction && returns && returns !== 'variant' && isKnownScalarType(returns) && !definition.isArray) {
		return (name) => `'${name}' is a Function returning ${definition.asType}, and a call cannot be assigned to. This is a VBE compile error: Function call on left-hand side of assignment must return Variant or Object.`;
	}
	return undefined;
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
	/** True for `wb.Name() = x`: the member is given arguments. */
	withArguments: boolean;
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
	// The member may be given arguments: `r.Address(False, False) = "B2"`.
	let memberIndex = lhs.length - 1;
	const withArguments = lhs[memberIndex].rawText === ')';
	if (withArguments) {
		let depth = 0;
		for (; memberIndex >= 0; memberIndex--) {
			const raw = lhs[memberIndex].rawText;
			depth += raw === ')' ? 1 : raw === '(' ? -1 : 0;
			if (depth === 0) {
				break;
			}
		}
		memberIndex--;
	}
	const memberTok = lhs[memberIndex];
	if (!memberTok || !tokenName(memberTok) || lhs[memberIndex - 1]?.rawText !== '.') {
		return undefined;
	}
	// A target is one receiver chain ending in the member. Anything else
	// before the `=` is another statement comparing the member: an ElseIf or
	// Case header, a single-line If's condition, a call given the comparison
	// (`Debug.Print w.Part = "a"`). ReDim's `ElseIf ReDimUI.SenderPart =
	// "plus" Then` compiles, and was reported as assigning to 'ElseIf
	// ReDimUI.SenderPart'.
	if (!isMemberStatementChainThrough(lhs, 0, memberIndex)) {
		return undefined;
	}
	if (lhs.some((tok) => tok.kind === 'operator' && tok.rawText === '=')) {
		return undefined;
	}
	return {
		member: tokenName(memberTok)!,
		label: source
			.slice(span.start + lhs[0].start, span.start + lhs[lhs.length - 1].end)
			.trim(),
		memberSpan: {
			start: span.start + memberTok.start,
			end: span.start + memberTok.end,
		},
		valueTokens: toks.slice(equalsIndex + 1),
		usesSet,
		withArguments,
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
	const variantArrayFunctions = arrayOnlyVariantFunctions(source, mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procedure = member;
		const env = typeEnvironmentFor(symbols, member);
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const { resolveExpressionType, resolveQualifiedExpressionType } =
			sourceBindingTypeResolvers(symbols, procSym, projectVisibleSymbols);
		// What a Variant holds at a statement, and how often each name is
		// written anywhere: a Variant named once is never assigned, so Empty.
		let shapesAt: ((stmt: LeafStatementNode) => ReadonlyMap<string, FixedArrayBound>) | undefined;
		let mentions: Map<string, number> | undefined;
		let reaching: ReturnType<typeof straightLineAssignments> | undefined;
		let valuesAt: ((stmt: LeafStatementNode) => ReadonlyMap<string, KnownLocalValue>) | undefined;
		let unreachable: ReadonlySet<BodyNode> | undefined;
		let written: ReadonlySet<string> | undefined;
		const arrayValueAt = (stmt: LeafStatementNode, name: string): ArrayValue | undefined => {
			const lower = name.toLowerCase();
			const local = procSym?.children?.find((child) => child.name.toLowerCase() === lower);
			const type = normalizeType(local?.asType);
			if (local?.kind !== 'localVariable' || local.visibility === 'Static' || local.isArray || (type !== undefined && type !== 'variant')) {
				return undefined;
			}
			const shape = (shapesAt ??= knownArrayShapesAt(source, symbols, procedure, activity, moduleOptionBase(mod, activity)))(stmt).get(lower);
			if (shape) {
				return { element: shape.origin === 'Split(...)' ? 'string' : 'variant', text: `'${name}', which holds an array from ${shape.origin}` };
			}
			mentions ??= nameMentions(source, procedure, activity);
			return mentions.get(lower) === 1 ? { element: 'empty', text: `'${name}', which is never assigned and so is Empty` } : undefined;
		};
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkAssignmentSpan(span, stmt);
			}
		}, activity);

		// `v = Null` then `s = v`: the Null a Variant local holds here, from
		// its last assignment in a straight line (issue #239). A single-line
		// If's branch sees what held before the If, less whatever the If
		// touches.
		function nullHeldAt(stmt: LeafStatementNode, span: Span, valueTokens: readonly VbaToken[]): { name: string; span: Span; returns?: boolean } | undefined {
			const value = valueTokens.filter((tok) => tok.kind !== 'comment');
			// `n = F()` with F a Function of the module returning Null (issue #448).
			const called = value.length > 0 ? functionResultAt(value, 0, knownFunctionResults(source, mod, activity), procedure, symbols) : undefined;
			if (called && called.end === value.length - 1) {
				const at = { start: span.start + value[0].start, end: span.start + value[called.end].end };
				return called.result.kind === 'null' ? { name: source.slice(at.start, at.end), span: at, returns: true } : undefined;
			}
			const name = value.length === 1 ? tokenName(value[0]) : undefined;
			if (!name) {
				return undefined;
			}
			const lower = name.toLowerCase();
			const local = procSym?.children?.find((child) => child.name.toLowerCase() === lower);
			const type = normalizeType(local?.asType);
			if (local?.kind !== 'localVariable' || local.visibility === 'Static' || local.isArray || (type !== undefined && type !== 'variant')) {
				return undefined;
			}
			reaching ??= straightLineAssignments(source, procedure.body, activity);
			const held = reaching.get(stmt)?.get(lower)?.filter((tok) => tok.kind !== 'comment');
			return held?.length === 1 && tokenText(held[0]) === 'null'
				? { name, span: { start: span.start + value[0].start, end: span.start + value[0].end } }
				: undefined;
		}

		// `n = 1 + Null`, `s = "a" + v` with v holding Null: arithmetic, `+`,
		// unary minus, Not, a comparison and Abs give Null when an operand is
		// Null, and so do Xor and Eqv; And, Or and Imp are judged only with
		// Null on every side; `&` never does (issues #324 and #556, measured
		// in Excel 16.0).
		function nullExpressionAt(stmt: LeafStatementNode, span: Span, valueTokens: readonly VbaToken[]): { text: string; span: Span } | undefined {
			const value = unwrapOuterParens(valueTokens.filter((tok) => tok.kind !== 'comment'));
			if (value.length < 2) {
				return undefined;
			}
			const holdsNull = (tok: VbaToken): boolean => tokenText(tok) === 'null'
				|| nullHeldAt(stmt, span, [tok]) !== undefined;
			const yieldsNull = (toks: readonly VbaToken[]): boolean => {
				const part = unwrapOuterParens([...toks]);
				if (part.length === 1) {
					return holdsNull(part[0]);
				}
				const head = tokenText(part[0]);
				if (head === '-' || head === 'not') {
					return yieldsNull(part.slice(1));
				}
				if (head === 'abs' && part[1]?.rawText === '(' && matchParenFrom(part, 1) === part.length - 1) {
					return yieldsNull(part.slice(2, -1));
				}
				const operands: VbaToken[][] = [[]];
				const operators: string[] = [];
				let depth = 0;
				for (const tok of part) {
					depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
					const word = tok.kind === 'operator' ? tok.rawText : tokenText(tok);
					const current = operands[operands.length - 1];
					if (depth === 0 && current.length > 0 && NULL_PROPAGATING.has(word) && tok.kind !== 'stringLiteral') {
						operators.push(word);
						operands.push([]);
					} else {
						current.push(tok);
					}
				}
				if (operators.length === 0 || operators.includes('&') || operands.some((operand) => operand.length === 0)) {
					return false;
				}
				// And, Or and Imp give a value when the other side decides it
				// (issue #556, measured in Excel 16.0): Null And 0 is 0, but
				// Null And 1 is Null; 40000 Or Null is 40000, but 0 Or Null is
				// Null; Null Imp 12 is 12 and False Imp Null is True, but
				// Null Imp False is Null. Judged with one operator only.
				const logical = operators.find((operator) => operator === 'and' || operator === 'or' || operator === 'imp');
				if (logical) {
					if (operators.length !== 1) {
						return operands.every((operand) => yieldsNull(operand));
					}
					const [left, right] = operands.map((operand) => (yieldsNull(operand) ? 'null' : literalNumber(operand)));
					const decided = (other: number | 'null' | undefined, otherOnLeft: boolean): boolean => other === 'null'
						|| (other !== undefined && (logical === 'and' ? other !== 0 : logical === 'or' ? other === 0 : otherOnLeft ? other !== 0 : other === 0));
					return (left === 'null' && decided(right, false)) || (right === 'null' && decided(left, true));
				}
				return operands.some((operand) => yieldsNull(operand));
			};
			return yieldsNull(value)
				? { text: value.map((tok) => tok.rawText).join(' ').replace(/ ?([()]) ?/g, '$1'), span: { start: span.start + value[0].start, end: span.start + value[value.length - 1].end } }
				: undefined;
		}

		// `s = "b"` then `n = s`: the String a local or an array element
		// is known to hold here, from its last assignment in a straight line.
		function knownStringAt(stmt: LeafStatementNode, span: Span, valueTokens: readonly VbaToken[], expected: string): InferredArgumentType | undefined {
			const value = unwrapOuterParens(valueTokens.filter((tok) => tok.kind !== 'comment'));
			if (value.length === 0 || (unreachable ??= unreachableStatementsIn(source, procedure, symbols, activity)).has(stmt)) {
				return undefined;
			}
			const valueSpan = { start: span.start + value[0].start, end: span.start + value[value.length - 1].end };
			const label = `'${source.slice(valueSpan.start, valueSpan.end)}', which holds`;
			// `n = F()` with F a Function of the module returning "abc" (issue #448).
			const called = functionResultAt(value, 0, knownFunctionResults(source, mod, activity), procedure, symbols);
			if (called && called.end === value.length - 1) {
				return called.result.kind === 'string'
					? { type: 'String', label: `'${source.slice(valueSpan.start, valueSpan.end)}', which returns ${JSON.stringify(called.result.value)}`, span: valueSpan, stringValue: called.result.value }
					: undefined;
			}
			if (value.length === 1) {
				const lower = tokenName(value[0])?.toLowerCase();
				const known = lower ? (valuesAt ??= knownLocalLiteralValuesAt(source, procedure, symbols, activity))(stmt).get(lower) : undefined;
				if (known?.kind === 'string' && !known.contentMutated) {
					return { type: 'String', label: `${label} ${JSON.stringify(known.value)}`, span: valueSpan, stringValue: known.value as string };
				}
				// A `String * 3` local named nowhere else holds three Chr(0),
				// which convert to no number, Boolean or date (issue #451,
				// measured in Excel 16.0).
				const local = lower ? procSym?.children?.find((child) => child.name.toLowerCase() === lower) : undefined;
				const length = local?.kind === 'localVariable' && local.visibility !== 'Static' && !local.isArray && /^\d+$/.test(local.fixedLength ?? '') ? Number(local.fixedLength) : undefined;
				if (length !== undefined && length >= 1 && (mentions ??= nameMentions(source, procedure, activity)).get(lower!) === 1) {
					return { type: 'String', label: `'${value[0].rawText}', a String * ${length} never assigned, which holds ${length} Chr(0)`, span: valueSpan, stringValue: '\u0000'.repeat(length) };
				}
				return undefined;
			}
			// `"a" & "b"`, `Left("abc", 1)`, `o & 5` (issue #405). A Date written
			// as text converts back to a Date, so that target is left alone.
			const knownAt = (valuesAt ??= knownLocalLiteralValuesAt(source, procedure, symbols, activity))(stmt);
			const spelled = spelledText(value, (lower) => knownAt.get(lower), env, sourceNames);
			if (spelled) {
				return spelled.standIn && normalizeType(expected) === 'date'
					? undefined
					: { type: 'String', label: spelled.standIn ? `${label} a Date written as text` : spelled.named !== undefined ? `${label} ${spelled.named}` : `${label} ${JSON.stringify(spelled.text)}`, span: valueSpan, stringValue: spelled.text };
			}
			// `v = Array("1", "b")` then `n = v(1)` (issue #260).
			written ??= elementsWrittenIn(source, procedure, activity);
			if (written.has(tokenName(value[0])?.toLowerCase() ?? '')) {
				return undefined;
			}
			const shapes = (shapesAt ??= knownArrayShapesAt(source, symbols, procedure, activity, moduleOptionBase(mod, activity)))(stmt);
			const element = elementOperandStartingAt(value, 0, shapes, moduleOptionBase(mod, activity));
			return element && element.last === value.length - 1 && typeof element.value === 'string'
				? { type: 'String', label: `${label} ${JSON.stringify(element.value)}`, span: valueSpan, stringValue: element.value }
				: undefined;
		}

		function checkAssignmentSpan(span: Span, stmt: LeafStatementNode): void {
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
			// `Dim a()` with no As clause is an array of Variant (issue #222).
			const untypedArray = !targetType.asType && (() => {
				const shape = declaredShapeForSourceBinding(symbols, procSym, projectVisibleSymbols, assignment.name, 'assignmentTarget');
				return shape.resolved && shape.shape?.isArray === true;
			})();
			const declaredExpected = (targetType.resolved
				? targetType.asType
				: env.get(assignment.name.toLowerCase())) ?? (untypedArray ? 'Variant' : undefined);
			// A variable As an Enum is a Long: `x = "abc"` raises 13 and
			// `x = 3000000000#` 6 (issue #436, measured in Excel 16.0).
			const enumName = declaredExpected?.split('.').pop()?.toLowerCase();
			const expected = enumName && [...(symbols.root.children ?? []), ...(projectVisibleSymbols ?? [])]
				.some((sym) => sym.kind === 'enum' && sym.name.toLowerCase() === enumName) ? 'Long' : declaredExpected;
			// `Sheet1 = 5` compiles as a Let through the document's default
			// member, and a Worksheet or Workbook has none (issue #225).
			if (!expected && !targetType.resolved && isDocumentModuleName(assignment.name, memberCtx)) {
				// Word refuses `ThisDocument = 5` while compiling (issue #228).
				if (memberCtx.model?.hostName === 'Word') {
					push(
						'setRequiresObject',
						`'${assignment.name}' names the document module itself, which no assignment can replace. This is a VBE compile error: Invalid use of property.`,
						assignment.span,
					);
					return;
				}
				push(
					'setRequired',
					`'${assignment.name}' names the document module itself: a Let reaches it through its default member, which a document does not have, and a Set cannot replace it either. This will raise Run-time error '438': Object doesn't support this property or method.`,
					assignment.span,
				);
				return;
			}
			if (!expected) {
				return;
			}
			if (isKnownObjectAssignmentType(expected, memberCtx)) {
				// The VBE compiles a bare `=` to an object variable as a Let
				// through the type's default member (issue #107): `r = 5`
				// writes the Range's Value. What is reported is what the
				// default member makes of it.
				const verdict = objectLetAssignmentVerdict(expected, memberCtx);
				// `x = 5` on a Word Paragraph: its default member Range holds an
				// object, which a Let cannot write (issue #462, measured in Word 16.0).
				// A DAO Recordset's Fields holds an object too (issue #464).
				const holding = verdict !== 'noDefault' ? objectHoldingDefault(expected, memberCtx) : undefined;
				if (holding) {
					push(
						'invalidPropertyUse',
						`Assignment to '${assignment.name}' reaches the default member ${holding.name} of ${expected}, which holds an object (${holding.returns}), so a Let cannot write it. This is a VBE compile error: Invalid use of property.`,
						assignment.span,
					);
					return;
				}
				// A class whose default member is a Property Get with no Let:
				// `c = 5` does not compile (issue #256, measured in Excel 16.0).
				const readOnlyDefault = verdict === 'lets' ? readOnlyProjectDefault(expected, memberCtx) : undefined;
				if (readOnlyDefault) {
					push(
						'readonlyMemberAssignment',
						`Assignment to '${assignment.name}' reaches the default member ${readOnlyDefault} of ${expected}, a Property Get with no Property Let. This is a VBE compile error: Invalid use of property.`,
						assignment.span,
					);
					return;
				}
				if (verdict === 'argument') {
					push(
						'setRequired',
						// A Collection says "Argument not optional"; Excel's
						// collections, Hyperlinks and Workbooks among them, say
						// "Invalid use of property" (issue #221, measured).
						`Assignment to '${assignment.name}' requires Set: the default member of ${expected} takes an argument, so a Let cannot reach it. This is a VBE compile error: ${normalizeType(expected) === 'collection' ? 'Argument not optional' : 'Invalid use of property'}.`,
						assignment.span,
					);
				} else if (verdict === 'noDefault') {
					// While the object is Nothing the Let raises 91, and 438 only
					// once it holds one (issue #193). The object-state walk says
					// which; this rule owns the report either way, since the fix is
					// the Set.
					const state = objectLetStateAt(source, mod, procedure, symbols, memberCtx, activity, assignment.span.start);
					const lower = assignment.name.toLowerCase();
					const declared = procSym?.children?.find((child) => child.name.toLowerCase() === lower)
						?? symbols.root.children?.find((child) => child.name.toLowerCase() === lower);
					const error = state === 'unset'
						? `It is still Nothing here, so this will raise Run-time error '91': Object variable or With block variable not set.`
						: state === 'set' || declared?.isAutoInstantiated
							? `This will raise Run-time error '438': Object doesn't support this property or method.`
							: `This will raise Run-time error '438': Object doesn't support this property or method, or '91' while it is Nothing.`;
					push(
						'setRequired',
						`Assignment to '${assignment.name}' requires Set: ${expected} has no default member for a Let to reach. ${error}`,
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
			const arrayProblem = arrayAssignmentProblem(
				assignment,
				span.start,
				targetShape,
				(name) => {
					const resolved = declaredShapeForSourceBinding(symbols, procSym, projectVisibleSymbols, name, 'expression');
					return resolved.resolved ? resolved.shape : shapes.get(name.toLowerCase());
				},
				(name) => arrayValueAt(stmt, name),
				(tokens) => inferArgumentType(tokens, span.start, env, moduleSignatures, sourceNames, source, memberCtx, resolveExpressionType, resolveQualifiedExpressionType),
				sourceNames,
				(name) => variantArrayFunctions.has(name.toLowerCase())
					&& !procSym?.children?.some((child) => child.name.toLowerCase() === name.toLowerCase()),
			);
			if (arrayProblem) {
				push(arrayProblem.code, arrayProblem.message, arrayProblem.span);
				return;
			}
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
			const nullCall = nullFromChoice(assignment.valueTokens, symbols);
			if (nullCall && isKnownScalarType(normalizeType(expected) ?? '')) {
				push(
					'assignmentTypeMismatch',
					`Assignment to '${assignment.name}' expects ${expected}, but ${nullCall.why}, so it returns Null. Null cannot be coerced to this scalar type. This will raise Run-time error '94': Invalid use of Null.`,
					{ start: span.start + nullCall.first.start, end: span.start + nullCall.last.end },
				);
				return;
			}
			const nullSource = nullHeldAt(stmt, span, assignment.valueTokens);
			if (nullSource && isKnownScalarType(normalizeType(expected) ?? '')) {
				push(
					'assignmentTypeMismatch',
					`Assignment to '${assignment.name}' expects ${expected}, but '${nullSource.name}' ${nullSource.returns ? 'returns Null' : 'holds Null here'}. Null cannot be coerced to this scalar type. This will raise Run-time error '94': Invalid use of Null.`,
					nullSource.span,
				);
				return;
			}
			const nullValue = isKnownScalarType(normalizeType(expected) ?? '') ? nullExpressionAt(stmt, span, assignment.valueTokens) : undefined;
			if (nullValue) {
				push(
					'assignmentTypeMismatch',
					`Assignment to '${assignment.name}' expects ${expected}, but '${nullValue.text}' is Null: an operator on Null gives Null. Null cannot be coerced to this scalar type. This will raise Run-time error '94': Invalid use of Null.`,
					nullValue.span,
				);
				return;
			}
			const knownString = isKnownScalarType(normalizeType(expected) ?? '') ? knownStringAt(stmt, span, assignment.valueTokens, expected) : undefined;
			if (knownString) {
				const reason = incompatibilityReason(expected, knownString);
				if (reason) {
					push(
						'assignmentTypeMismatch',
						`Assignment to '${assignment.name}' expects ${expected}, but ${knownString.label} here. ${reason.replace('This string literal', 'This string')}`,
						knownString.span,
					);
				}
				return;
			}
			if (!actual) {
				return;
			}
			const reason = incompatibilityReason(expected, actual);
			if (!reason) {
				return;
			}
			// A constant out of the target's range overflows; its type is no
			// mismatch: `b = vbTrue` stores -1 in a Byte (issue #326).
			const bounds = actual.numericConstantName !== undefined && actual.numericValue !== undefined ? numericLiteralBounds(normalizeType(expected) ?? '') : undefined;
			push(
				'assignmentTypeMismatch',
				bounds
					? `Assignment to '${assignment.name}' stores ${actual.numericConstantName}, which is ${actual.numericValue}, in ${/^[AEIOU]/.test(bounds.label) ? 'an' : 'a'} ${bounds.label}, whose range is ${bounds.min} to ${bounds.max}. This will raise Run-time error '6': Overflow.`
					: `Assignment to '${assignment.name}' expects ${expected}, but got ${actual.label}. ${reason}`,
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

/**
 * A Choose, Switch or IIf whose literal arguments make it return Null
 * (issue #243, measured in Excel 16.0): `Choose(5, 1, 2)` and
 * `Choose(0, "a")` name no choice, `Switch(False, 1)` finds no True
 * condition, and `IIf(False, 1, Null)` takes its Null. A module procedure of
 * the same name is the module's.
 */
function nullFromChoice(
	valueTokens: readonly VbaToken[],
	symbols: ReturnType<typeof buildModuleSymbols>,
): { why: string; first: VbaToken; last: VbaToken } | undefined {
	const toks = valueTokens.filter((tok) => tok.kind !== 'comment');
	const start = tokenText(toks[0]) === 'vba' && toks[1]?.rawText === '.' ? 2 : 0;
	const fn = tokenText(toks[start]);
	if ((fn !== 'choose' && fn !== 'switch' && fn !== 'iif') || toks[start + 1]?.rawText !== '(' || matchParenFrom(toks, start + 1) !== toks.length - 1) {
		return undefined;
	}
	if (start === 0 && (symbols.root.children ?? []).some((child) => child.name.toLowerCase() === fn)) {
		return undefined;
	}
	const args = splitTopLevelTokenGroups(toks, start + 2, ',', toks.length - 1);
	const literal = (arg: readonly VbaToken[]): number | undefined => {
		const word = arg.length === 1 ? tokenText(arg[0]) : '';
		if (word === 'true' || word === 'false') {
			return word === 'true' ? -1 : 0;
		}
		const signed = arg.length === 2 && arg[0].rawText === '-';
		const number = signed ? arg[1] : arg.length === 1 ? arg[0] : undefined;
		const value = number?.kind === 'integerLiteral' ? Number(number.rawText.replace(/[%&^]$/, '')) : undefined;
		return value === undefined || !Number.isFinite(value) ? undefined : signed ? -value : value;
	};
	const shown = toks.slice(start).map((tok) => tok.rawText).join('').replace(/,/g, ', ');
	const result = { first: toks[0], last: toks[toks.length - 1] };
	if (fn === 'choose') {
		const index = args.length > 1 ? literal(args[0]) : undefined;
		const choices = args.length - 1;
		return index !== undefined && (index < 1 || index > choices)
			? { ...result, why: `${shown} names no choice: its index ${index} is not 1 to ${choices}` }
			: undefined;
	}
	if (fn === 'switch') {
		const conditions = args.filter((_arg, i) => i % 2 === 0);
		return args.length % 2 === 0 && conditions.every((arg) => literal(arg) === 0)
			? { ...result, why: `${shown} has no condition that is True` }
			: undefined;
	}
	const condition = args.length === 3 ? literal(args[0]) : undefined;
	const chosen = condition === undefined ? undefined : args[condition === 0 ? 2 : 1];
	return chosen?.length === 1 && tokenText(chosen[0]) === 'null'
		? { ...result, why: `${shown} takes the branch that is Null` }
		: undefined;
}

/** An array value, by its element type, or an Empty Variant. */
interface ArrayValue {
	element: string;
	text: string;
}

/**
 * What an array target takes, and what an array value goes into (issue #194,
 * each measured in Excel 16.0):
 *
 *  - A fixed array takes no assignment, and a dynamic array takes no scalar:
 *    `a = b` into `Dim a(1)`, `a = "abc"`, `a = 5`, `a = Join(...)` do not
 *    compile ("Can't assign to array"). A Byte array takes a String.
 *  - A dynamic array takes an array variable of its own element type only:
 *    Long() from Integer(), and Variant() from Long(), do not compile.
 *  - Array() gives Variant() and Split and Filter give String(), so either
 *    into another element type raises 13: `Dim a() As String: a = Array("x")`.
 *    So does an Empty Variant, and so does an array into a scalar.
 */
function arrayAssignmentProblem(
	assignment: { name: string; span: Span; valueTokens: VbaToken[] },
	baseOffset: number,
	targetShape: DeclaredValueShape | undefined,
	sourceShape: (name: string) => DeclaredValueShape | undefined,
	variantValue: (name: string) => ArrayValue | undefined,
	scalarType: (tokens: VbaToken[]) => InferredArgumentType | undefined,
	sourceNames: SourceNameScope,
	returnsVariantArray: (name: string) => boolean = () => false,
): { code: 'arrayTargetAssignment' | 'assignmentTypeMismatch'; message: string; span: Span } | undefined {
	const value = assignment.valueTokens.filter((tok) => tok.kind !== 'comment');
	if (value.length === 0) {
		return undefined;
	}
	const valueSpan = { start: baseOffset + value[0].start, end: baseOffset + value[value.length - 1].end };
	const shown = value.length === 1 ? value[0].rawText : 'this value';
	const name = value.length === 1 ? tokenName(value[0]) : undefined;
	const named = name ? sourceShape(name) : undefined;
	// Only into an array: into a scalar, Empty would run.
	const called = targetShape?.isArray && isWholeCall(value) && returnsVariantArray(value[0].rawText)
		? { element: 'variant', text: `${value[0].rawText}(...), which returns Array(...) or Empty,` }
		: undefined;
	const produced = arrayProducedBy(value, sourceNames) ?? called ?? (name && !named?.isArray ? variantValue(name) : undefined);
	const targetType = elementType(targetShape?.asType);
	if (targetShape?.isArray) {
		const elements = `an array of ${(targetShape.asType ?? 'Variant').replace(/\s*\(\s*\)\s*$/, '')}`;
		const cannot = (what: string): { code: 'arrayTargetAssignment'; message: string; span: Span } => ({
			code: 'arrayTargetAssignment',
			message: `Can't assign to array: '${assignment.name}' is ${elements}, and ${what}. This is a VBE compile error.`,
			span: assignment.span,
		});
		if (targetShape.isFixedArray) {
			return cannot('a fixed-size array takes no assignment whole');
		}
		if (named?.isArray) {
			return elementType(named.asType) === targetType ? undefined : cannot(`'${name}' is an array of ${named.asType ?? 'Variant'}`);
		}
		// A Function declared to return a typed array is held to the same
		// rule as an array variable: `a = StrArr()` into Long() or Variant()
		// does not compile (issue #222, measured in Excel 16.0).
		const returned = !produced && isWholeCall(value) ? scalarType(value) : undefined;
		if (returned && /\(\s*\)\s*$/.test(returned.type)) {
			const element = returned.type.replace(/\s*\(\s*\)\s*$/, '');
			return elementType(element) === targetType ? undefined : cannot(`${value[0].rawText}(...) returns an array of ${element}`);
		}
		if (produced) {
			if (produced.element === targetType) {
				return undefined;
			}
			const what = produced.element === 'empty' ? produced.text : `${produced.text} holds ${produced.element === 'string' ? 'String' : 'Variant'} elements`;
			return {
				code: 'assignmentTypeMismatch',
				message: `Assignment to '${assignment.name}' expects ${elements}, but ${what}. This will raise Run-time error '13': Type mismatch.`,
				span: valueSpan,
			};
		}
		// A scalar is a literal, or a VBA function that returns one: an
		// expression's inferred type can miss an array (a UDT field, a
		// Function returning Byte()).
		const literal = value.length === 1 && ['stringLiteral', 'integerLiteral', 'floatLiteral'].includes(value[0].kind);
		const runtimeCall = scalarRuntimeCall(value, sourceNames);
		const scalar = literal || runtimeCall ? scalarType(value) : undefined;
		const scalarKind = normalizeType(scalar?.type);
		if (scalar && scalarKind && !/\(\s*\)\s*$/.test(scalar.type) && isKnownScalarType(scalarKind) && !(targetType === 'byte' && scalarKind === 'string')) {
			return cannot(`${shown} is a ${scalar.type}, not an array`);
		}
		return undefined;
	}
	if (produced && produced.element !== 'empty' && targetShape && isKnownScalarType(targetType)) {
		return {
			code: 'assignmentTypeMismatch',
			message: `Assignment to '${assignment.name}' expects ${targetShape.asType}, but ${produced.text} is an array. This will raise Run-time error '13': Type mismatch.`,
			span: valueSpan,
		};
	}
	return undefined;
}

/**
 * The module's Functions declared As Variant whose every assignment to their
 * own name is `Array(...)`: each returns an array of Variant, or Empty when
 * no assignment runs. Into an array of another element type both raise 13
 * (issue #222, measured in Excel 16.0: `a = VarArr()` into Long()).
 */
function arrayOnlyVariantFunctions(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): Set<string> {
	const out = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || member.procKind !== 'Function' || member.typeSuffix) {
			continue;
		}
		const returns = normalizeType(member.returnType);
		if (returns !== undefined && returns !== 'variant') {
			continue;
		}
		const lower = member.name.toLowerCase();
		let onlyArrays = true;
		forEachStatement(member.body, (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				let toks = statementTokens(source, span);
				// A one-line If's own span runs to its branches, which come next
				// as spans of their own; only its condition is read here.
				const branches = stmt.kind === 'Statement' && stmt.singleLineIfBranches;
				if (branches && span === stmt.span) {
					const then = toks.findIndex((tok) => tokenText(tok) === 'then');
					toks = then >= 0 ? toks.slice(0, then) : toks;
					if (toks.some((tok) => tokenName(tok)?.toLowerCase() === lower)) {
						onlyArrays = false;
					}
					continue;
				}
				if (!toks.some((tok) => tokenName(tok)?.toLowerCase() === lower)) {
					continue;
				}
				const target = bareAssignmentTarget(source, span);
				// A one-line If's Then branch runs to its Else.
				const value = (target?.valueTokens.filter((tok) => tok.kind !== 'comment') ?? [])
					.filter((tok, index, all) => !(index === all.length - 1 && tokenText(tok) === 'else'));
				const isArrayCall = tokenText(value[0]) === 'array' && value[1]?.rawText === '(' && matchParenFrom(value, 1) === value.length - 1;
				const readsSelf = value.some((tok) => tokenName(tok)?.toLowerCase() === lower);
				if (target?.name.toLowerCase() !== lower || !isArrayCall || readsSelf) {
					onlyArrays = false;
				}
			}
		}, activity);
		if (onlyArrays) {
			out.add(lower);
		}
	}
	return out;
}

/** Whether a bare name no declaration resolves is a document module's: Sheet1, ThisWorkbook. */
function isDocumentModuleName(name: string, memberCtx: MemberCompletionContext): boolean {
	const lower = name.toLowerCase();
	return (memberCtx.projectClassMembers ?? []).some((type) => type.kind === 'document' && type.name.toLowerCase() === lower);
}

/** Whether the value is one call and nothing more: `F()`, `F(1, 2)`. */
function isWholeCall(value: readonly VbaToken[]): boolean {
	return tokenName(value[0]) !== undefined && value[1]?.rawText === '(' && matchParenFrom(value, 1) === value.length - 1;
}

/** An array's element type, normalized: "Byte()" and "Byte" are both byte. */
function elementType(asType: string | undefined): string {
	return normalizeType(asType?.replace(/\s*\(\s*\)\s*$/, '')) ?? 'variant';
}

/** Whether the value is one call to a VBA runtime function that returns a scalar: `Join(...)`. */
function scalarRuntimeCall(value: readonly VbaToken[], sourceNames: SourceNameScope): boolean {
	let index = 0;
	if (tokenText(value[0]) === 'vba' && value[1]?.rawText === '.') {
		index = 2;
	}
	const name = tokenName(value[index]);
	let paren = index + 1;
	if (value[paren]?.rawText === '$') {
		paren++;
	}
	if (!name || value[paren]?.rawText !== '(' || matchParenFrom(value, paren) !== value.length - 1) {
		return false;
	}
	if (index === 0 && runtimeCallableSourceShadowed(name, sourceNames)) {
		return false;
	}
	const returns = normalizeType(resolveRuntimeFunction(name)?.returns);
	return returns !== undefined && isKnownScalarType(returns);
}

/** The array a call returns, by element type: Array() is Variant(), Split and Filter are String(). */
function arrayProducedBy(value: readonly VbaToken[], sourceNames: SourceNameScope): ArrayValue | undefined {
	let index = 0;
	if (tokenText(value[0]) === 'vba' && value[1]?.rawText === '.') {
		index = 2;
	}
	const callee = tokenText(value[index]);
	if (value[index + 1]?.rawText !== '(' || matchParenFrom(value, index + 1) !== value.length - 1) {
		return undefined;
	}
	if (index === 0 && runtimeCallableSourceShadowed(value[0].rawText, sourceNames)) {
		return undefined;
	}
	const text = `${value.slice(0, index + 1).map((tok) => tok.rawText).join('')}(...)`;
	if (callee === 'array') {
		return { element: 'variant', text };
	}
	if (callee === 'split' || callee === 'filter') {
		return { element: 'string', text };
	}
	return undefined;
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

/**
 * The compile error the VBE gives an assignment to a read-only host property,
 * or undefined when the assignment compiles or the models cannot say which
 * error it is. Measured in Excel, Word and PowerPoint 16.0 (issue #198):
 *
 * - A property of type Variant or Object takes either statement: the value
 *   goes to whatever the property returns when the code runs.
 * - A Let to a scalar property is "Can't assign to read-only property", except
 *   on Excel's dispatch-only interfaces (Range, Shape, Font, ...), where it is
 *   "Wrong number of arguments or invalid property assignment", or
 *   "Assignment to constant not permitted" when the property takes
 *   parameters (Range.Address).
 * - A Set to an object property is "Invalid use of property". A Let to one
 *   goes to the returned object's default member, so it is not decided here.
 * - A Set to a scalar property gives the Let's error on a dispatch-only
 *   interface and "Invalid use of property" on a dual one, but "Type mismatch"
 *   when a dual property takes parameters (Word's Range.XML), and the Word,
 *   PowerPoint and Office models do not record a property's parameters. So a
 *   Set is judged on Excel's types only, whose parameters the model has.
 */
function hostReadOnlyAssignmentError(
	target: MemberCompletion,
	usesSet: boolean,
	memberCtx: MemberCompletionContext,
): string | undefined {
	if (target.kind !== 'property') {
		return undefined;
	}
	const declared = target.declaredType?.trim() ?? '';
	if (!declared || /^(?:Variant|Object)$/i.test(declared)) {
		return undefined;
	}
	const scalar = isKnownScalarType(normalizeType(declared) ?? '')
		|| resolveHostEnum(declared, memberCtx.model) !== undefined;
	const dispatchOnly = isDispatchOnlyHostType(target.owner, memberCtx.model);
	const withParameters = signatureDeclaresParameters(target.signature);
	if (!scalar) {
		return usesSet && target.returns && !withParameters ? 'Invalid use of property' : undefined;
	}
	if (dispatchOnly) {
		return withParameters
			? 'Assignment to constant not permitted'
			: 'Wrong number of arguments or invalid property assignment';
	}
	if (!usesSet) {
		return "Can't assign to read-only property";
	}
	return target.owner.startsWith('Excel.') ? 'Invalid use of property' : undefined;
}

/** Whether the receiver of the member ending at `offset` binds only at run time. */
function lateBoundReceiver(source: string, offset: number, memberCtx: MemberCompletionContext): boolean {
	const receiver = resolveReceiverTypeAt(source, offset, memberCtx);
	return receiver === undefined || isLateBoundTypeKey(receiver);
}

/**
 * A Set whose value cannot be the target's type. A scalar value is refused
 * when the module compiles: "Type mismatch" into a variable (`Set r = 5`),
 * "Object required" through a Property Set (`Set h.Item = 5`). An object of
 * the wrong class compiles, and raises 13 when the Set runs, so under On
 * Error Resume Next it is handled (issue #202, measured in Excel 16.0).
 */
function pushObjectAssignmentMismatch(
	push: PushFn,
	label: string,
	expected: string | undefined,
	actual: { label: string; span?: Span } | undefined,
	reason: string,
	fallback: Span,
	scalarError: 'Type mismatch' | 'Object required',
): void {
	if (reason === SCALAR_OBJECT_ASSIGNMENT_REASON) {
		push(
			'setRequiresObject',
			`Set assigns an object to '${label}', which expects ${expected}, but ${actual?.label} is not an object. This is a VBE compile error: ${scalarError}.`,
			actual?.span ?? fallback,
		);
		return;
	}
	push(
		'assignmentObjectTypeMismatch',
		`Object assignment to '${label}' expects ${expected}, but got ${actual?.label}. ${reason} This will raise Run-time error '13': Type mismatch.`,
		actual?.span ?? fallback,
	);
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
	const projectClasses = (memberCtx.projectClassMembers?.length ?? 0) > 0;
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
		if (target?.access === 'read-only' && target.writable === undefined) {
			const vbeError = hostReadOnlyAssignmentError(target, assignment.usesSet, memberCtx);
			if (vbeError && !lateBoundReceiver(source, assignment.memberSpan.end, memberCtx)) {
				push(
					'readonlyMemberAssignment',
					`Cannot assign to read-only property '${assignment.label}'. This is a VBE compile error: ${vbeError}.`,
					assignment.memberSpan,
				);
			}
			return;
		}
		// `Range("A1").Font.Size = 500`: a value the host refuses (issue #204).
		if (target && target.writable === undefined && !assignment.usesSet && !assignment.withArguments) {
			const problem = hostPropertyValueProblem(target, assignment.valueTokens);
			const value = assignment.valueTokens.filter((tok) => tok.kind !== 'comment');
			if (problem && value.length > 0) {
				push('hostPropertyValueOutOfRange', problem, { start: span.start + value[0].start, end: span.start + value[value.length - 1].end });
				return;
			}
		}
		// `Set f.T1 = Nothing`: a form's control is no variable to Set (issue
		// #226, measured in Excel 16.0: "Invalid use of property").
		if (assignment.usesSet && !assignment.withArguments && target && target.writable === undefined
			&& /^MSForms\./i.test(target.returns ?? '')
			&& (memberCtx.projectClassMembers ?? []).some((type) => type.kind === 'userform' && type.name === target.owner)) {
			push(
				'setRequiresObject',
				`'${assignment.label}' is a control on the form ${target.owner}, which no Set can replace. This is a VBE compile error: Invalid use of property.`,
				assignment.memberSpan,
			);
			return;
		}
		// The project-class checks read a bare property target only.
		if (!projectClasses || assignment.withArguments || !target || target.writable === undefined) {
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
			// A Property Set takes the Set whatever the Let and Get are typed:
			// `Set c.M = New Collection` runs beside a Long Let (issue #414).
			if (expected && isKnownScalarType(normalizeType(expected) ?? '') && !target.setAccessor) {
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
				pushObjectAssignmentMismatch(push, assignment.label, expected, actual, reason, assignment.memberSpan, 'Object required');
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
	activity?: ConditionalActivityTracker,
): ProcedureStatementVisitor {
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const { resolveExpressionType, resolveQualifiedExpressionType } =
			sourceBindingTypeResolvers(symbols, procSym, projectVisibleSymbols);
		// What an Object or Variant local holds at a statement (issue #246).
		let heldAt: ReturnType<typeof heldObjectsAt> | undefined;
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkSetSpan(span, stmt);
			}
		};

		function checkSetSpan(span: Span, stmt: LeafStatementNode): void {
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
			// `Set Sheet1 = Nothing`: a document's name is no variable to Set
			// (issue #225, measured in Excel 16.0).
			if (!targetDeclaredType.resolved && isDocumentModuleName(target.name, memberCtx)) {
				push(
					'setRequiresObject',
					`'${target.name}' names the document module itself, which no Set can replace. This is a VBE compile error: Invalid use of property.`,
					target.span,
				);
				return;
			}
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
				let shown = actual;
				let reason = objectAssignmentIncompatibilityReason(
					expected,
					actual,
					memberCtx,
				);
				// `Set o = New Flat1` then `Set c = o`: the class an Object holds
				// is checked as the Set runs (issue #246, measured in Excel 16.0).
				const value = target.valueTokens.filter((tok) => tok.kind !== 'comment');
				if (!reason && value.length === 1 && tokenName(value[0])) {
					heldAt ??= heldObjectsAt(source, member, symbols, activity);
					const held = heldAt(stmt).classes.get(tokenName(value[0])!.toLowerCase());
					if (held) {
						shown = { type: held, label: `'${value[0].rawText}', which holds a ${held} here`, span: { start: span.start + value[0].start, end: span.start + value[0].end } };
						reason = objectAssignmentIncompatibilityReason(expected, shown, memberCtx);
					}
				}
				const sheets = reason ? undefined : sheetsFromCollectionProperty(value, expected, sourceNames, memberCtx);
				if (sheets) {
					shown = { type: 'Excel.Sheets', label: `'${sheets.text}', which returns a Sheets object`, span: { start: span.start + value[0].start, end: span.start + value[value.length - 1].end } };
					reason = `Excel's Worksheets and Charts properties return a Sheets object, never a ${sheets.collection} one.`;
				}
				if (reason) {
					pushObjectAssignmentMismatch(push, target.name, expected, shown, reason, target.span, 'Type mismatch');
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
	// A number is no more a target than a string is: `Mid(5, 1) = "x"` is a
	// "Syntax error" as well (issue #213, measured in Excel 16.0).
	const literalKinds = ['stringLiteral', 'integerLiteral', 'floatLiteral'];
	if (!target || target.length !== 1 || !literalKinds.includes(target[0].kind)) {
		return undefined; // target is not exactly one literal
	}
	return {
		span: { start: span.start + target[0].start, end: span.start + target[0].end },
		message:
			"The target of a Mid statement must be a writable String variable, not a " +
			`${target[0].kind === 'stringLiteral' ? 'string literal' : 'number'}. Assigning into a literal is a compile error.`,
	};
}

/**
 * The text a String expression spells out, or undefined. `standIn` marks a
 * Date written as text, a Date literal or local or CStr of one, whose text
 * the locale decides: it is never a number or a Boolean, so the text stands
 * in for it only to say that (issue #405, measured in Excel 16.0).
 */
interface SpelledText {
	text: string;
	standIn: boolean;
	/**
	 * What the text is, where the host or the locale decides the words: a
	 * month name, a type name, a cell address. Such a word is never a
	 * number, a Boolean or a date, and the text stands in for it (issue
	 * #457, measured in Excel 16.0).
	 */
	named?: string;
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const STRCONV_CASES: Readonly<Record<string, number>> = { vbuppercase: 1, vblowercase: 2, vbpropercase: 3 };

/**
 * A call or property whose String result is fixed, or is a word no locale
 * reads as a number, a Boolean or a date (issue #457, each measured in
 * Excel 16.0 into a Long, a Double, a Boolean and a Date): Chr, Hex, Oct,
 * Space, String and StrConv over literals; MonthName, WeekdayName, Format of
 * a Date literal as a month or day name, and TypeName; a Range's Address;
 * Application.Name and PathSeparator. `Hex(9)` is "9", which converts.
 */
function fixedTextPart(part: readonly VbaToken[], known: (lower: string) => KnownLocalValue | undefined, env: ReadonlyMap<string, string>, sourceNames: SourceNameScope): SpelledText | undefined {
	const exact = (text: string): SpelledText => ({ text, standIn: false });
	const named = (text: string, what: string): SpelledText => ({ text, standIn: false, named: what });
	const first = tokenText(part[0]);
	if (part.length === 3 && first === 'application' && part[1].rawText === '.') {
		const member = tokenText(part[2]);
		return member === 'name' ? named('Microsoft Excel', 'the application name') : member === 'pathseparator' ? named('\\', 'the path separator') : undefined;
	}
	// `Range("A1").Address`, `Cells(1, 2).Address(False, False)`. The text
	// stands in with no digit: a string with one may be a number somewhere.
	if ((first === 'range' || first === 'cells') && part[1]?.rawText === '(') {
		const close = matchParenFrom(part, 1);
		const tail = part.slice(close + 1);
		const address = tail.length >= 2 && tail[0].rawText === '.' && tokenText(tail[1]) === 'address'
			&& (tail.length === 2 || (tail[2].rawText === '(' && matchParenFrom(tail, 2) === tail.length - 1));
		return close > 0 && address ? named('address', 'a cell address') : undefined;
	}
	const open = part[1]?.rawText === '$' ? 2 : 1;
	const fn = tokenName(part[0])?.toLowerCase();
	if (!fn || runtimeCallableSourceShadowed(fn, sourceNames) || part[open]?.rawText !== '(' || matchParenFrom(part, open) !== part.length - 1) {
		return undefined;
	}
	const args = splitTopLevelTokenGroups(part, open + 1, ',', part.length - 1);
	const whole = (k: number): number | undefined => {
		const arg = args[k];
		return arg?.length === 1 && arg[0].kind === 'integerLiteral' && /^\d+$/.test(arg[0].rawText) ? Number(arg[0].rawText) : undefined;
	};
	const n = whole(0);
	switch (fn) {
		case 'typename':
			return args.length === 1 ? named('Integer', 'a type name') : undefined;
		case 'monthname':
			return n !== undefined && n >= 1 && n <= 12 && args.length <= 2 ? named(MONTH_NAMES[n - 1], 'a month name') : undefined;
		case 'weekdayname':
			return n !== undefined && n >= 1 && n <= 7 && args.length <= 3 ? named(DAY_NAMES[n - 1], 'a day name') : undefined;
		case 'format': {
			const pattern = args.length === 2 && args[1].length === 1 && args[1][0].kind === 'stringLiteral' ? stringLiteralValue(args[1][0].rawText).toLowerCase() : undefined;
			const date = args[0]?.length === 1 && args[0][0].kind === 'dateLiteral';
			return date && (pattern === 'mmm' || pattern === 'mmmm') ? named('January', 'a month name')
				: date && (pattern === 'ddd' || pattern === 'dddd') ? named('Sunday', 'a day name') : undefined;
		}
		case 'chr':
			return args.length === 1 && n !== undefined && n >= 32 && n <= 126 ? exact(String.fromCharCode(n)) : undefined;
		case 'hex':
		case 'oct':
			return args.length === 1 && n !== undefined && n <= 2147483647 ? exact(n.toString(fn === 'hex' ? 16 : 8).toUpperCase()) : undefined;
		case 'space':
			return args.length === 1 && n !== undefined && n <= 1000 ? exact(' '.repeat(n)) : undefined;
		case 'string': {
			const fill = args[1]?.length === 1 && args[1][0].kind === 'stringLiteral' ? stringLiteralValue(args[1][0].rawText) : '';
			return args.length === 2 && n !== undefined && n <= 1000 && fill !== '' ? exact(fill[0].repeat(n)) : undefined;
		}
		case 'strconv': {
			const subject = args.length === 2 ? spelledText(args[0], known, env, sourceNames) : undefined;
			const kind = args[1]?.length === 1 ? (whole(1) ?? STRCONV_CASES[tokenText(args[1][0])]) : undefined;
			if (!subject || subject.standIn || subject.named !== undefined) {
				return undefined;
			}
			const lower = subject.text.toLowerCase();
			return kind === 1 ? exact(subject.text.toUpperCase()) : kind === 2 ? exact(lower)
				: kind === 3 ? exact(lower.replace(/(^|[^a-z0-9])([a-z])/g, (_, before: string, letter: string) => before + letter.toUpperCase())) : undefined;
		}
	}
	return undefined;
}

/** The text functions folded over known text: `Left("abc", 1)` is "a". */
/** The binary operators whose result is Null when an operand is (issue #324). `&` is here to be refused. */
const NULL_PROPAGATING: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod', '=', '<>', '<', '>', '<=', '>=', 'and', 'or', 'xor', 'eqv', 'imp', '&']);

/** The number a literal operand is, True as -1 and False as 0: `1`, `-2.5`, `True`. */
function literalNumber(operand: readonly VbaToken[]): number | undefined {
	const toks = unwrapOuterParens([...operand]);
	const sign = toks.length === 2 && toks[0].rawText === '-' ? -1 : 1;
	const tok = toks.length === 1 ? toks[0] : toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+') ? toks[1] : undefined;
	const word = tokenText(tok);
	if (word === 'true' || word === 'false') {
		return sign * (word === 'true' ? -1 : 0);
	}
	if (tok?.kind !== 'integerLiteral' && tok?.kind !== 'floatLiteral') {
		return undefined;
	}
	const value = Number(tok.rawText.replace(/[%&^!#@]$/, ''));
	return Number.isFinite(value) ? sign * value : undefined;
}

const TEXT_FUNCTIONS: ReadonlySet<string> = new Set(['cstr', 'left', 'right', 'mid', 'ucase', 'lcase', 'trim', 'ltrim', 'rtrim']);

/**
 * `"a" & "b"`, `Left("abc", 1)`, `o & 5` with o a Boolean known to be True:
 * each part of a `&` chain a literal, a known local, or a text function over
 * those. A whole number is written in digits and a Boolean as True or False.
 */
function spelledText(
	toks: readonly VbaToken[],
	known: (lower: string) => KnownLocalValue | undefined,
	env: ReadonlyMap<string, string>,
	sourceNames: SourceNameScope,
): SpelledText | undefined {
	let text = '';
	let standIn = false;
	const parts = splitTopLevelTokenGroups(toks, 0, '&');
	for (const part of parts) {
		const spelled = spelledPart(unwrapOuterParens(part), known, env, sourceNames);
		if (!spelled) {
			return undefined;
		}
		// A month name or address stands in only for itself: beside other
		// text, as in `"1 " & MonthName(1)`, it may make a date.
		if (spelled.named !== undefined) {
			return parts.length === 1 ? spelled : undefined;
		}
		text += spelled.text;
		standIn ||= spelled.standIn;
	}
	return { text, standIn };
}

function spelledPart(
	part: VbaToken[],
	known: (lower: string) => KnownLocalValue | undefined,
	env: ReadonlyMap<string, string>,
	sourceNames: SourceNameScope,
): SpelledText | undefined {
	const exact = (text: string): SpelledText => ({ text, standIn: false });
	if (part.length === 2 && part[0].rawText === '-' && /^\d+[%&]?$/.test(part[1].rawText)) {
		return exact(`-${Number(part[1].rawText.replace(/[%&]$/, ''))}`);
	}
	if (part.length === 1) {
		const tok = part[0];
		if (tok.kind === 'stringLiteral') {
			return exact(stringLiteralValue(tok.rawText));
		}
		if (tok.kind === 'integerLiteral' && /^\d+[%&]?$/.test(tok.rawText)) {
			return exact(String(Number(tok.rawText.replace(/[%&]$/, ''))));
		}
		if (tok.kind === 'dateLiteral') {
			return { text: tok.rawText, standIn: true };
		}
		const word = tokenText(tok);
		if (word === 'true' || word === 'false') {
			return exact(word === 'true' ? 'True' : 'False');
		}
		const lower = tokenName(tok)?.toLowerCase();
		if (!lower) {
			return undefined;
		}
		const type = normalizeType(env.get(lower));
		if (type === 'date') {
			return { text: `[${tok.rawText}]`, standIn: true };
		}
		const value = known(lower);
		if (value?.kind === 'string' && !value.contentMutated) {
			return exact(value.value as string);
		}
		if (value?.kind === 'number' && type === 'boolean') {
			return exact(value.value === 0 ? 'False' : 'True');
		}
		if (value?.kind === 'number' && (type === 'byte' || type === 'integer' || type === 'long') && Number.isInteger(value.value)) {
			return exact(String(value.value));
		}
		return undefined;
	}
	const fixed = fixedTextPart(part, known, env, sourceNames);
	if (fixed) {
		return fixed;
	}
	// `Left$` lexes as Left and a `$`.
	const open = part[1]?.rawText === '$' ? 2 : 1;
	const fn = tokenName(part[0])?.toLowerCase();
	if (!fn || !TEXT_FUNCTIONS.has(fn) || runtimeCallableSourceShadowed(fn, sourceNames) || part[open]?.rawText !== '(' || matchParenFrom(part, open) !== part.length - 1) {
		return undefined;
	}
	const args = splitTopLevelTokenGroups(part, open + 1, ',', part.length - 1);
	const subject = spelledText(args[0], known, env, sourceNames);
	// Only CStr passes a Date written as text on: Left of it depends on the locale.
	if (!subject || ((subject.standIn || subject.named !== undefined) && fn !== 'cstr')) {
		return undefined;
	}
	const count = (k: number): number | undefined => {
		const arg = args[k];
		return arg?.length === 1 && /^\d+$/.test(arg[0].rawText) ? Number(arg[0].rawText) : undefined;
	};
	const s = subject.text;
	switch (fn) {
		case 'cstr':
			return args.length === 1 ? subject : undefined;
		case 'ucase':
			return args.length === 1 ? exact(s.toUpperCase()) : undefined;
		case 'lcase':
			return args.length === 1 ? exact(s.toLowerCase()) : undefined;
		case 'trim':
			return args.length === 1 ? exact(s.replace(/^ +| +$/g, '')) : undefined;
		case 'ltrim':
			return args.length === 1 ? exact(s.replace(/^ +/, '')) : undefined;
		case 'rtrim':
			return args.length === 1 ? exact(s.replace(/ +$/, '')) : undefined;
		case 'left':
		case 'right': {
			const n = args.length === 2 ? count(1) : undefined;
			return n === undefined ? undefined : exact(fn === 'left' ? s.slice(0, n) : s.slice(Math.max(0, s.length - n)));
		}
		case 'mid': {
			const start = count(1);
			const length = args.length === 3 ? count(2) : args.length === 2 ? s.length : undefined;
			return start === undefined || start < 1 || length === undefined ? undefined : exact(s.substr(start - 1, length));
		}
	}
	return undefined;
}
