import {sourceSetterAssignment, invalidSetterAssignmentArity} from '../setterAssignment';
import { classMemberValues } from '../../symbols/classMemberFacts';
import { memberParameterCounts } from '../memberParameterCounts';
import { isLeafStatement } from '../../parser/nodes';
import { projectSetterValueType, projectGetterKnownValue } from '../../completion/projectSetterValueType';
// Rule family: assignment validation (audit #0).
//
// Extracted verbatim from analyzeModule.ts: Const reassignment, scalar and
// member assignment type compatibility, Set assignment validation, and
// missing Function/Property Get return assignments.

import { createAssignmentCoercionType } from '../assignmentCoercionType';
import { assignmentTargetFromTokens, assignmentTargetName } from '../../completion/assignmentTarget';
import {
	isLateBoundTypeKey,
	resolveReceiverTypeAt,
	signatureDeclaresParameters,
	type MemberCompletion,
	type MemberCompletionContext,
} from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { isDispatchOnlyHostType, resolveHostEnum } from '../../host/hostModel';
import { formulaStringProblem, hostPropertyValueProblem, hostUnionPropertyValueProblem } from './hostPropertyValues';
import {
	matchParenFrom,
	splitTopLevelTokenGroups,
} from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import { MAX_EXPRESSION_DEPTH } from '../../parser/expressionLimits';
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
import {procedureParamsFromSymbol} from '../../symbols/symbolModel';
import type {
	VbaProcedureSignature,
	VbaProcedureParam,
	VbaProjectClassMember,
	VbaProjectClassMembers,
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
	validateArity,
	splitArgSlots,
	type InferredArgumentType,
	extractQualifiedCall,
} from '../callExtraction';
import {
	buildModuleTypeSignatures,
	createObjectAssignmentTypeResolver,
	createObjectDefaultQueries,
	createObjectTypeImplementationLookup,
	createProjectInterfaceSharingLookup,
	callableSignatureForCall,
	callableTypeSignaturesFor,
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	declaredTypeForSourceBinding,
	type DeclaredValueShape,
	incompatibilityReason,
	objectHoldingDefault,
	readOnlyHostDefault,
	getterMayReturnObject,
	objectLetAssignmentVerdict,
	sameByRefType,
	inferArgumentType,
	defTypeOf,
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
import { operatorYieldsNull } from '../nullOperators';

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
function readOnlyProjectDefault(type: string, projectClassNamed: ReturnType<typeof createObjectDefaultQueries>['projectClassNamed']): string | undefined {
	const lower = type.trim().split('.').pop()?.toLowerCase();
	const cls = lower ? projectClassNamed(lower) : undefined;
	const member = cls?.exhaustive === true ? cls.members.find((candidate) => candidate.defaultMember) : undefined;
	return member && member.kind === 'property' && !member.letAccessor && member.writable !== true ? member.name : undefined;
}

function invalidGetterArgumentCount(source: string, name: string, span: Span, params: readonly VbaProcedureParam[], tokens: VbaToken[], base: number): boolean {
	const split = tokens.length ? splitArgSlots(tokens,base) : {slots:[],spans:[]};
	let invalid = false;
	validateArity(source,{name,params:params.map(param => ({...param,optional:Boolean(param.optional),paramArray:Boolean(param.paramArray)}))},
		{name,nameSpan:span,slots:split.slots,slotSpans:split.spans,sliceStart:base},()=>{invalid=true;});
	return invalid;
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
	hasArguments: boolean;
	argumentTokens: VbaToken[];
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
		hasArguments: withArguments && lhs.length > memberIndex + 3,
		argumentTokens: withArguments ? lhs.slice(memberIndex+2,-1) : [],
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
	const isDocumentModule = projectTypeNameLookup(memberCtx, 'document', false);
	const isFormOwner = projectTypeNameLookup(memberCtx, 'userform', true);
	const defaultQueries = createObjectDefaultQueries(memberCtx);
	const resolveObjectType = defaultQueries.resolveType;
	const shareInterfaces = createProjectInterfaceSharingLookup(memberCtx);
	const implementsType = createObjectTypeImplementationLookup();
	const objectAssignmentReason = (expected: string | undefined, actual: ReturnType<typeof inferArgumentType>) =>
		objectAssignmentIncompatibilityReason(expected, actual, memberCtx, resolveObjectType, shareInterfaces, implementsType);
	// Declared-type facts are stable within this rule invocation. Value and
	// object-state facts below still depend on the individual statement.
	const objectTypes = new Map<string, {
		isObject: boolean;
		verdict: ReturnType<typeof objectLetAssignmentVerdict>;
		holding: ReturnType<typeof objectHoldingDefault>;
		readOnlyDefault: string | undefined;
	}>();
	const objectFactsFor = (type: string) => {
		let facts = objectTypes.get(type);
		if (!facts) {
			const isObject = resolveObjectType(type) !== undefined;
			const verdict = isObject ? defaultQueries.verdictFor(type) : 'unknown';
			const holding = isObject && verdict !== 'noDefault' ? objectHoldingDefault(type, memberCtx) : undefined;
			const readOnlyDefault = isObject && !holding
				? readOnlyProjectDefault(type, defaultQueries.projectClassNamed) ?? readOnlyHostDefault(type, memberCtx)
				: undefined;
			facts = { isObject, verdict, holding, readOnlyDefault };
			objectTypes.set(type, facts);
		}
		return facts;
	};
	const checkGetterObjectDefault = (type: string, label: string, span: Span): boolean => {
		const facts = objectFactsFor(type);
		if (facts.holding || facts.readOnlyDefault) {
			const member = facts.holding?.name ?? facts.readOnlyDefault;
			push('invalidPropertyUse', `Assignment through '${label}' reaches the default member ${member} of ${type}, which has no writable Let contract. This is a VBE compile error: Invalid use of property.`, span);
			return true;
		}
		if (facts.verdict === 'argument') {
			push('argumentCount', `Argument not optional: '${label}' returns ${type}, whose default member requires an index before a Let can reach it. This is a VBE compile error.`, span);
			return true;
		}
		if (facts.verdict === 'noDefault') {
			push('runtimeMemberNotFound', `'${label}' returns ${type}, which has no default member to receive this Let assignment. This will raise Run-time error '438': Object doesn't support this property or method, or error '91' if the getter returns Nothing.`, span);
			return true;
		}
		return false;
	};
	// Base depends on this module/activity pass, not on a procedure or value.
	// Resolve it only when array folding needs it; zero is a cached result too.
	let optionBase: number | undefined;
	const optionBaseFor = (): number => optionBase ??= moduleOptionBase(mod, activity);
	// Null-choice shadowing checks only direct module names, not the broader
	// local/project runtime scope. At most three names are queried per pass.
	let choiceModuleNames: Map<string, boolean> | undefined;
	const choiceModuleNameDeclared = (lower: string): boolean => {
		choiceModuleNames ??= new Map();
		if (!choiceModuleNames.has(lower)) {
			choiceModuleNames.set(lower, (symbols.root.children ?? []).some((child) => child.name.toLowerCase() === lower));
		}
		return choiceModuleNames.get(lower)!;
	};
	// The Collection.Count guard considers any project surface named Collection.
	// Resolve this exact predicate only when queried, across the whole rule pass.
	let projectCollection: boolean | undefined;
	const projectDeclaresCollection = (): boolean => projectCollection ??=
		(memberCtx.projectClassMembers ?? []).some((type) => type.name.toLowerCase() === 'collection');
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	// Enum assignment compatibility is a name query, not a full symbol scan
	// per assignment. Keep this index within the current rule pass.
	const enumNames = new Set([...(symbols.root.children ?? []), ...(projectVisibleSymbols ?? [])]
		.filter((symbol) => symbol.kind === 'enum')
		.flatMap(symbol => [symbol.name.toLowerCase(), `${symbol.moduleName}.${symbol.name}`.toLowerCase()]));
	const coercionType = createAssignmentCoercionType(memberCtx, enumNames);
	const memberCoercionType = createAssignmentCoercionType(memberCtx);
	const setterNames = new Set([...(symbols.root.children ?? []), ...(projectVisibleSymbols ?? [])]
		.filter(symbol => symbol.kind === 'propertyLet').map(symbol => symbol.name.toLowerCase()));
	const variantArrayFunctions = arrayOnlyVariantFunctions(source, mod, activity);
	const getterNames = new Set([...(symbols.root.children ?? []), ...(projectVisibleSymbols ?? [])].filter(symbol => symbol.kind === 'propertyGet').map(symbol => symbol.name.toLowerCase()));
	let ownGetterValues: ReturnType<typeof classMemberValues> | undefined;

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procedure = member;
		const env = typeEnvironmentFor(symbols, member);
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		// These value helpers inspect the same direct child declaration. Preserve
		// first-match semantics, but share queried names (and misses) per procedure.
		let localSymbols: Map<string, VbaSymbol> | undefined;
		const localSymbolNamed = (lower: string): VbaSymbol | undefined => {
			if (!localSymbols) {
				localSymbols = new Map();
				for (const child of procSym?.children ?? []) {
					const key = child.name.toLowerCase();
					if (!localSymbols.has(key)) { localSymbols.set(key, child); }
				}
			}
			return localSymbols.get(lower);
		};
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
			const local = localSymbolNamed(lower);
			const type = normalizeType(local?.asType);
			if (local?.kind !== 'localVariable' || local.visibility === 'Static' || local.isArray || (type !== undefined && type !== 'variant')) {
				return undefined;
			}
			const shape = (shapesAt ??= knownArrayShapesAt(source, symbols, procedure, activity, optionBaseFor()))(stmt).get(lower);
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
			const local = localSymbolNamed(lower);
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
			const yieldsNull = (toks: readonly VbaToken[]): boolean => operatorYieldsNull(toks, holdsNull);
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
				const local = lower ? localSymbolNamed(lower) : undefined;
				const length = local?.kind === 'localVariable' && local.visibility !== 'Static' && !local.isArray && /^\d+$/.test(local.fixedLength ?? '') ? Number(local.fixedLength) : undefined;
				if (length !== undefined && length >= 1 && (mentions ??= nameMentions(source, procedure, activity)).get(lower!) === 1) {
					return { type: 'String', label: `'${value[0].rawText}', a String * ${length} never assigned, which holds ${length} Chr(0)`, span: valueSpan, stringValue: '\u0000'.repeat(length) };
				}
				return undefined;
			}
			// `"a" & "b"`, `Left("abc", 1)`, `o & 5` (issue #405). A Date written
			// as text converts back to a Date, so that target is left alone.
			const knownAt = (valuesAt ??= knownLocalLiteralValuesAt(source, procedure, symbols, activity))(stmt);
			const spelled = spelledText(value, (lower) => knownAt.get(lower), env, sourceNames, true);
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
			const shapes = (shapesAt ??= knownArrayShapesAt(source, symbols, procedure, activity, optionBaseFor()))(stmt);
			const element = elementOperandStartingAt(value, 0, shapes, optionBaseFor());
			return element && element.last === value.length - 1 && typeof element.value === 'string'
				? { type: 'String', label: `${label} ${JSON.stringify(element.value)}`, span: valueSpan, stringValue: element.value }
				: undefined;
		}

		// `a(0) = New Collection` with a an array As Collection: a Let into the
		// element reaches the default member Item, which needs an argument, as
		// it does into a variable (issue #306, measured in Excel 16.0).
		function checkElementLet(span: Span): void {
			const element = arrayElementTarget(source, span, symbols, procSym, projectVisibleSymbols);
			if (!element || element.usesSet) {
				return;
			}
			const declared = declaredTypeForSourceBinding(symbols, procSym, projectVisibleSymbols, element.name, 'assignmentTarget');
			const expected = declared.resolved ? declared.asType : undefined;
			const facts = expected ? objectFactsFor(expected) : undefined;
			if (facts?.readOnlyDefault) {
				push('readonlyMemberAssignment', `Assignment to '${element.label}' reaches the default member ${facts.readOnlyDefault} of ${expected}, a Property Get with no Property Let. This is a VBE compile error: Invalid use of property.`, element.span);
				return;
			}
			if (expected && objectFactsFor(expected).verdict === 'argument') {
				push(
					'setRequired',
					`Assignment to '${element.label}' requires Set: the default member of ${expected} takes an argument, so a Let cannot reach it. This is a VBE compile error: ${normalizeType(expected) === 'collection' ? 'Argument not optional' : 'Invalid use of property'}.`,
					element.span,
				);
			}
		}

		function checkBareGetter(span: Span): boolean {
			if (!getterNames.size) { return false; }
			const tokens = statementTokensAfterLeadingLabel(source, span);
			let first = firstExecutableTokenIndex(tokens);
			if (tokenText(tokens[first]) === 'let') { first++; }
			const root = tokenName(tokens[first]);
			if (!root || !getterNames.has(root.toLowerCase())) { return false; }
			const equals = topLevelOperatorIndex(tokens, '=');
			if (equals < 0) { return false; }
			const target = assignmentTargetFromTokens(tokens.slice(0, equals + 1)), named = target && assignmentTargetName(target);
			if (!target || named?.index !== 0) { return false; }
			const binding = sourceIdentifierBinding(symbols, procSym, projectVisibleSymbols, root, 'assignmentTarget');
			if (binding.scope === 'ambiguous' || binding.definitions.some(def => def.kind === 'propertyLet' || def.kind === 'propertySet')) { return false; }
			const getter = binding.definitions.find(def => def.kind === 'propertyGet');
			if (!getter || getter === procSym) { return false; }
			const declared = getter.asType ?? (getter.moduleName.toLowerCase() === symbols.moduleName.toLowerCase() ? defTypeOf(symbols, getter.name) : 'Variant');
			const parameters = procedureParamsFromSymbol(getter);
			if (!named.indexed && parameters.some(param => !param.optional && !param.paramArray)) {
				push('argumentCount', `Argument not optional: '${root}' requires a getter argument.`, {start:span.start+target[0].start,end:span.start+target[0].end});
				return true;
			}
			const resultIndexed = named.indexed && target.length > named.index + 3 && parameters.length === 0;
			const nameSpan = {start:span.start+target[0].start,end:span.start+target[0].end};
			if (!resultIndexed && invalidGetterArgumentCount(source,root,nameSpan,parameters,named.indexed?target.slice(named.index+2,-1):[],span.start)) { return true; }
			if (declared && getterMayReturnObject(declared,memberCtx) && !resultIndexed && checkGetterObjectDefault(declared,root,nameSpan)) { return true; }
			if (normalizeType(declared) !== 'variant') { return false; }

			const value = getter.moduleName.toLowerCase() === symbols.moduleName.toLowerCase()
				? (ownGetterValues ??= classMemberValues(source, symbols.root.children ?? [])).get(getter.name.toLowerCase())
				: projectGetterKnownValue(memberCtx, getter.moduleName, getter.name);
			if (value !== 'scalar' && value !== 'empty') { return false; }
			push('variantValueMisuse', `'${root}' has only a Property Get returning a Variant that holds no object. The Let writes through its returned value, which cannot receive a property assignment. This will raise Run-time error '424': Object required.`, { start: span.start + target[0].start, end: span.start + target[0].end });
			return true;
		}

		function checkBareSetter(span: Span, stmt: LeafStatementNode): boolean {
			if (!setterNames.size) { return false; }
			const tokens = statementTokensAfterLeadingLabel(source, span);
			// Only names known to have a Let need syntax/binding work. If branches
			// are visited separately, so the full If span is not an assignment here.
			let first = firstExecutableTokenIndex(tokens);
			if (tokens[first]?.kind === 'keyword' && tokenText(tokens[first]) === 'if') { return false; }
			if (tokens[first]?.kind === 'keyword' && tokenText(tokens[first]) === 'let') { first++; }
			const root = tokenName(tokens[first]);
			if (!root || !setterNames.has(root.toLowerCase())) { return false; }
			const equals = topLevelOperatorIndex(tokens, '=');
			if (equals < 0) { return false; }
			const target = assignmentTargetFromTokens(tokens.slice(0, equals + 1));
			const named = target && assignmentTargetName(target);
			if (!target || named?.index !== 0) { return false; }
			const name = tokenName(target[0]);
			if (!name || !setterNames.has(name.toLowerCase())) { return false; }
			const binding = sourceIdentifierBinding(symbols, procSym, projectVisibleSymbols, name, 'assignmentTarget');
			if (binding.scope === 'ambiguous') { return false; }
			const setter = binding.definitions.find(definition => definition.kind === 'propertyLet');
			const parameter = setter?.children?.filter(child => child.kind === 'parameter').at(-1);
			const declared = parameter?.asType ?? (parameter?.moduleName.toLowerCase() === symbols.moduleName.toLowerCase() ? defTypeOf(symbols, parameter.name) : undefined)
				?? (setter && projectSetterValueType(memberCtx, setter.moduleName, setter.name));
			if (parameter?.isArray) {
				const value = tokens.slice(equals + 1);
				const actual = inferArgumentType(value, span.start, env, moduleSignatures, sourceNames,
					source, memberCtx, resolveExpressionType, resolveQualifiedExpressionType);
				checkArraySetterValue(name, declared, value, span.start, actual, coercionType, resolveExpressionType, resolveQualifiedExpressionType, push, sourceNames, projectDeclaresCollection, defaultQueries.verdictFor, (name) => arrayValueAt(stmt, name));
				return true;
			}
			if (!declared) { return false; }
			const expected = coercionType(declared);
			if (!isKnownScalarType(normalizeType(expected) ?? '')) { return false; }
			const actual = inferArgumentType(tokens.slice(equals + 1), span.start, env, moduleSignatures, sourceNames,
				source, memberCtx, resolveExpressionType, resolveQualifiedExpressionType);
			const problem = arrayAssignmentProblem({ name, span: actual?.span ?? span, valueTokens: tokens.slice(equals + 1) }, span.start,
				{ asType: expected, isArray: false, isFixedArray: false }, () => undefined,
				name => arrayValueAt(stmt, name), () => actual, sourceNames);
			if (problem) { push(problem.code, problem.message, problem.span); return true; }
			const reason = actual && incompatibilityReason(expected, actual);
			if (reason) {
				push('assignmentTypeMismatch', `Assignment to '${name}' expects ${declared}, but got ${actual!.label}. ${reason}`, actual!.span);
			}
			return true;
		}

		function checkAssignmentSpan(span: Span, stmt: LeafStatementNode): void {
			const setter = sourceSetterAssignment(source, span, symbols, procSym, projectVisibleSymbols, memberCtx);
			if (setter && invalidSetterAssignmentArity(setter, source, push)) { return; }
			if (checkBareSetter(span, stmt) || checkBareGetter(span)) { return; }
			const assignment = bareAssignmentTarget(source, span);
			if (!assignment) {
				checkElementLet(span);
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
				? targetType.asType ?? (untypedArray ? undefined : defTypeOf(symbols, assignment.name))
				: env.get(assignment.name.toLowerCase())) ?? (untypedArray ? 'Variant' : undefined);
			// A variable As an Enum is a Long: `x = "abc"` raises 13 and
			// `x = 3000000000#` 6 (issue #436, measured in Excel 16.0).
			const expected = declaredExpected ? coercionType(declaredExpected) : undefined;
			// `Sheet1 = 5` compiles as a Let through the document's default
			// member, and a Worksheet or Workbook has none (issue #225).
			if (!expected && !targetType.resolved && isDocumentModule(assignment.name)) {
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
			const objectType = objectFactsFor(expected);
			if (objectType.isObject) {
				// The VBE compiles a bare `=` to an object variable as a Let
				// through the type's default member (issue #107): `r = 5`
				// writes the Range's Value. What is reported is what the
				// default member makes of it.
				const verdict = objectType.verdict;
				// `x = 5` on a Word Paragraph: its default member Range holds an
				// object, which a Let cannot write (issue #462, measured in Word 16.0).
				// A DAO Recordset's Fields holds an object too (issue #464).
				const holding = objectType.holding;
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
				// So does a Word Document, whose Name takes none (issue #438).
				const readOnlyDefault = objectType.readOnlyDefault;
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
					const state = objectLetStateAt(source, mod, procedure, symbols, memberCtx, activity, assignment.span.start, defaultQueries);
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
			const nullCall = nullFromChoice(assignment.valueTokens, choiceModuleNameDeclared);
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
			memberCoercionType,
			activity,
			push,
			projectDeclaresCollection,
			isFormOwner,
			objectAssignmentReason,
			resolveObjectType,
			defaultQueries.verdictFor,
			checkGetterObjectDefault,
			arrayValueAt,
			resolveExpressionType,
			resolveQualifiedExpressionType,
			symbols,
		);
	}
}

/** A Let's final array parameter is subject to compile-time shape rules. */
function checkArraySetterValue(
	label: string, expected: string | undefined, tokens: readonly VbaToken[], baseOffset: number,
	actual: ReturnType<typeof inferArgumentType>, coercionType: (type: string) => string,
	resolveType: SourceDeclaredTypeResolver | undefined,
	resolveQualifiedType: SourceQualifiedDeclaredTypeResolver | undefined, push: PushFn,
	sourceNames: SourceNameScope, projectDeclaresCollection: () => boolean,
	objectVerdict: (type: string | undefined) => ReturnType<typeof objectLetAssignmentVerdict>,
	variantValue: (name: string) => ArrayValue | undefined,
): void {
	const raw = tokens.filter(token => token.kind !== 'comment' && token.kind !== 'newline');
	if (!raw.length) { return; }
	let value = raw;
	if (raw[0].rawText === '(') {
		// Collect pairs once: peeling N nested groups with N scans is quadratic.
		const opens: number[] = [], pairs = new Map<number, number>();
		for (let i = 0; i < raw.length; i++) {
			if (raw[i].rawText === '(') { opens.push(i); }
			else if (raw[i].rawText === ')') { const open = opens.pop(); if (open !== undefined) { pairs.set(open, i); } }
		}
		let first = 0, last = raw.length - 1;
		while (pairs.get(first) === last) { first++; last--; }
		value = raw.slice(first, last + 1);
	}
	const expectedType = normalizeType(expected), actualType = normalizeType(actual?.type);
	const typedArray = actual && /\(\s*\)\s*$/.test(actual.type);
	const wrongElement = typedArray && expectedType && actualType && isKnownScalarType(expectedType)
		&& isKnownScalarType(actualType) && !sameByRefType(actualType, expectedType);
	const indexed = value[1]?.rawText === '(' && matchParenFrom(value, 1) === value.length - 1;
	const name = value.length === 1 || indexed ? tokenName(value[0]) : undefined;
	const qualified = value.length === 3 && value[1].rawText === '.' && tokenName(value[0]) && tokenName(value[2]);
	const declared = name ? resolveType?.(name) : qualified ? resolveQualifiedType?.(value[0].rawText, value[2].rawText) : undefined;
	const span = { start: baseOffset + raw[0].start, end: baseOffset + raw[raw.length - 1].end };
	const arrayVariable = declared?.isArray && ['localVariable', 'moduleVariable', 'parameter'].includes(declared.kind ?? '');
	if (arrayVariable && (!indexed || value.length === 3)) {
		push('arrayTargetAssignment', `Assignment to '${label}' passes a whole array to a Property Let value parameter. This is a VBE compile error: Can't assign to array.`, span);
		return;
	}
	// Byte-array assignments convert a String value, but not a whole String array.
	const stringElement = arrayVariable && indexed && value.length > 3 && normalizeType(declared?.asType) === 'string';
	if (expectedType === 'byte' && ((actualType === 'string' && !typedArray) || stringElement)) { return; }
	if (value.length === 2 && tokenText(value[0]) === 'new' && normalizeType(value[1].rawText) === 'collection' && !projectDeclaresCollection()) {
		push('argumentCount', `Assignment to '${label}' reads the Collection's default member Item, which requires an index. This is a VBE compile error: Argument not optional.`, span);
		return;
	}
	const variantVariable = declared?.resolved && !declared.isArray
		&& ['localVariable', 'moduleVariable', 'parameter'].includes(declared.kind ?? '')
		&& (normalizeType(declared.asType) ?? 'variant') === 'variant';
	if (expectedType === 'byte' && (variantVariable || arrayProducedBy(value, sourceNames))) {
		const problem = arrayAssignmentProblem({ name: label, span, valueTokens: value }, baseOffset,
			{ asType: 'Byte', isArray: true, isFixedArray: false },
			name => { const binding = resolveType?.(name); return binding?.resolved ? { asType: binding.asType, isArray: !!binding.isArray, isFixedArray: false } : undefined; },
			variantValue, () => actual, sourceNames);
		if (problem) { push(problem.code, problem.message, problem.span); }
		return;
	}
	if (wrongElement || (arrayVariable && indexed && value.length > 3) || variantVariable
		|| (actual && !/\(\s*\)\s*$/.test(actual.type) && isKnownScalarType(normalizeType(coercionType(actual.type)) ?? ''))
		|| arrayProducedBy(value, sourceNames) || (actual && !typedArray && objectVerdict(actual.type) === 'noDefault')) {
		push('argumentShapeMismatch', `Assignment to '${label}' passes ${actual?.label ?? 'an array element'}, but the Property Let value parameter requires an array. This is a VBE compile error: Type mismatch: array or user-defined type expected.`, span);
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
	moduleNameDeclared: (lower: string) => boolean,
): { why: string; first: VbaToken; last: VbaToken } | undefined {
	const toks = valueTokens.filter((tok) => tok.kind !== 'comment');
	const start = tokenText(toks[0]) === 'vba' && toks[1]?.rawText === '.' ? 2 : 0;
	const fn = tokenText(toks[start]);
	if ((fn !== 'choose' && fn !== 'switch' && fn !== 'iif') || toks[start + 1]?.rawText !== '(' || matchParenFrom(toks, start + 1) !== toks.length - 1) {
		return undefined;
	}
	if (start === 0 && moduleNameDeclared(fn)) {
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
			if (!onlyArrays) {
				return;
			}
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
						return;
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
					return;
				}
			}
		}, activity);
		if (onlyArrays) {
			out.add(lower);
		}
	}
	return out;
}

/** Query one project kind without repeating scans, preserving early matches. */
function projectTypeNameLookup(
	memberCtx: MemberCompletionContext,
	kind: VbaProjectClassMembers['kind'],
	caseSensitive: boolean,
): (name: string) => boolean {
	let names: Set<string> | undefined;
	let nextIndex = 0;
	return (name) => {
		names ??= new Set();
		const query = caseSensitive ? name : name.toLowerCase();
		if (names.has(query)) { return true; }
		const surfaces = memberCtx.projectClassMembers ?? [];
		// Metadata is stable within this rule pass. Resume after the last examined
		// surface; a successful query need not inspect the remaining project.
		while (nextIndex < surfaces.length) {
			const type = surfaces[nextIndex++];
			if (type.kind !== kind) { continue; }
			const declared = caseSensitive ? type.name : type.name.toLowerCase();
			names.add(declared);
			if (declared === query) { return true; }
		}
		return false;
	};
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
		if (raises) {
			return;
		}

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

// These getter-only Variant properties return scalar values, never an object
// whose default property could receive a Let. Their assignments compile but
// cannot execute. Keep object-valued getters (e.g. Worksheet.UsedRange) alone.
// https://learn.microsoft.com/en-us/office/vba/api/excel.range.height
const RANGE_READONLY_VALUES = new Set(['height', 'width', 'left', 'top', 'text', 'countlarge', 'hasarray', 'hasformula']);

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
	coercionType: (declared: string) => string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	projectDeclaresCollection: () => boolean,
	isFormOwner: (name: string) => boolean,
	objectAssignmentReason: (expected: string | undefined, actual: ReturnType<typeof inferArgumentType>) => string | undefined,
	resolveObjectType: ReturnType<typeof createObjectAssignmentTypeResolver>,
	objectVerdict: (type: string | undefined) => ReturnType<typeof objectLetAssignmentVerdict>,
	checkGetterObjectDefault: (type: string, label: string, span: Span) => boolean,
	arrayValueAt: (stmt: LeafStatementNode, name: string) => ArrayValue | undefined,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
	symbols?: ReturnType<typeof buildModuleSymbols>,
): void {
	const projectClasses = (memberCtx.projectClassMembers?.length ?? 0) > 0;
	let valuesAt: ReturnType<typeof knownLocalLiteralValuesAt> | undefined;
	let booleanNames: Set<string> | undefined;
	const checkStatement = (span: Span, stmt: BodyNode): void => {
		// A local known to hold a number, for a host property's limits (issue
		// #346). A Boolean holding True is no -1 there: Excel takes True where
		// it refuses -1, and refuses False as it does 0 (issue #630, measured
		// in Excel 16.0).
		const known = (tokens: readonly VbaToken[]): number | undefined => {
			const value = tokens.filter((tok) => tok.kind !== 'comment');
			const lower = value.length === 1 ? tokenName(value[0])?.toLowerCase() : undefined;
			const held = lower && symbols ? (valuesAt ??= knownLocalLiteralValuesAt(source, member, symbols, activity))(stmt).get(lower) : undefined;
			if (held?.kind !== 'number') {
				return undefined;
			}
			// Any direct child with this exact declared type counts, including a
			// later duplicate. Held values still come from the current statement.
			if (!booleanNames) {
				booleanNames = new Set();
				for (const child of procedureSymbolFor(symbols!, member)?.children ?? []) {
					if (child.asType?.toLowerCase() === 'boolean') {
						booleanNames.add(child.name.toLowerCase());
					}
				}
			}
			const boolean = booleanNames.has(lower!);
			return boolean && held.value !== 0 ? undefined : held.value as number;
		};
		const assignment = memberAssignmentTarget(source, span);
		if (symbols) {
			const setter = sourceSetterAssignment(source, span, symbols, procedureSymbolFor(symbols, member), undefined, memberCtx);
			if (setter && invalidSetterAssignmentArity(setter, source, () => {})) { return; }
		}
		if (!assignment) {
			return;
		}
		// `c.Count = 2` on a local Collection: Count is a Long Function (issue
		// #305, measured in Excel 16.0).
		const collectionCount = /^([A-Za-z]\w*)\.count$/i.exec(assignment.label);
		if (collectionCount && !assignment.withArguments && normalizeType(env.get(collectionCount[1].toLowerCase())) === 'collection'
			&& !projectDeclaresCollection()) {
			push(
				'readonlyMemberAssignment',
				`Cannot assign to '${assignment.label}': a Collection's Count is a Function returning Long. This is a VBE compile error: Function call on left-hand side of assignment must return Variant or Object.`,
				assignment.memberSpan,
			);
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
			} else if (!vbeError && target.kind === 'property' && target.owner === 'Excel.Range'
				&& RANGE_READONLY_VALUES.has(target.name.toLowerCase())) {
				const alternative = /^(height|width)$/i.test(target.name)
					? ` Use ${target.name.toLowerCase() === 'height' ? 'RowHeight' : 'ColumnWidth'} to change it.` : '';
				push('hostReadonlyValueAssignment',
					`Cannot assign to read-only property '${assignment.label}': it returns a value and has no setter. This assignment fails when it runs.${alternative}`,
					assignment.memberSpan);
			}
			return;
		}
		// `Range("A1").Formula = "=SUM(B1"`: a formula Excel cannot parse
		// (issue #276). A warning: a cell formatted as Text takes it.
		const formula = target && !assignment.usesSet && !assignment.withArguments ? formulaStringProblem(target, assignment.valueTokens) : undefined;
		if (formula) {
			const value = assignment.valueTokens.filter((tok) => tok.kind !== 'comment');
			push('formulaStringUnparsed', formula, { start: span.start + value[0].start, end: span.start + value[value.length - 1].end });
			return;
		}
		// `Range("A1").Font.Size = 500`: a value the host refuses (issue #204).
		if (target && target.writable === undefined && !assignment.usesSet && !assignment.withArguments) {
			const problem = hostPropertyValueProblem(target, assignment.valueTokens, known);
			const value = assignment.valueTokens.filter((tok) => tok.kind !== 'comment');
			if (problem && value.length > 0) {
				push('hostPropertyValueOutOfRange', problem, { start: span.start + value[0].start, end: span.start + value[value.length - 1].end });
				return;
			}
		}
		// Host scalar setters use the same provable VBA coercions as variables.
		// Variant, Object and object-valued getters need host-specific knowledge.
		if (target?.kind === 'property' && target.access === 'read/write' && target.writable === undefined
			&& !assignment.usesSet && !assignment.withArguments && target.declaredType
			&& (isKnownScalarType(normalizeType(target.declaredType) ?? '') || resolveHostEnum(target.declaredType, memberCtx.model))
			&& !lateBoundReceiver(source, assignment.memberSpan.end, memberCtx)) {
			const actual = inferArgumentType(assignment.valueTokens, span.start, env, moduleSignatures,
				sourceNames, source, memberCtx, resolveExpressionType, resolveQualifiedExpressionType);
			const expected = resolveHostEnum(target.declaredType, memberCtx.model) ? 'Long' : target.declaredType;
			const arrayProblem = arrayAssignmentProblem({ name: assignment.label, span: assignment.memberSpan, valueTokens: assignment.valueTokens }, span.start,
				{ asType: expected, isArray: false, isFixedArray: false }, () => undefined,
				name => isLeafStatement(stmt) ? arrayValueAt(stmt, name) : undefined, () => actual, sourceNames);
			if (arrayProblem) { push(arrayProblem.code, arrayProblem.message, arrayProblem.span); return; }
			const reason = actual && incompatibilityReason(target.declaredType, actual);
			if (reason) {
				push('assignmentTypeMismatch', `Assignment to '${assignment.label}' expects ${target.declaredType}, but got ${actual!.label}. ${reason}`, actual!.span);
				return;
			}
		}
		// `ActiveSheet.Visible = "abc"`: a receiver of several host types,
		// each refusing the value alike (issue #416).
		if ((!target || !target.owner.includes('.')) && target?.writable === undefined && !assignment.usesSet && !assignment.withArguments) {
			const resolved = resolveReceiverTypeAt(source, assignment.memberSpan.start, memberCtx);
			const parts = resolved?.startsWith('union:') ? resolved.slice('union:'.length).split('|') : [];
			const problem = parts.length > 0 ? hostUnionPropertyValueProblem(parts, assignment.member, assignment.valueTokens) : undefined;
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
			&& isFormOwner(target.owner)) {
			push(
				'setRequiresObject',
				`'${assignment.label}' is a control on the form ${target.owner}, which no Set can replace. This is a VBE compile error: Invalid use of property.`,
				assignment.memberSpan,
			);
			return;
		}
		// Source accessors accept their value after the index arguments as well.
		// An indexed field or array-valued property belongs to typeMembers.ts
		// and propertyUse.ts, rather than a scalar setter-value check.
		const indexedAccessor = target && (assignment.usesSet ? target.setAccessor : (target.letAccessor
			|| (target.writable === false && target.signature !== undefined
				&& (signatureDeclaresParameters(target.signature) || normalizeType(target.returns ?? target.declaredType) !== 'string'))));
		if (!projectClasses || (assignment.withArguments && !indexedAccessor) || !target || target.writable === undefined || (target.isArray && !target.writeIsArray)) {
			return;
		}
		if (target.writable === false) {
			// A Let can write through the object returned by Get; it does not replace the property.
			if (!assignment.usesSet && getterMayReturnObject(target.returns ?? target.declaredType, memberCtx)) {
				const parameters = memberParameterCounts(target.signature);
				if (assignment.withArguments && parameters.total === 0 && assignment.hasArguments) { return; } // Result indexing has its own value diagnostic.
				if (!assignment.withArguments && parameters.required > 0) { return; } // Argument-count owns this target.

				const returned = target.returns ?? target.declaredType;
				const getterParams = target.procedureParams?.propertyGet;
				if (getterParams && invalidGetterArgumentCount(source,target.name,assignment.memberSpan,getterParams,assignment.argumentTokens,span.start)) { return; }
				if (returned && checkGetterObjectDefault(returned,assignment.label,assignment.memberSpan)) { return; }
				const type = normalizeType(returned);
				if ((!type || type === 'variant') && (target.knownValue === 'scalar' || target.knownValue === 'empty')) {
					push('variantValueMisuse', `'${assignment.label}' has only a Property Get returning a Variant that holds no object. The Let writes through its returned value, which cannot receive a property assignment. This will raise Run-time error '424': Object required.`, assignment.memberSpan);
				}
				return;
			}
			push(
				'readonlyMemberAssignment',
				`Cannot assign to read-only property '${assignment.label}'.`,
				assignment.memberSpan,
			);
			return;
		}
		if (!assignment.usesSet && target.writeIsArray) {
			const actual = inferArgumentType(assignment.valueTokens, span.start, env, moduleSignatures, sourceNames,
				source, memberCtx, resolveExpressionType, resolveQualifiedExpressionType);
			checkArraySetterValue(assignment.label, target.writeType, assignment.valueTokens, span.start, actual, coercionType,
				resolveExpressionType, resolveQualifiedExpressionType, push, sourceNames, projectDeclaresCollection, objectVerdict, (name) => isLeafStatement(stmt) ? arrayValueAt(stmt, name) : undefined);
			return; // The value is an array parameter, never a scalar value to coerce.
		}
		const declaredExpected = target.writeType ?? target.returns;
		const expected = declaredExpected ? coercionType(declaredExpected) : undefined;
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
			const reason = objectAssignmentReason(expected, actual);
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
		if (!target.letAccessor && isKnownObjectAssignmentType(expected, memberCtx, resolveObjectType)) {
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
		const arrayProblem = arrayAssignmentProblem({ name: assignment.label, span: assignment.memberSpan, valueTokens: assignment.valueTokens }, span.start,
			{ asType: expected, isArray: false, isFixedArray: false }, () => undefined,
			name => isLeafStatement(stmt) ? arrayValueAt(stmt, name) : undefined,
			tokens => inferArgumentType(tokens, span.start, env, moduleSignatures, sourceNames, source, memberCtx, resolveExpressionType, resolveQualifiedExpressionType), sourceNames);
		if (arrayProblem) { push(arrayProblem.code, arrayProblem.message, arrayProblem.span); return; }
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
			`Assignment to '${assignment.label}' expects ${declaredExpected}, but got ${actual.label}. ${reason}`,
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
			checkStatement(span, stmt);
		}
	}, activity);
}

/**
 * `v(0) = x` or `Set v(0) = x` where v is a declared array: the element as
 * an assignment target, with the array's name, the span from the name to
 * the closing parenthesis, and the value.
 */
function arrayElementTarget(
	source: string,
	span: Span,
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
): { name: string; label: string; span: Span; valueTokens: VbaToken[]; usesSet: boolean } | undefined {
	const toks = statementTokens(source, span);
	let i = firstExecutableTokenIndex(toks);
	const head = tokenText(toks[i]);
	if (head === 'set' || head === 'let') {
		i++;
	}
	const name = tokenName(toks[i]);
	if (!name || toks[i + 1]?.rawText !== '(') {
		return undefined;
	}
	const close = matchParenFrom(toks, i + 1);
	if (close < 0 || toks[close + 1]?.rawText !== '=' || close + 2 >= toks.length) {
		return undefined;
	}
	const shape = declaredShapeForSourceBinding(symbols, procSym, projectVisibleSymbols, name, 'assignmentTarget');
	if (!shape.resolved || shape.shape?.isArray !== true) {
		return undefined;
	}
	return {
		name,
		label: toks.slice(i, close + 1).map((tok) => tok.rawText).join(''),
		span: { start: span.start + toks[i].start, end: span.start + toks[close].end },
		valueTokens: toks.slice(close + 2),
		usesSet: head === 'set',
	};
}

export function checkSetAssignments(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
	activity?: ConditionalActivityTracker,
): ProcedureStatementVisitor {
	const isDocumentModule = projectTypeNameLookup(memberCtx, 'document', false);
	const isProjectClass = projectTypeNameLookup(memberCtx, 'class', false);
	const resolveObjectType = createObjectAssignmentTypeResolver(memberCtx);
	const shareInterfaces = createProjectInterfaceSharingLookup(memberCtx);
	const implementsType = createObjectTypeImplementationLookup();
	// Form metadata is stable within this rule invocation; query only the names
	// actually used, retaining the first matching control and missing results.
	let formResolved = false;
	let currentForm: VbaProjectClassMembers | undefined;
	let controls: Map<string, VbaProjectClassMember | undefined> | undefined;
	function formControl(lower: string): VbaProjectClassMember | undefined {
		if (!memberCtx.meProjectType) { return undefined; }
		if (!formResolved) {
			const formName = memberCtx.meProjectType.toLowerCase();
			currentForm = memberCtx.projectClassMembers?.find((type) => type.kind === 'userform' && type.name.toLowerCase() === formName);
			formResolved = true;
		}
		if (!currentForm) { return undefined; }
		controls ??= new Map();
		if (!controls.has(lower)) {
			controls.set(lower, currentForm.members.find((member) => member.name.toLowerCase() === lower && /^MSForms\./i.test(member.returns ?? '')));
		}
		return controls.get(lower);
	}
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	let moduleDeclaredNames: ReadonlySet<string> | undefined;
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		let procedureDeclaredNames: ReadonlySet<string> | undefined;
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
			// `Set v(0) = x` into an element of a declared array is judged as
			// a Set into a variable of its element type (issue #306).
			const element = arrayElementTarget(source, span, symbols, procSym, projectVisibleSymbols);
			const target = setAssignmentTarget(source, span) ?? (element?.usesSet ? element : undefined);
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
			if (!targetDeclaredType.resolved && isDocumentModule(target.name)) {
				push(
					'setRequiresObject',
					`'${target.name}' names the document module itself, which no Set can replace. This is a VBE compile error: Invalid use of property.`,
					target.span,
				);
				return;
			}
			// `Set Answer = Nothing` inside the form: a control is no
			// variable to Set (issue #315, measured in Excel 16.0).
			const lowerTarget = target.name.toLowerCase();
			// Only a form context needs this exact direct-declaration shadow check.
			// Keep it local to this invocation: project names and enum members
			// in the broader runtime shadow scope do not count as declarations here.
			const declaredHere = !!memberCtx.meProjectType && (
				(procedureDeclaredNames ??= new Set((procSym?.children ?? []).map((symbol) => symbol.name.toLowerCase()))).has(lowerTarget)
				|| (moduleDeclaredNames ??= new Set((symbols.root.children ?? []).map((symbol) => symbol.name.toLowerCase()))).has(lowerTarget)
			);
			const control = declaredHere ? undefined : formControl(lowerTarget);
			if (control) {
				push(
					'setRequiresObject',
					`'${target.name}' is a control on this form, which no Set can replace. This is a VBE compile error: Invalid use of property.`,
					target.span,
				);
				return;
			}
			const expected = targetDeclaredType.resolved
				? targetDeclaredType.asType
				: env.get(target.name.toLowerCase());
			const targetType = normalizeType(expected);
			// `Set t = Prompt` with t As MSForms.TextBox and Prompt a Label on
			// this form (issue #315, measured in Excel 16.0: 13).
			const controlClass = /^(?:msforms\.)?(textbox|label|listbox|combobox|checkbox|optionbutton|togglebutton|commandbutton|frame|multipage|tabstrip|scrollbar|spinbutton|image)$/i.exec(expected?.trim() ?? '')?.[1]?.toLowerCase();
			// Both target extractors slice significant statement tokens.
			const value = target.valueTokens;
			const valueControl = controlClass && value.length === 1 && tokenName(value[0]) && !env.has(tokenName(value[0])!.toLowerCase())
				? formControl(tokenName(value[0])!.toLowerCase())
				: undefined;
			const valueClass = valueControl?.returns?.slice('MSForms.'.length).toLowerCase();
			if (valueControl && valueClass && valueClass !== controlClass) {
				push(
					'assignmentObjectTypeMismatch',
					`Object assignment to '${target.name}' expects ${expected}, but '${value[0].rawText}' is a control of class ${valueControl.returns}. This will raise Run-time error '13': Type mismatch.`,
					{ start: span.start + value[0].start, end: span.start + value[0].end },
				);
				return;
			}
			// `Set v = 5` is refused whatever v is: a literal is never an object
			// reference ("Object required", issue #125, measured in Excel 16.0).
			if ((!targetType || targetType === 'variant') && value.length === 1 && isScalarLiteralToken(value[0])) {
				push(
					'setRequiresObject',
					`Set assigns an object reference, but ${value[0].rawText} is a literal value. This is a VBE compile error: Object required.`,
					{ start: span.start + value[0].start, end: span.start + value[0].end },
				);
				return;
			}
			if (!targetType || !isKnownScalarType(targetType)) {
				if (!resolveObjectType(expected)) {
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
					resolveObjectType,
					shareInterfaces,
					implementsType,
				);
				// `Set o = New Flat1` then `Set c = o`: the class an Object holds
				// is checked as the Set runs (issue #246, measured in Excel 16.0).
				if (!reason && value.length === 1 && tokenName(value[0])) {
					heldAt ??= heldObjectsAt(source, member, symbols, activity);
					const held = heldAt(stmt).classes.get(tokenName(value[0])!.toLowerCase());
					if (held) {
						shown = { type: held, label: `'${value[0].rawText}', which holds a ${held} here`, span: { start: span.start + value[0].start, end: span.start + value[0].end } };
						reason = objectAssignmentIncompatibilityReason(expected, shown, memberCtx, resolveObjectType, shareInterfaces, implementsType);
					}
				}
				// `Set c = ActiveSheet`: a Worksheet or a Chart, never a
				// Collection or a class of the project (issue #306, measured in
				// Excel 16.0: 13).
				if (!reason && value.length === 1 && tokenText(value[0]) === 'activesheet' && !sourceNames.runtimeShadows.has('activesheet')
					&& (targetType === 'collection' || (targetType !== undefined && isProjectClass(targetType)))) {
					shown = { type: 'Object', label: 'ActiveSheet, a Worksheet or a Chart', span: { start: span.start + value[0].start, end: span.start + value[0].end } };
					reason = `ActiveSheet holds a sheet, never a ${expected}.`;
				}
				const sheets = reason ? undefined : sheetsFromCollectionProperty(value, expected, sourceNames, memberCtx);
				if (sheets) {
					shown = { type: 'Excel.Sheets', label: `'${sheets.text}', which returns a Sheets object`, span: { start: span.start + value[0].start, end: span.start + value[value.length - 1].end } };
					reason = `Excel's Worksheets and Charts properties return a Sheets object, never a ${sheets.collection} one.`;
				}
				if (reason) {
					pushObjectAssignmentMismatch(push, target === element ? element.label : target.name, expected, shown, reason, target.span, 'Type mismatch');
				}
				return;
			}
			if (target === element) {
				return; // an element of a scalar array, which the VBE has not been asked about
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
	const argToks = toks.slice(parenIndex + 1, close);
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

/** `Range("c1").Address` as Excel gives it with no arguments: "$C$1", or "$A$1:$B$2". */
function literalRangeAddress(part: readonly VbaToken[]): string | undefined {
	const toks = part.filter((tok) => tok.kind !== 'comment');
	if (toks.length !== 6 || tokenText(toks[0]) !== 'range' || toks[1].rawText !== '(' || toks[2].kind !== 'stringLiteral' || toks[3].rawText !== ')'
		|| toks[4].rawText !== '.' || tokenText(toks[5]) !== 'address') {
		return undefined;
	}
	const cells = stringLiteralValue(toks[2].rawText).toUpperCase().split(':');
	if (cells.length > 2 || !cells.every((cell) => /^[A-Z]{1,3}[1-9]\d*$/.test(cell))) {
		return undefined;
	}
	return cells.map((cell) => cell.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2')).join(':');
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
function fixedTextPart(part: readonly VbaToken[], known: (lower: string) => KnownLocalValue | undefined, env: ReadonlyMap<string, string>, sourceNames: SourceNameScope, nesting: number): SpelledText | undefined {
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
	// `Split(Range("C1").Address, "$")(1)`, the column letter "C" (issue #457).
	if (fn === 'split' && part[open]?.rawText === '(' && !runtimeCallableSourceShadowed(fn, sourceNames)) {
		const close = matchParenFrom(part, open);
		const index = part.slice(close + 1);
		const [text, separator, ...rest] = splitTopLevelTokenGroups(part, open + 1, ',', close);
		const address = text ? literalRangeAddress(text) : undefined;
		const at = index.length === 3 && index[0].rawText === '(' && index[1].kind === 'integerLiteral' && index[2].rawText === ')' ? Number(index[1].rawText) : undefined;
		if (address && separator?.length === 1 && separator[0].kind === 'stringLiteral' && rest.length === 0 && at !== undefined) {
			const sep = stringLiteralValue(separator[0].rawText);
			const element = sep ? address.split(sep)[at] : undefined;
			return element === undefined ? undefined : exact(element);
		}
		return undefined;
	}
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
			const subject = args.length === 2 ? spelledText(args[0], known, env, sourceNames, false, nesting + 1) : undefined;
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
	whole = false,
	nesting = 0,
): SpelledText | undefined {
	// Recover conservatively on unfinished or deeply nested editor input, as
	// the expression parser and other recursive folders do. Siblings share
	// the current depth; only descending into a call argument consumes it.
	if (nesting >= MAX_EXPRESSION_DEPTH) {
		return undefined;
	}
	let text = '';
	let standIn = false;
	const parts = splitTopLevelTokenGroups(toks, 0, '&');
	// A number on its own is no text: only `&` makes one of it. `d = -657435`
	// is a Long, which Overflow judges (issue #329). Inside `CStr(40000)` it
	// is the text the function makes (issue #624).
	const lone = whole && parts.length === 1 ? unwrapOuterParens(parts[0]) : [];
	if (lone.length > 0 && lone.length <= 2 && lone[lone.length - 1].kind === 'integerLiteral' && (lone.length === 1 || lone[0].rawText === '-')) {
		return undefined;
	}
	for (const part of parts) {
		const spelled = spelledPart(unwrapOuterParens(part), known, env, sourceNames, nesting);
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
	nesting: number,
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
	const fixed = fixedTextPart(part, known, env, sourceNames, nesting);
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
	const subject = spelledText(args[0], known, env, sourceNames, false, nesting + 1);
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
