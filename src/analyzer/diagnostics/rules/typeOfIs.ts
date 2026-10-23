// `TypeOf <expr> Is <Type>` always-False detection (MS-VBAL 5.6 TypeOf...Is).
//
// `TypeOf x Is Y` tests whether x's runtime object is-a Y. When x is declared as
// a known, concrete object type A that can never hold a Y, the test is provably
// always False - a dead branch and almost always a bug.
//
// This is the first type rule to consume the §5.6 expression AST directly: it
// reads the parsed `TypeOfIsExpr` nodes that Slices 1-3 produce (notably in
// `If TypeOf x Is Y Then` branch conditions, now parsed expressions) instead of
// re-tokenizing. It reuses the existing object-compatibility tables
// (`objectAssignmentIncompatibilityReason`, which already encodes host/project
// identity and `Implements` interface relationships) so no new type matrix is
// introduced.
//
// No-false-positive discipline:
//   - Only IdentifierExpr operands with a known declared OBJECT type are checked.
//   - `Object`/`Variant`/unknown/scalar operands or targets stay quiet.
//   - The operand type A must be CONCRETE: a host type, or a project class that
//     no project class `Implements`. An interface-typed operand could hold a
//     subtype that is-a Y, so it stays quiet (avoids the diamond false positive).
//   - Fires only when A and Y are mutually assignment-incompatible (neither is-a
//     the other), proven via the existing object tables in both directions.

import type {
	BodyNode,
	ExprNode,
	IdentifierExpr,
	ModuleNode,
	Span,
	TypeOfIsExpr,
} from '../../parser/nodes';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { qualifiedProcedureKey, type VbaProcedureSignature, type VbaSymbol } from '../../symbols/symbolModel';
import type { CallableTypeSignature } from '../callExtraction';
import {
	bareCallableSourceShadowed,
	callableTypeSignaturesFor,
	createObjectAssignmentTypeResolver,
	createObjectTypeImplementationLookup,
	createProjectInterfaceSharingLookup,
	isKnownScalarType,
	objectAssignmentIncompatibilityReason,
	sourceIdentifierBound,
	sourceNameScopeFor,
	type SourceNameScope,
	typeEnvironmentFor,
} from '../typeInference';
import {
	ARGUMENT_TYPED_FUNCTIONS,
	BARE_VALUE_RESULT_FUNCTIONS,
	TYPEOF_KEYWORD_OPERANDS,
	VALUE_RESULT_FUNCTIONS,
} from '../../runtime/vbaValueResults';
import { tokenizeCached } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import { parseExpression } from '../../parser/parseExpression';
import { forEachSubExpression } from '../exprWalk';
import { activeModuleMembers, blockFooterLineSpan, blockHeaderLineSpan, isInactiveNode, statementTokens, tokenText } from '../walker';
import type { ProcedureExpressionVisitor } from '../exprWalk';

/**
 * Rule: `TypeOf` requires an object expression before `Is` (`TypeOf x Is Y`). A
 * malformed `TypeOf Is Y` (no operand) is a compile error - VBE "Syntax error"
 * (oracle-verified `corpus_excel_syntax_006_compile`). Detected by a whole-source
 * scan for an adjacent `TypeOf` `Is` pair (an operand always sits between them in
 * valid code); inactive conditional-compilation regions are skipped (no-FP).
 */
export function checkTypeOfMissingOperand(
	source: string,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const toks = tokenizeCached(source).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	for (let i = 0; i + 1 < toks.length; i++) {
		const word = (toks[i].canonicalText ?? toks[i].rawText).toLowerCase();
		if (word !== 'typeof') {
			continue;
		}
		const nextWord = (toks[i + 1].canonicalText ?? toks[i + 1].rawText).toLowerCase();
		if (nextWord !== 'is') {
			continue;
		}
		const span = { start: toks[i].start, end: toks[i + 1].end };
		if (activity?.isInactive(span)) {
			continue;
		}
		push("typeofMissingOperand", "'TypeOf' requires an object expression before 'Is'.", span);
	}
}

interface TypeOfQueries {
	resolveType: ReturnType<typeof createObjectAssignmentTypeResolver>;
	shareInterfaces: ReturnType<typeof createProjectInterfaceSharingLookup>;
	implementsType: ReturnType<typeof createObjectTypeImplementationLookup>;
	isImplemented: (key: string, display: string) => boolean;
}

export function checkTypeOfIsCompatibility(
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureExpressionVisitor {
	// Project metadata is stable for this analysis query. Keep all derived
	// lookups lazy and query-owned so the next query observes fresh metadata.
	let implementedNames: Set<string> | undefined;
	const queries: TypeOfQueries = {
		resolveType: createObjectAssignmentTypeResolver(memberCtx),
		shareInterfaces: createProjectInterfaceSharingLookup(memberCtx),
		implementsType: createObjectTypeImplementationLookup(),
		isImplemented: (key, display) => {
			if (!implementedNames) {
				implementedNames = new Set();
				// The original exclusion considered every supplied surface kind.
				for (const type of memberCtx.projectClassMembers ?? []) {
					for (const name of type.implements ?? []) { implementedNames.add(name.toLowerCase()); }
				}
			}
			return implementedNames.has(key) || implementedNames.has(display.toLowerCase());
		},
	};
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		return (expr) => {
			if (expr.exprKind === 'TypeOfIsExpr') {
				checkTypeOfIs(expr, env, memberCtx, queries, push);
			}
		};
	};
}

function checkTypeOfIs(
	expr: TypeOfIsExpr,
	env: ReadonlyMap<string, string>,
	memberCtx: MemberCompletionContext,
	queries: TypeOfQueries,
	push: PushFn,
): void {
	if (expr.operand.exprKind !== 'IdentifierExpr') {
		return; // v1: only simple identifier operands have a known declared type
	}
	const operandName = (expr.operand as IdentifierExpr).name;
	const declared = env.get(operandName.toLowerCase());
	if (!declared) {
		return; // undeclared / unknown type -> quiet
	}
	const operandType = queries.resolveType(declared);
	const targetType = queries.resolveType(expr.typeName);
	if (!operandType || !targetType) {
		return; // not both known object types -> quiet
	}
	if (operandType.kind === 'generic' || targetType.kind === 'generic') {
		return; // Object operand or `Is Object` (always True) -> quiet
	}
	if (operandType.key === targetType.key) {
		return; // same type (always True) -> quiet
	}
	// Concrete-operand gate: an interface-typed operand could hold a subtype that
	// is-a the target, so only fire when A cannot be an interface implemented by
	// some class (host types are never user-implementable).
	if (operandType.kind === 'project' && queries.isImplemented(operandType.key, operandType.display)) {
		return;
	}
	// Mutual incompatibility: neither type is assignable to the other (reuses the
	// existing object tables, which already honour `Implements`). If either is
	// assignable, `TypeOf` could be True, so stay quiet.
	const operandCanBeTarget = objectAssignmentIncompatibilityReason(
		expr.typeName,
		{ type: declared, label: declared, span: expr.span },
		memberCtx,
		queries.resolveType,
		queries.shareInterfaces,
		queries.implementsType,
	) === undefined;
	const targetCanBeOperand = objectAssignmentIncompatibilityReason(
		declared,
		{ type: expr.typeName, label: expr.typeName, span: expr.span },
		memberCtx,
		queries.resolveType,
		queries.shareInterfaces,
		queries.implementsType,
	) === undefined;
	if (operandCanBeTarget || targetCanBeOperand) {
		return;
	}
	push(
		'typeOfIsAlwaysFalse',
		`'TypeOf ... Is ${targetType.display}' is always False: '${operandName}' is declared As ${operandType.display}, which is never ${targetType.display}.`,
		expr.span,
	);
}

/**
 * Rule: the binary `Is` object-identity operator requires object-reference
 * operands on both sides (MS-VBAL 5.6). When an operand is PROVABLY a non-object
 * - a scalar value literal, or an identifier declared As a known scalar type
 * (numeric / String / Boolean / Date) - `a Is b` is a VBE compile error
 * ("Object required"). Oracle-verified rejected at compile:
 * `is_scalar_long_var_compile`, `is_scalar_string_var_compile`,
 * `is_two_scalar_vars_compile`, `is_literal_integer_operands_probe`,
 * `is_literal_string_operands_probe`; with the Variant control
 * `is_variant_var_ok_compile` and the Object control `is_object_var_vs_nothing_probe`
 * accepted.
 *
 * No-false-positive discipline: fires only when an operand is PROVABLY scalar.
 * Variant (can hold an object reference), Object/class-typed operands,
 * `Nothing`/`Null`/`Empty` literals, member access, call results, arrays, and any
 * unresolved / undeclared identifier all stay quiet. `TypeOf x Is Y` is a
 * distinct `TypeOfIsExpr`, not a binary `Is`, so it is unaffected.
 */
export function checkIsOperatorOperands(
	symbols: ReturnType<typeof buildModuleSymbols>,
	push: PushFn,
	projectProcedures?: ReadonlyMap<string, readonly VbaProcedureSignature[]>,
	projectVisibleSymbols?: readonly VbaSymbol[],
	objectNames: ReadonlySet<string> = new Set(),
): ProcedureExpressionVisitor {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	let enums: ReadonlySet<string> | undefined;
	const enumNames = (): ReadonlySet<string> => enums ??= enumTypeNames(symbols, projectVisibleSymbols);
	return (member) => {
		const env = withoutControlLocals(typeEnvironmentFor(symbols, member), symbols.implicitLocals?.get(member.span.start), objectNames);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		// A VBA function's name answers the source, or the module's own
		// object, before VBA: a control or a predeclared class named Timer
		// makes `Timer` that object (issue #1419, measured in all four hosts).
		const shadowed = (name: string): boolean => env.has(name.toLowerCase())
			|| objectNames.has(name.toLowerCase())
			|| sourceIdentifierBound(symbols, procSym, projectVisibleSymbols, name, 'expression');
		const resultOf = (form: OperandForm) => (expr: ExprNode): CallResult | undefined =>
			callResultType(expr, env, moduleSignatures, sourceNames) ?? runtimeValueResult(expr, form, env, shadowed);
		const calls = resultOf('is');
		const typeOfCalls = resultOf('typeof');
		const operand = (side: ExprNode, results: (expr: ExprNode) => CallResult | undefined) =>
			nonObjectOperand(side, env, results) ?? enumOperand(side, env, enumNames());
		return (expr) => {
			// `TypeOf n Is Collection` with n a Long is the same compile error,
			// Type mismatch (issue #1411, measured in Excel, Word, PowerPoint
			// and Access 16.0). A literal operand is a Syntax error, which
			// expressions.ts reports (issue #1416).
			if (expr.exprKind === 'TypeOfIsExpr') {
				const offender = expr.operand.exprKind === 'LiteralExpr' ? undefined : operand(expr.operand, typeOfCalls);
				if (offender) {
					push(
						'isOperatorNonObject',
						`'TypeOf ... Is' requires an object operand, but ${offender.detail}, which is not an object.`,
						offender.span,
					);
				}
				return;
			}
			if (expr.exprKind !== 'BinaryExpr' || expr.operator !== 'Is') {
				return;
			}
			const offender = operand(expr.left, calls) ?? operand(expr.right, calls);
			if (offender) {
				push(
					'isOperatorNonObject',
					`The 'Is' operator requires object operands, but ${offender.detail}, which is not an object.`,
					offender.span,
				);
			}
		};
	};
}

/** A call's declared result, and the name to show for the call. */
interface CallResult {
	type: string;
	display: string;
}

/**
 * The declared result of a call to a project or same-module Function,
 * `Bar()` or `Other.Foo()`, typed by the DefType of the module that declares
 * it. Undefined for a variable's element, a name a source variable takes,
 * and VBA's own functions, which were not measured (issue #1411).
 */
function callResultType(
	expr: ExprNode,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope,
): CallResult | undefined {
	if (expr.exprKind !== 'IndexExpr') {
		return undefined;
	}
	const callee = expr.callee;
	if (callee.exprKind === 'IdentifierExpr') {
		const lower = callee.name.toLowerCase();
		if (env.has(lower) || bareCallableSourceShadowed(callee.name, sourceNames)) {
			return undefined;
		}
		// VBA's own functions are left out: the runtime table types `Right`,
		// `Mid`, `InStr`, `Split` and others as values where the VBE treats
		// their result as a Variant (issue #1416's measurements), so only a
		// source declaration says what a call returns here.
		const type = moduleSignatures.get(lower)?.returnType;
		return type ? { type, display: callee.name } : undefined;
	}
	if (callee.exprKind === 'MemberAccessExpr' && callee.accessKind !== 'bang' && callee.object?.exprKind === 'IdentifierExpr') {
		const qualifier = callee.object.name;
		const type = env.has(qualifier.toLowerCase()) || sourceNames.callableShadows.has(qualifier.toLowerCase())
			? undefined
			: moduleSignatures.get(qualifiedProcedureKey(qualifier, callee.member))?.returnType;
		return type ? { type, display: `${qualifier}.${callee.member}` } : undefined;
	}
	return undefined;
}

/**
 * The Enums this module and the project make visible, lowercase. A name
 * another kind of symbol of the project also takes is left out.
 */
function enumTypeNames(
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
): ReadonlySet<string> {
	const enums = new Set<string>();
	const others = new Set<string>();
	for (const sym of [...(symbols.root.children ?? []), ...(projectVisibleSymbols ?? [])]) {
		(sym.kind === 'enum' ? enums : others).add(sym.name.toLowerCase());
	}
	for (const name of others) {
		enums.delete(name);
	}
	return enums;
}

/**
 * `n Is Nothing` and `TypeOf n Is Collection` with n declared As an Enum,
 * of this module or another: Type mismatch (issue #1416, measured in
 * Excel, Word, PowerPoint and Access 16.0).
 */
function enumOperand(
	operand: ExprNode,
	env: ReadonlyMap<string, string>,
	enums: ReadonlySet<string>,
): { span: Span; detail: string } | undefined {
	let expr = operand;
	while (expr.exprKind === 'ParenExpr') {
		expr = expr.inner;
	}
	if (expr.exprKind !== 'IdentifierExpr') {
		return undefined;
	}
	const declared = env.get(expr.name.toLowerCase());
	return declared && enums.has(declared.trim().toLowerCase())
		? { span: expr.span, detail: `'${expr.name}' is declared As ${declared}, an Enum` }
		: undefined;
}

/**
 * The environment without the implicit locals a control's name takes. With
 * DefLng and no Option Explicit, `Picker = 1` in a form whose ComboBox is
 * Picker assigns the control's Value; it declares no Long (issue #1419).
 */
function withoutControlLocals(
	env: ReadonlyMap<string, string>,
	implicitLocals: ReadonlySet<string> | undefined,
	objectNames: ReadonlySet<string>,
): ReadonlyMap<string, string> {
	const controls = [...implicitLocals ?? []].filter((name) => objectNames.has(name));
	if (controls.length === 0) {
		return env;
	}
	const out = new Map(env);
	for (const name of controls) {
		out.delete(name);
	}
	return out;
}

/** Whether a call stands under `Is` or under `TypeOf`, which judge some calls differently. */
type OperandForm = 'is' | 'typeof';

/**
 * A call to one of VBA's own functions whose result is a value under `Is`
 * or `TypeOf` (vbaValueResults.ts, issue #1416): `CDbl(1)`, `VBA.Mid$(s, 1)`,
 * `Erl`. Undefined when `shadowed` says a source or object name takes the
 * function's name or `VBA`'s, and for a bare keyword-named TypeOf operand,
 * which is a Syntax error reported in expressions.ts.
 */
function runtimeValueResult(
	expr: ExprNode,
	form: OperandForm,
	env: ReadonlyMap<string, string>,
	shadowed: (name: string) => boolean,
): CallResult | undefined {
	const target = expr.exprKind === 'IndexExpr' ? expr.callee : expr;
	const args = expr.exprKind === 'IndexExpr' ? expr.args : [];
	let name: string;
	let display: string;
	let qualified: boolean;
	if (target.exprKind === 'IdentifierExpr') {
		if (shadowed(target.name)) {
			return undefined;
		}
		name = target.name;
		display = target.name;
		qualified = false;
		if (target.typeSuffix && target.typeSuffix !== '$') {
			return undefined;
		}
		name += target.typeSuffix ?? '';
	} else if (
		target.exprKind === 'MemberAccessExpr' && target.accessKind !== 'bang'
		&& target.object?.exprKind === 'IdentifierExpr' && target.object.name.toLowerCase() === 'vba'
		&& !shadowed('VBA')
	) {
		if (target.typeSuffix && target.typeSuffix !== '$') {
			return undefined;
		}
		name = `${target.member}${target.typeSuffix ?? ''}`;
		display = `VBA.${name}`;
		qualified = true;
	} else {
		return undefined;
	}
	const lower = name.toLowerCase();
	if (!qualified && form === 'typeof' && TYPEOF_KEYWORD_OPERANDS.has(lower)) {
		return undefined;
	}
	const value = VALUE_RESULT_FUNCTIONS.has(lower)
		|| (!qualified && BARE_VALUE_RESULT_FUNCTIONS.has(lower))
		|| (!qualified && form === 'is' && ARGUMENT_TYPED_FUNCTIONS.has(lower) && args.length > 0
			&& args.every((arg) => arg.value !== undefined && arg.value !== null && provablyScalar(arg.value, env, shadowed)));
	return value ? { type: 'value', display: `${display}${args.length > 0 || expr.exprKind === 'IndexExpr' ? '(...)' : ''}` } : undefined;
}

/** Arithmetic and `&`, which give a value when both sides are values. */
const SCALAR_OPERATORS: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'Mod', '&']);

/** An argument that is provably a value: a literal, a typed variable or element, a value call, or arithmetic on those. */
function provablyScalar(expr: ExprNode, env: ReadonlyMap<string, string>, shadowed: (name: string) => boolean): boolean {
	if (expr.exprKind === 'ParenExpr') {
		return provablyScalar(expr.inner, env, shadowed);
	}
	if (expr.exprKind === 'UnaryExpr' && expr.operator === '-') {
		return provablyScalar(expr.operand, env, shadowed);
	}
	if (expr.exprKind === 'BinaryExpr') {
		return SCALAR_OPERATORS.has(expr.operator) && provablyScalar(expr.left, env, shadowed) && provablyScalar(expr.right, env, shadowed);
	}
	return nonObjectOperand(expr, env, (inner) => runtimeValueResult(inner, 'is', env, shadowed)) !== undefined;
}

/**
 * The same check on the conditions the expression walk does not parse: a
 * single-line If's, and a Do or While loop's (issue #325, measured in Excel
 * 16.0: each is a compile error, Type mismatch).
 */
export function checkIsOperandsInConditions(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	projectProcedures?: ReadonlyMap<string, readonly VbaProcedureSignature[]>,
	projectVisibleSymbols?: readonly VbaSymbol[],
	objectNames?: ReadonlySet<string>,
): void {
	const checkOperands = checkIsOperatorOperands(symbols, push, projectProcedures, projectVisibleSymbols, objectNames);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const checkExpression = checkOperands(member);
		const check = (span: Span, from: number, to: (toks: readonly VbaToken[]) => number): void => {
			// Absolute offsets, so the parsed nodes carry the source's spans.
			const toks = statementTokens(source, span)
				.map((tok) => ({ ...tok, start: span.start + tok.start, end: span.start + tok.end }));
			const end = to(toks);
			if (end <= from) {
				return;
			}
			const parsed = parseExpression(toks, from, end).expr;
			if (!parsed) {
				return;
			}
			forEachSubExpression(parsed, checkExpression);
		};
		const visit = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if (isInactiveNode(activity, node)) {
					continue;
				}
				if (node.kind === 'Statement' && node.singleLineIfBranches) {
					check(node.span, 1, (toks) => (tokenText(toks[0]) === 'if' ? toks.findIndex((tok) => tokenText(tok) === 'then') : -1));
				} else if (node.kind === 'DoBlock' || node.kind === 'WhileBlock') {
					const header = blockHeaderLineSpan(source, node.span);
					const headToks = statementTokens(source, header);
					const from = tokenText(headToks[0]) === 'do' && ['while', 'until'].includes(tokenText(headToks[1])) ? 2 : tokenText(headToks[0]) === 'while' ? 1 : -1;
					if (from > 0) {
						check(header, from, (toks) => toks.length);
					}
					if (node.kind === 'DoBlock') {
						const footer = blockFooterLineSpan(source, node.span);
						const footToks = statementTokens(source, footer);
						if (tokenText(footToks[0]) === 'loop' && ['while', 'until'].includes(tokenText(footToks[1]))) {
							check(footer, 2, (toks) => toks.length);
						}
					}
				}
				if ('body' in node && Array.isArray(node.body)) {
					visit(node.body as BodyNode[]);
				}
			}
		};
		visit(member.body);
	}
}

/** Describes a provably non-object (scalar) operand of `Is`, or undefined. */
function nonObjectOperand(
	operand: ExprNode,
	env: ReadonlyMap<string, string>,
	callResult?: (expr: ExprNode) => CallResult | undefined,
): { span: Span; detail: string } | undefined {
	// `(n) Is Nothing` is the same Type mismatch (issue #1415, measured in
	// Excel, Word, PowerPoint and Access 16.0).
	let expr = operand;
	while (expr.exprKind === 'ParenExpr') {
		expr = expr.inner;
	}
	// `a(0) Is Nothing` for a Long array: its element is a Long (issue #1415).
	if (expr.exprKind === 'IndexExpr' && expr.callee.exprKind === 'IdentifierExpr') {
		const declared = env.get(expr.callee.name.toLowerCase());
		const element = declared?.replace(/\s*\(\s*\)\s*$/, '').trim();
		if (element && isKnownScalarType(element.toLowerCase())) {
			return { span: expr.span, detail: `an element of '${expr.callee.name}' is ${/^[aeiou]/i.test(element) ? 'an' : 'a'} ${element}` };
		}
	}
	const result = callResult?.(expr);
	if (result?.type === 'value') {
		return { span: expr.span, detail: `VBA's '${result.display}' returns a value` };
	}
	if (result && isKnownScalarType(result.type.replace(/\s*\(\s*\)\s*$/, '').trim().toLowerCase())) {
		return { span: expr.span, detail: `'${result.display}' returns ${result.type}` };
	}
	if (expr.exprKind === 'LiteralExpr') {
		const k = expr.literalKind;
		if (k === 'integer' || k === 'float' || k === 'string' || k === 'date' || k === 'boolean') {
			return { span: expr.span, detail: `'${expr.raw}' is a ${k} literal` };
		}
		return undefined; // Nothing / Null / Empty -> not provably scalar
	}
	if (expr.exprKind === 'IdentifierExpr') {
		const declared = env.get(expr.name.toLowerCase());
		if (!declared) {
			return undefined; // undeclared / unknown -> quiet
		}
		// Match declared names literally: vbLong/vbString can name source classes.
		// Strip only a trailing array marker when recognizing built-in scalars.
		const raw = declared.replace(/\s*\(\s*\)\s*$/, '').trim().toLowerCase();
		if (isKnownScalarType(raw)) {
			return { span: expr.span, detail: `'${expr.name}' is declared As ${declared}` };
		}
		return undefined; // Variant / Object / class -> quiet
	}
	return undefined; // member / call / paren / array / New / unary -> quiet (v1)
}
