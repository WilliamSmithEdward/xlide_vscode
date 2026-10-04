// Rule family: expression-syntax rules (audit #0).
//
// Extracted verbatim from analyzeModule.ts: unbalanced parentheses, invalid
// operator sequences, division by a provably-zero divisor, and the
// parenthesized/parenless call-shape rules.

import {
	explicitCallStatementArgumentWithoutParens,
	explicitCallStatementTarget,
	standaloneEmptyParenthesizedCallStatement,
	standaloneMultiArgParenthesizedCallStatement,
} from '../../call/callContext';
import type { HostObjectModel } from '../../host/excelObjectModel';
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	evaluateIntegerConstantExpression,
	type IntegerConstantLookup,
	resolveRawIntegerConstants,
} from '../../constants/integerConstantExpression';
import { tokenizeCached } from '../../lexer/tokenize';
import { relationalOperatorAt } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	BodyNode,
	IfBranchNode,
	LeafStatementNode,
	ModuleNode,
	Span,
} from '../../parser/nodes';
import { isLeafStatement } from '../../parser/nodes';
import {
	resolveRuntimeFunction,
	runtimeAllowsExplicitCall,
} from '../../runtime/vbaRuntime';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import {
	qualifiedProcedureKey,
	type VbaProcedureSignature,
	type VbaSymbol,
} from '../../symbols/symbolModel';
import {
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import {
	callableAcceptsZeroArguments,
	type CallableTypeSignature,
} from '../callExtraction';
import {
	collectModuleLiteralIntegerConstants,
	foldIntegerExpressionTokens,
} from '../constExpr';
import {
	defaultedStraightLine,
	bareCallableSourceShadowed,
	callableSignatureFor,
	callableTypeSignaturesFor,
	declaredTypeForSourceBinding,
	isKnownScalarType,
	isNumericType,
	isProvablyNonNumericString,
	stringConstantsInScope,
	knownLocalLiteralValuesAt,
	unreachableStatementsIn,
	type KnownLocalValue,
	nonnumericStringArithmeticOperand,
	normalizeType,
	procedureIntegerConstantLookup,
	resolveExactMemberCompletion,
	stringLiteralValue,
	runtimeCallableSourceShadowed,
	type SourceDeclaredTypeResolver,
	type SourceNameScope,
	sourceNameScopeFor,
	typeEnvironmentFor,
} from '../typeInference';
import { isInvalidBooleanString, isInvalidDateString } from '../stringConversion';
import { isBareOrVbaQualifiedIntrinsicCall } from './shared';
import {
	elementOperandEndingAt,
	elementsWrittenIn,
	elementOperandStartingAt,
	knownArrayShapesAt,
	moduleOptionBase,
	type ElementOperand,
	type FixedArrayBound,
} from './arrays';
import { moduleTypes } from '../typeFields';
import { checkEachCounterPass, loopCountersAt } from '../loopCounters';
import { isKnownNumber, typeMemberStatesAt, type MemberState } from '../typeMemberState';
import {
	absoluteSpan,
	bareAssignmentTarget,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	firstExecutableTokenIndex,
	matchParenFrom,
	rawExpressionTokens,
	statementTokens,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
	type ProcedureStatementVisitor,
} from '../walker';
import { functionIntegerResult, knownFunctionResults } from '../functionResults';
import { EMPTY_COLLECTION, straightLineAssignments } from '../straightLineValues';
import { conditionValue } from '../conditionValue';

/**
 * Rule: every parenthesis must be matched within its logical statement. VBA has
 * no cross-statement parentheses (a `(` is closed before the line ends unless a
 * `_` line-continuation joins the next physical line, which the lexer already
 * folds into trivia), so an open `(` left dangling at a statement boundary, or a
 * `)` with no matching `(`, is always the VBE "Expected: )" / "Syntax error".
 *
 * The scan walks the whole module's token stream, tracking paren depth and
 * resetting at each logical-statement boundary (a `newline` token or a depth-0
 * `:` statement separator). Only literal `(`/`)` punctuation tokens count -
 * parentheses inside strings, comments, date literals, and `[bracketed]` names
 * are distinct token kinds, so they can never create a false positive. At most
 * one diagnostic is reported per statement.
 */
export function checkUnbalancedParens(
	source: string,
	push: PushFn,
	activity?: ConditionalActivityTracker,
): void {
	// Text under an inactive `#If` arm is never compiled, and `#If False Then`
	// is a common place to park notes (issue #102).
	const toks = activity
		? tokenizeCached(source).filter((tok) => !activity.isInactive({ start: tok.start, end: tok.end }))
		: tokenizeCached(source);
	let depth = 0;
	const openOffsets: number[] = [];
	let flagged = false;

	const flush = (): void => {
		if (!flagged && depth > 0) {
			const off = openOffsets[0];
			push(
				'unbalancedParens',
				"Unbalanced parentheses: a ')' is missing.",
				{ start: off, end: off + 1 },
			);
		}
		depth = 0;
		openOffsets.length = 0;
		flagged = false;
	};

	for (const tok of toks) {
		if (tok.kind === 'newline') {
			flush();
			continue;
		}
		if (tok.kind === 'colon' && depth === 0) {
			flush();
			continue;
		}
		if (tok.kind !== 'punctuation') {
			continue;
		}
		if (tok.rawText === '(') {
			depth++;
			openOffsets.push(tok.start);
		} else if (tok.rawText === ')') {
			if (depth === 0) {
				if (!flagged) {
					push(
						'unbalancedParens',
						"Unbalanced parentheses: an unexpected ')' was found.",
						{ start: tok.start, end: tok.end },
					);
					flagged = true;
				}
			} else {
				depth--;
				openOffsets.pop();
			}
		}
	}
	flush();
}

/**
 * Rule: a `Call` statement must wrap its arguments in parentheses. After the
 * `Call` keyword the callee chain (identifier, then any run of `.member` or
 * `(...)` groups) is consumed; any token left over is an unparenthesised
 * argument - the VBE "Expected: (" error. Unbalanced parentheses are left to the
 * dedicated rule.
 */
export function checkCallParens(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): ProcedureStatementVisitor {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		return (stmt) => {
			const invalidCallTarget = invalidExplicitCallTarget(source, stmt.span, moduleSignatures, sourceNames);
			if (invalidCallTarget) {
				push(
					'invalidExplicitCallTarget',
					`'${invalidCallTarget.name}' cannot be used as the target of an explicit Call statement.`,
					invalidCallTarget.span,
				);
				return;
			}
			const at = explicitCallStatementArgumentWithoutParens(source, stmt.span);
			if (at) {
				push(
					'callRequiresParens',
					'A Call statement requires parentheses around its argument list.',
					at,
				);
			}
			const bare = implicitParenthesizedBareCallableCall(source, stmt.span, moduleSignatures, sourceNames);
			if (bare) {
				push(
					'callStatementForbidsParens',
					bareCallForbidsParensMessage(bare.name, moduleSignatures, sourceNames),
					bare.span,
				);
			}
			const multiArg = implicitParenthesizedMultiArgCall(source, stmt.span, moduleSignatures, sourceNames);
			if (multiArg) {
				push(
					'callStatementMultiArgParens',
					`A standalone call cannot enclose multiple arguments in parentheses; use 'Call ${multiArg.name}(...)' or remove the parentheses ('${multiArg.name} arg1, arg2'). VBA rejects this form as a compile error.`,
					multiArg.span,
				);
			}
			const implicit = implicitParenthesizedMemberCall(source, stmt.span, memberCtx);
			if (implicit) {
				push(
					'callStatementForbidsParens',
					'Standalone zero-argument member calls cannot use empty parentheses unless they are prefixed with Call or used in an expression.',
					implicit.span,
				);
			}
		};
	};
}

function bareCallForbidsParensMessage(
	name: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
): string {
	const runtime = !moduleSignatures.has(name.toLowerCase()) &&
		!runtimeCallableSourceShadowed(name, sourceNames)
		? resolveRuntimeFunction(name)
		: undefined;
	if (runtime && !runtimeAllowsExplicitCall(runtime)) {
		return `Standalone '${runtime.name}()' cannot use empty parentheses in statement context; use '${runtime.name}' as a statement or use it in an expression.`;
	}
	return 'Standalone zero-argument procedure calls cannot use empty parentheses unless they are prefixed with Call or used in an expression.';
}

function invalidExplicitCallTarget(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
): { name: string; span: Span } | undefined {
	const target = explicitCallStatementTarget(source, span);
	if (!target) {
		return undefined;
	}
	if (
		moduleSignatures.has(target.name.toLowerCase()) ||
		runtimeCallableSourceShadowed(target.name, sourceNames)
	) {
		return undefined;
	}
	const runtime = resolveRuntimeFunction(target.name);
	if (!runtime || runtimeAllowsExplicitCall(runtime)) {
		return undefined;
	}
	return { name: runtime.name, span: target.span };
}

export function checkInvalidExpressionSyntax(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		return (stmt) => {
			const incompleteMember = incompleteMemberAccess(source, stmt.span, {
				scalarTypes: env,
				resolveScalarType: (name) => declaredTypeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'memberReceiver',
				),
			});
			if (incompleteMember) {
				push(
					'invalidExpressionSyntax',
					"Incomplete member access: type a member name after '.'.",
					incompleteMember.span,
				);
				return;
			}
			const unsupportedQuestion = unsupportedQuestionMarkOperator(source, stmt.span);
			if (unsupportedQuestion) {
				push(
					'invalidExpressionSyntax',
					"VBA does not support the '?' conditional operator in code modules; use If...Then...Else, or IIf(...) only when both branches are safe to evaluate.",
					unsupportedQuestion.span,
				);
				return;
			}
			const hit = invalidOperatorSequence(source, stmt.span);
			if (hit) {
				push(
					'invalidExpressionSyntax',
					hit.message ?? `Invalid operator sequence '${hit.text}'; this will fail to compile as a syntax error.`,
					hit.span,
				);
				return;
			}
			const juxtaposed = juxtaposedRhsValues(source, stmt.span);
			if (juxtaposed) {
				push(
					'invalidExpressionSyntax',
					`Unexpected '${juxtaposed.text}' after a complete expression; expected end of statement. This will fail to compile as a syntax error.`,
					juxtaposed.span,
				);
			}
		};
	};
}

const NON_UNARY_BINARY_OPERATORS = new Set([
	'*',
	'/',
	'\\',
	'^',
	'&',
	'=',
	'<',
	'>',
	'<=',
	'>=',
	'<>',
	':=',
	'like',
	'is',
	'and',
	'or',
	'xor',
	'eqv',
	'imp',
	'mod',
]);

/**
 * Whether the `&` at `index` is glued to a name before it, which makes it the
 * name's Long type-declaration character (`total&`) rather than the
 * concatenation operator. `s$`, `n%`, `x!`, `d#` and `c@` lex the same way;
 * only `&` doubles as an operator, so only it needs asking.
 */
export function isGluedTypeSuffixAmpersand(toks: readonly VbaToken[], index: number): boolean {
	const tok = toks[index];
	const prev = toks[index - 1];
	return tok?.kind === 'operator' && tok.rawText === '&' && prev !== undefined
		&& prev.end === tok.start && tokenName(prev) !== undefined;
}

function invalidOperatorSequence(
	source: string,
	span: Span,
): { text: string; span: Span; message?: string } | undefined {
	const toks = statementTokens(source, span);
	// A Case statement's Is-comparison clause (MS-VBAL 5.4.2.10, `Case Is > 5`)
	// uses `Is` as grammar, not as the object-identity operator, so the
	// word-operator scan would mis-read `Is >` as an impossible operator run.
	// Case clauses are grammar, not value expressions; skip the whole statement
	// (the Select/Case rules own its structure).
	if (tokenText(toks[firstExecutableTokenIndex(toks)]) === 'case') {
		return undefined;
	}
	for (let i = 0; i < toks.length; i++) {
		// `Not` with nothing after it to negate: a line holding only `Not`, or
		// `x = Not` (issue #234, measured in Excel 16.0: Syntax error).
		if (tokenText(toks[i]) === 'not' && toks[i].kind === 'keyword' && (i === toks.length - 1 || toks[i + 1].kind === 'comment')) {
			return {
				text: toks[i].rawText,
				span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
				message: "'Not' has nothing after it to negate. This is a VBE compile error: Syntax error.",
			};
		}
		if (!isNonUnaryBinaryOperator(toks[i])) {
			continue;
		}
		// `total& = 3`: an `&` glued to the name before it is the Long
		// type-declaration character, not concatenation. The VBE reads it that
		// way whatever follows - `a& b` is a syntax error there, `a &b` is a
		// concatenation (issue #100, measured in Excel 16.0).
		if (isGluedTypeSuffixAmpersand(toks, i)) {
			continue;
		}
		// `x^=5` assigns the LongLong x^ (issue #369, measured in Excel 16.0);
		// elsewhere a glued `^` is the power operator, `a^2`.
		if (toks[i].rawText === '^' && i === firstExecutableTokenIndex(toks) + 1 && toks[i - 1].end === toks[i].start
			&& tokenName(toks[i - 1]) !== undefined && toks[i + 1]?.rawText === '=') {
			continue;
		}
		// `a < > b` is one relational operator written as two tokens (MS-VBAL
		// 5.6.9.5), and the VBE reads it as `a <> b` (issue #87).
		const operatorEnd = i + (relationalOperatorAt(toks, i)?.length ?? 1) - 1;
		let end = operatorEnd;
		while (isNonUnaryBinaryOperator(toks[end + 1])) {
			end++;
		}
		if (end > operatorEnd || operatorEnd === toks.length - 1) {
			const first = toks[i];
			const last = toks[end];
			return {
				text: source.slice(span.start + first.start, span.start + last.end),
				span: { start: span.start + first.start, end: span.start + last.end },
			};
		}
		i = operatorEnd;
	}
	return undefined;
}

const JUXTAPOSABLE_VALUE_KINDS = new Set([
	'integerLiteral',
	'floatLiteral',
	'dateLiteral',
	'stringLiteral',
	'identifier',
	'bracketedIdentifier',
]);

function isJuxtaposableValueStart(tok: VbaToken | undefined): boolean {
	return tok !== undefined && JUXTAPOSABLE_VALUE_KINDS.has(tok.kind);
}

function endsJuxtaposableValue(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return false;
	}
	// A digit run glued to `&` lexes as a &-suffixed integer literal, but the
	// VBE can read that `&` as CONCATENATION - it does when the digits overflow
	// Long (VBE oracle suffix_long_amp_glued_concat_accepted: `s = 3000000000&"x"`
	// is accepted) - so a &-suffixed integer literal never provably ends a
	// value. The in-range form (`n = 5& 1`) is under-reported by design: a
	// missed diagnostic beats a false positive on the oracle-verified concat.
	if (tok.kind === 'integerLiteral' && tok.rawText.endsWith('&')) {
		return false;
	}
	return JUXTAPOSABLE_VALUE_KINDS.has(tok.kind) || tok.rawText === ')' || tok.rawText === ']';
}

/**
 * Detects two juxtaposed value expressions in an assignment RHS - a complete value
 * (literal / identifier / call / index) immediately followed by another value
 * starter with no operator between, e.g. `n = 1 n 1` or `n = 1 MsgBox("hello") 1`.
 * That is a VBE "Expected: end of statement" syntax error which the lenient parser
 * otherwise silently drops to a raw statement.
 *
 * Scoped to the TOP LEVEL of an assignment RHS (a top-level standalone `=` on a
 * statement that is not a non-assignment leader) so it cannot misfire on: implicit
 * call statements (`MsgBox x` - no `=`), a call written with a space (`Foo (x)` -
 * the next token is `(`, not a value start), jagged-array access (`arr(1)(2)` - `(`
 * again), a trailing type-suffix/operator (`Count&` - `&` is not a value start), or
 * anything inside parentheses (depth > 0 is skipped).
 */
function juxtaposedRhsValues(
	source: string,
	span: Span,
): { text: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	if (toks.length === 0 || isNonAssignmentStatementLeader(tokenText(toks[firstExecutableTokenIndex(toks)]))) {
		return undefined;
	}
	const eq = topLevelOperatorIndex(toks, '=');
	if (eq < 0) {
		return undefined;
	}
	const at = juxtaposedValueIndex(toks, eq + 1);
	return at < 0 ? undefined : { text: toks[at].rawText, span: absoluteSpan(span, toks[at]) };
}

/**
 * The index of a value that follows a complete value with no operator between,
 * `asdf qwer` or `1 n`, from `from` on at the top level; -1 when there is none.
 * A Const's value is read the same way (issue #234).
 */
export function juxtaposedValueIndex(toks: readonly VbaToken[], from: number): number {
	let depth = 0;
	for (let i = from; i + 1 < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
			continue;
		}
		if (raw === ')' || raw === ']') {
			depth = depth > 0 ? depth - 1 : 0;
		}
		if (depth !== 0) {
			continue;
		}
		if (endsJuxtaposableValue(toks[i]) && isJuxtaposableValueStart(toks[i + 1])) {
			return i + 1;
		}
	}
	return -1;
}

function unsupportedQuestionMarkOperator(
	source: string,
	span: Span,
): { span: Span } | undefined {
	const question = statementTokens(source, span).find(
		(tok) => tok.kind === 'operator' && tok.rawText === '?',
	);
	return question ? { span: absoluteSpan(span, question) } : undefined;
}

export function incompleteMemberAccess(
	source: string,
	span: Span,
	options: {
		includeLeadingDot?: boolean;
		scalarTypes?: ReadonlyMap<string, string>;
		resolveScalarType?: SourceDeclaredTypeResolver;
	} = {},
): { span: Span } | undefined {
	const toks = statementTokens(source, span);
	for (let i = 0; i < toks.length; i++) {
		const tok = toks[i];
		if (tok.rawText !== '.') {
			continue;
		}
		if (i === 0 && !options.includeLeadingDot) {
			continue;
		}
		const next = toks[i + 1];
		if (next && tokenName(next)) {
			continue;
		}
		const receiverName = i > 0 ? tokenName(toks[i - 1]) : undefined;
		if (receiverName) {
			const resolvedType = options.resolveScalarType?.(receiverName);
			const asType = resolvedType?.resolved
				? resolvedType.asType
				: options.scalarTypes?.get(receiverName.toLowerCase());
			const normalized = normalizeType(asType);
			if (normalized && isKnownScalarType(normalized)) {
				continue;
			}
		}
		return { span: absoluteSpan(span, tok) };
	}
	return undefined;
}

export function isNonUnaryBinaryOperator(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return false;
	}
	// VBA word operators (And/Or/Xor/Eqv/Imp/Mod/Like/Is) lex as 'keyword'
	// tokens, so accept those alongside symbolic 'operator' tokens (e.g. ':=')
	// by matching on the lowercased text rather than the token kind.
	if (tok.kind !== 'operator' && tok.kind !== 'keyword') {
		return false;
	}
	return NON_UNARY_BINARY_OPERATORS.has(tokenText(tok));
}

export function checkDivisionByZeroExpressions(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	hostModel?: HostObjectModel,
): ProcedureStatementVisitor {
	const projectConstants = resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map());
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, projectConstants);
	const types = moduleTypes(source, mod, activity);
	return (member) => {
		const constants = procedureIntegerConstantLookup(
			member, moduleConstants, symbols, projectVisibleSymbols, activity, hostModel,
		);
		// A numeric member of a Type local is 0 until an assignment stores
		// another literal (issue #253): `1 / t.a` raises 11.
		const membersAt = types.size === 0 ? undefined : typeMemberStatesAt(source, symbols, member, types, activity, moduleOptionBase(mod, activity));
		let members: ReadonlyMap<string, MemberState> = new Map();
		// A local the procedure never assigns is 0, and one whose every
		// assignment is `d = 0` is 0 too (issue #119): `10 / d` raises 11.
		// So is one the last assignment before the division sets to 0, though
		// a later one changes it (issue #180).
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		let known: ReadonlyMap<string, KnownLocalValue> = new Map();
		// A Function of the module that returns 0: `10 / F()` (issue #448).
		const results = knownFunctionResults(source, mod, activity);
		// A loop counter on its first and last passes: `For i = 0 To 3` then
		// `1 / i` divides by 0 on the first (issue #263).
		const counters = loopCountersAt(source, member.body, activity);
		let passValues: ReadonlyMap<string, number> = new Map();
		let heldNow: (lower: string) => readonly VbaToken[] | undefined = () => undefined;
		let emptyCollectionNow: (lower: string) => boolean = () => false;
		let defaulted: ReturnType<typeof defaultedStraightLine> | undefined;
		const lookup: IntegerConstantLookup = {
			get: (name) => {
				const pass = passValues.get(name.toLowerCase());
				if (pass !== undefined) {
					return pass;
				}
				const constant = constants.get(name);
				if (constant !== undefined) {
					return constant;
				}
				const local = known.get(name.toLowerCase());
				if (local?.kind === 'empty') {
					return 0;
				}
				// "0" divides as 0 (issue #491, measured in Excel 16.0).
				if (local?.kind === 'string' && !local.contentMutated && /^\s*[+-]?(0+\.?0*|\.0+)\s*$/.test(local.value as string)) {
					return 0;
				}
				// A Variant the straight line just set to Empty or CDec(0).
				const held = heldNow(name.toLowerCase());
				if (held && ((held.length === 1 && tokenText(held[0]) === 'empty') || zeroConversionCallEnd(held, 0) === held.length - 1)) {
					return 0;
				}
				const field = members.get(name.toLowerCase());
				if (isKnownNumber(field) && Number.isInteger(field.number)) {
					return field.number;
				}
				// `c.Count` of a Collection nothing has added to (issue #614,
				// measured in Excel 16.0: `10 / c.Count` raises 11).
				if (/^[a-z_][a-z0-9_]*[.]count$/i.test(name) && emptyCollectionNow(name.slice(0, -'.count'.length).toLowerCase())) {
					return 0;
				}
				if (local?.kind === 'number' && Number.isInteger(local.value)) {
					return local.value as number;
				}
				return local ? undefined : functionIntegerResult(name, results, member, symbols);
			},
		};
		const guards = divisionGuardRanges(member.body, activity);
		// A statement a known guard keeps from running (issue #273).
		const unreachable = unreachableStatementsIn(source, member, symbols, activity);
		// `d = 0.4` then `10 Mod d` (issue #239).
		const fractionOf = (lower: string): number | undefined => {
			const local = known.get(lower);
			return local?.kind === 'number' && constants.get(lower) === undefined ? (local.value as number) : undefined;
		};
		// A dividend that is Null, or a Variant the straight line just set to
		// Null: Null divided by zero is Null and raises nothing (issue #282,
		// measured in Excel 16.0).
		let reaching: ReturnType<typeof straightLineAssignments> | undefined;
		const nullAt = (stmt: LeafStatementNode) => (tok: VbaToken | undefined): boolean => {
			if (tokenText(tok) === 'null') {
				return true;
			}
			const lower = tok ? tokenName(tok)?.toLowerCase() : undefined;
			const held = lower ? (reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((t) => t.kind !== 'comment') : undefined;
			return held?.length === 1 && tokenText(held[0]) === 'null';
		};
		// A number past the Long range, written or just stored: a literal, or a
		// CDec of one (issue #502).
		const pastLongAt = (stmt: LeafStatementNode) => (tok: VbaToken | undefined): boolean => {
			const outside = (value: VbaToken | undefined): boolean => {
				const text = value?.kind === 'stringLiteral' ? value.rawText.slice(1, -1) : value?.kind === 'integerLiteral' || value?.kind === 'floatLiteral' ? value.rawText.replace(/[!#@%&^]$/, '') : undefined;
				const number = text !== undefined && /^\s*[+-]?\d+(\.\d*)?([eE][+-]?\d+)?\s*$/.test(text) ? Number(text) : NaN;
				return Number.isFinite(number) && Math.abs(Math.round(number)) > 2147483647;
			};
			if (outside(tok)) {
				return true;
			}
			const lower = tok ? tokenName(tok)?.toLowerCase() : undefined;
			const held = lower ? (reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((t) => t.kind !== 'comment') : undefined;
			const conversion = held?.length === 4 && tokenText(held[0]) === 'cdec' && held[1].rawText === '(' && held[3].rawText === ')';
			return held !== undefined && ((held.length === 1 && outside(held[0])) || (conversion && outside(held[2])));
		};
		return (stmt) => {
			if (unreachable.has(stmt)) {
				return;
			}
			known = valuesAt(stmt);
			heldNow = (lower) => (reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((t) => t.kind !== 'comment');
			emptyCollectionNow = (lower) => (defaulted ??= defaultedStraightLine(source, member, symbols, activity)).get(stmt)?.get(lower) === EMPTY_COLLECTION;
			// A single-line If's branch sees the members less what its condition names.
			members = membersAt ? membersAt(stmt, stmt.span.end) : members;
			checkEachCounterPass(source, stmt.span, counters.get(stmt), () => undefined, (values, report) => {
				passValues = values;
				for (const hit of divisionByZeroDivisors(source, stmt.span, lookup, guards, fractionOf, nullAt(stmt), pastLongAt(stmt))) {
					report('divisionByZero', hit.message, hit.span);
				}
			}, push);
			passValues = new Map();
		};
	};
}

/**
 * Rule: an arithmetic operator raises error 13 for a nonnumeric string operand
 * whatever the result goes into (issue #119; each measured in Excel 16.0):
 * `v = "abc" + 1` into a Variant, `Main = "abc" + 1` as a function result,
 * `Not "abc"`, `-"abc"`, `If "abc" = 1 Then`, and `s * 2` with s holding
 * "abc". The assignment and argument rules only saw the numeric-target case.
 * `+` and the comparisons need a NUMBER on the other side, because two strings
 * concatenate and compare as text; `- * / \ ^ Mod` and the unary forms
 * always coerce.
 */
export function checkStringArithmeticOperands(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const optionBase = moduleOptionBase(mod, activity);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const valuesAt = knownLocalLiteralValuesAt(source, member, symbols, activity);
		// A statement a known guard keeps from running (issue #273).
		const unreachable = unreachableStatementsIn(source, member, symbols, activity);
		let known: ReadonlyMap<string, KnownLocalValue> = new Map();
		// The arrays the locals hold at the statement, read only when an
		// operand indexes one: `v(1) + 1` with v = Array("1", "b") (issue #260).
		let shapesAt: ((stmt: LeafStatementNode) => ReadonlyMap<string, FixedArrayBound>) | undefined;
		let current: LeafStatementNode | undefined;
		let shapes: ReadonlyMap<string, FixedArrayBound> | undefined;
		let written: ReadonlySet<string> | undefined;
		const shapesHere = (): ReadonlyMap<string, FixedArrayBound> => {
			if (!current) {
				return EMPTY_SHAPES;
			}
			if (!shapes) {
				shapesAt ??= knownArrayShapesAt(source, symbols, member, activity, optionBase);
				written ??= elementsWrittenIn(source, member, activity);
				const all = shapesAt(current);
				shapes = written.size === 0 ? all : new Map([...all].filter(([lower]) => !written!.has(lower)));
			}
			return shapes;
		};
		const elementSpan = (spanStart: number, toks: readonly VbaToken[], element: ElementOperand): Span => ({
			start: spanStart + toks[element.first].start,
			end: spanStart + toks[element.last].end,
		});
		const elementString = (spanStart: number, toks: readonly VbaToken[], element: ElementOperand | undefined): string | undefined => {
			if (typeof element?.value !== 'string' || !isProvablyNonNumericString(element.value)) {
				return undefined;
			}
			const span = elementSpan(spanStart, toks, element);
			return `'${source.slice(span.start, span.end)}', which holds ${JSON.stringify(element.value)}`;
		};
		const stringConsts = stringConstantsInScope(symbols, member);
		const constantOf = (tok: VbaToken): string | undefined => {
			const name = tokenName(tok)?.toLowerCase();
			return name && !known.has(name) ? stringConsts.get(name) : undefined;
		};
		const nonnumericString = (tok: VbaToken | undefined): string | undefined => {
			if (!tok) {
				return undefined;
			}
			if (tok.kind === 'stringLiteral') {
				const value = stringLiteralValue(tok.rawText);
				return isProvablyNonNumericString(value) ? `string literal ${tok.rawText}` : undefined;
			}
			const name = tokenName(tok)?.toLowerCase();
			const local = name ? known.get(name) : undefined;
			if (local?.kind === 'string' && !local.contentMutated && isProvablyNonNumericString(local.value as string)) {
				return `'${tok.rawText}', which holds ${JSON.stringify(local.value)}`;
			}
			const constant = constantOf(tok);
			return constant !== undefined && isProvablyNonNumericString(constant) ? `constant '${tok.rawText}', which is ${JSON.stringify(constant)}` : undefined;
		};
		// A condition converts to Boolean: "True" and numbers run, "yes"
		// and " True " raise 13 (issue #191).
		const nonBooleanString = (tok: VbaToken | undefined): string | undefined => {
			if (tok?.kind === 'stringLiteral') {
				return isInvalidBooleanString(stringLiteralValue(tok.rawText)) ? `string literal ${tok.rawText}` : undefined;
			}
			const name = tok ? tokenName(tok)?.toLowerCase() : undefined;
			const local = name ? known.get(name) : undefined;
			if (local?.kind === 'string' && !local.contentMutated && isInvalidBooleanString(local.value as string)) {
				return `'${tok!.rawText}', which holds ${JSON.stringify(local.value)}`;
			}
			return undefined;
		};
		const numeric = (tok: VbaToken | undefined): boolean => {
			if (!tok) {
				return false;
			}
			// A date adds and compares as a number: #1/1/2000# + "abc" raises.
			if (tok.kind === 'integerLiteral' || tok.kind === 'floatLiteral' || tok.kind === 'dateLiteral') {
				return true;
			}
			const name = tokenName(tok)?.toLowerCase();
			if (!name) {
				return false;
			}
			if (known.get(name)?.kind === 'number') {
				return true;
			}
			const type = normalizeType(env.get(name));
			return type !== undefined && isNumericType(type);
		};
		// A comparison converts only between typed operands: a Variant on
		// either side compares without converting, so `v = "abc"` with v
		// holding 5 is False and `v = 5` with v holding "abc" is False too
		// (issue #268, measured in Excel 16.0). The number side is a literal
		// or a name declared a number type, and the string side a literal,
		// a String or a Const.
		const declaredType = (tok: VbaToken | undefined): string | undefined => {
			const name = tok ? tokenName(tok)?.toLowerCase() : undefined;
			return name ? normalizeType(env.get(name)) : undefined;
		};
		const typedNumber = (tok: VbaToken | undefined): boolean => {
			if (tok?.kind === 'integerLiteral' || tok?.kind === 'floatLiteral') {
				return true;
			}
			const type = declaredType(tok);
			return type !== undefined && isNumericType(type);
		};
		const typedString = (tok: VbaToken | undefined): boolean => {
			if (!tok || tok.kind === 'stringLiteral') {
				return tok !== undefined;
			}
			const type = declaredType(tok);
			return type === 'string' || (type === undefined && constantOf(tok) !== undefined);
		};
		/** The string a typed string operand holds, for a Date or Boolean comparison. */
		const typedStringValue = (tok: VbaToken | undefined): { value: string; what: string } | undefined => {
			if (!tok || !typedString(tok)) {
				return undefined;
			}
			if (tok.kind === 'stringLiteral') {
				return { value: stringLiteralValue(tok.rawText), what: `string literal ${tok.rawText}` };
			}
			const local = known.get(tokenName(tok)!.toLowerCase());
			if (local?.kind === 'string' && !local.contentMutated) {
				return { value: local.value as string, what: `'${tok.rawText}', which holds ${JSON.stringify(local.value)}` };
			}
			const constant = constantOf(tok);
			return constant === undefined ? undefined : { value: constant, what: `constant '${tok.rawText}', which is ${JSON.stringify(constant)}` };
		};
		/** A typed Date or Boolean compared with a string it cannot read: `d = "abc"`, `b = "yes"`. */
		const unreadableAs = (typed: VbaToken | undefined, other: VbaToken | undefined): { what: string; as: string } | undefined => {
			const type = declaredType(typed);
			const string = type === 'date' || type === 'boolean' ? typedStringValue(other) : undefined;
			if (!string) {
				return undefined;
			}
			const invalid = type === 'date' ? isInvalidDateString(string.value) : isInvalidBooleanString(string.value);
			return invalid ? { what: string.what, as: type === 'date' ? 'a Date' : 'a Boolean' } : undefined;
		};
		const reportConversion = (span: Span, what: string, into: string): void => {
			push('stringArithmeticCoercion', `${into.replace('WHAT', what)}. This will raise Run-time error '13': Type mismatch.`, span);
		};
		// A condition that is one string: `If answer Then`, `Do While "abc"`.
		const checkCondition = (spanStart: number, toks: readonly VbaToken[], from: number, to: number, keyword: string): void => {
			if (to - from !== 1) {
				return;
			}
			const what = nonBooleanString(toks[from]);
			if (what) {
				reportConversion(absoluteSpan({ start: spanStart, end: spanStart }, toks[from]), what, `'${keyword}' converts WHAT to Boolean`);
			}
		};
		const conditionEnd = (toks: readonly VbaToken[], from: number): number => {
			const then = toks.findIndex((tok, k) => k >= from && tokenText(tok) === 'then');
			const comment = toks.findIndex((tok, k) => k >= from && tok.kind === 'comment');
			return then >= 0 ? then : comment >= 0 ? comment : toks.length;
		};
		// Block headers and footers are no statements of their own, so the
		// walk reads them here: If, Do and Loop conditions, While, a For
		// loop's bounds and Case values against a number (issue #191).
		known = valuesAt(undefined);
		// Only a string literal, or a local known to hold a string, can be
		// reported; a header with neither is not read.
		const knowsStrings = [...known.values()].some((value) => value.kind === 'string');
		const mayHoldString = (span: Span): boolean => knowsStrings || source.slice(span.start, span.end).includes('"');
		let forReaching: ReturnType<typeof straightLineAssignments> | undefined;
		const visitBlocks = (body: readonly BodyNode[]): void => {
			for (const node of body) {
				if (activity?.isInactive(node.span) || unreachable.has(node) || !('body' in node) || !Array.isArray(node.body)) {
					continue;
				}
				// `For i = 1 To v` with v a Variant the straight line just set to
				// Null: a bound must be a number (issue #332, measured: 94).
				if (node.kind === 'ForBlock' && !node.each) {
					const header = blockHeaderLineSpan(source, node.span);
					const headToks = statementTokens(source, header);
					const to = headToks.findIndex((tok) => tokenText(tok) === 'to');
					const step = headToks.findIndex((tok) => tokenText(tok) === 'step');
					const eq = headToks.findIndex((tok) => tok.rawText === '=');
					for (const [from, until] of [[eq + 1, to], [to + 1, step > 0 ? step : headToks.length], [step + 1, step > 0 ? headToks.length : -1]]) {
						const lower = until - from === 1 && from > 0 ? tokenName(headToks[from])?.toLowerCase() : undefined;
						const held = lower ? (forReaching ??= straightLineAssignments(source, member.body, activity)).get(node)?.get(lower)?.filter((tok) => tok.kind !== 'comment') : undefined;
						if (held?.length === 1 && tokenText(held[0]) === 'null') {
							push('variantValueMisuse', `'${headToks[from].rawText}' holds Null here, and For needs a number for its bounds. This will raise Run-time error '94': Invalid use of Null.`, absoluteSpan(header, headToks[from]));
						}
					}
				}
				// The opening line sees what reaches the block: `x = "abc"` then
				// `While x`, though the body assigns x again (issue #424).
				known = valuesAt(node);
				const reachingString = [...known.values()].some((value) => value.kind === 'string');
				if (node.kind !== 'SelectBlock' && !reachingString && !mayHoldString(blockHeaderLineSpan(source, node.span)) && !(node.kind === 'DoBlock' && mayHoldString(blockFooterLineSpan(source, node.span)))) {
					known = valuesAt(undefined);
					visitBlocks(node.body as BodyNode[]);
					continue;
				}
				const header = blockHeaderLineSpan(source, node.span);
				const headToks = statementTokens(source, header);
				const head = tokenText(headToks[0]);
				scanOperators(header.start, headToks, -1);
				if (node.kind === 'IfBlock' && head === 'if') {
					checkCondition(header.start, headToks, 1, conditionEnd(headToks, 1), 'If');
				} else if ((node.kind === 'DoBlock' || node.kind === 'WhileBlock') && headToks.length > 1) {
					const keyword = tokenText(headToks[head === 'do' ? 1 : 0]);
					if (keyword === 'while' || keyword === 'until') {
						const from = head === 'do' ? 2 : 1;
						checkCondition(header.start, headToks, from, conditionEnd(headToks, from), keyword === 'while' ? 'While' : 'Until');
					}
				}
				if (node.kind === 'DoBlock') {
					const footer = blockFooterLineSpan(source, node.span);
					const footToks = statementTokens(source, footer);
					if (tokenText(footToks[0]) === 'loop' && (tokenText(footToks[1]) === 'while' || tokenText(footToks[1]) === 'until')) {
						// The Loop line runs after the body, with what the body leaves.
						const entering = known;
						known = valuesAt(undefined);
						scanOperators(footer.start, footToks, -1);
						checkCondition(footer.start, footToks, 2, conditionEnd(footToks, 2), tokenText(footToks[1]) === 'while' ? 'While' : 'Until');
						known = entering;
					}
				}
				if (node.kind === 'ForBlock' && !node.each && node.controlVariable && numeric({ kind: 'identifier', rawText: node.controlVariable, start: 0, end: 0 } as VbaToken)) {
					checkForBounds(header.start, headToks);
				}
				// The selector is judged with each value: a typed number, Date or
				// Boolean converts; a Variant does not.
				if (node.kind === 'SelectBlock' && head === 'select' && conditionEnd(headToks, 2) === 3) {
					for (const item of node.body) {
						if (isLeafStatement(item) && !activity?.isInactive(item.span)) {
							checkCaseValues(item.span, headToks[2]);
						}
					}
				}
				visitBlocks(node.body as BodyNode[]);
			}
		};
		// `For i = 1 To "abc"`: a bound converts to the counter's number.
		const checkForBounds = (spanStart: number, toks: readonly VbaToken[]): void => {
			const eq = toks.findIndex((tok) => tok.rawText === '=');
			const to = toks.findIndex((tok) => tokenText(tok) === 'to');
			const step = toks.findIndex((tok) => tokenText(tok) === 'step');
			const end = conditionEnd(toks, 0);
			const bounds: Array<[number, number, string]> = [[eq + 1, to, 'start'], [to + 1, step > 0 ? step : end, 'end'], ...(step > 0 ? [[step + 1, end, 'step'] as [number, number, string]] : [])];
			for (const [from, until, which] of bounds) {
				const what = until - from === 1 && from > 0 ? nonnumericString(toks[from]) : undefined;
				if (what) {
					reportConversion(absoluteSpan({ start: spanStart, end: spanStart }, toks[from]), what, `For converts WHAT to a number for its ${which}`);
				}
			}
		};
		// `Select Case 1` then `Case "abc"`: each value is compared as a number.
		// The selector is a typed number, Date or Boolean: a Variant selector
		// compares without converting (issue #268).
		const checkCaseValues = (span: Span, selector: VbaToken): void => {
			const toks = statementTokens(source, span);
			if (tokenText(toks[0]) !== 'case' || tokenText(toks[1]) === 'else') {
				return;
			}
			const end = conditionEnd(toks, 1);
			const judge = (value: VbaToken | undefined): boolean => {
				if (!value) {
					return false;
				}
				const unreadable = unreadableAs(selector, value);
				if (unreadable) {
					reportConversion(absoluteSpan(span, value), unreadable.what, `Case compares WHAT with ${unreadable.as}, which cannot read it`);
					return true;
				}
				// A Case value converts to the selector's type even from a
				// Variant: `Select Case n` on a Long with `Case w`, w holding
				// "abc", raises 13 (measured), where `w = n` runs.
				const what = typedNumber(selector) ? nonnumericString(value) : undefined;
				if (what) {
					reportConversion(absoluteSpan(span, value), what, 'Case compares WHAT with a number');
				}
				return what !== undefined;
			};
			let from = 1;
			for (let k = 1; k <= end; k++) {
				if (k === end || toks[k].rawText === ',') {
					if (k - from === 1) {
						judge(toks[from]);
					} else if (k - from === 3 && tokenText(toks[from]) === 'is' && toks[from + 1].kind === 'operator') {
						// `Case Is > "abc"` compares the same way (issue #243).
						judge(toks[from + 2]);
					} else if (k - from === 3 && tokenText(toks[from + 1]) === 'to') {
						// `Case "a" To "z"`: the first end that fails (issue #268).
						if (!judge(toks[from])) {
							judge(toks[from + 2]);
						}
					}
					from = k + 1;
				}
			}
		};
		// Logical operators convert a string operand to a number, and bind
		// loosest: in `s = "yes" Or t` the string is the `=`'s. An operand is
		// judged only when it stands alone between the operator and a
		// boundary.
		const standsAlone = (toks: readonly VbaToken[], index: number): boolean => {
			const tok = toks[index];
			if (!tok) {
				return true;
			}
			const word = tokenText(tok);
			return tok.rawText === '(' || tok.rawText === ')' || tok.rawText === ',' || tok.rawText === ':' || tok.kind === 'comment'
				|| LOGICAL_OPERATORS.has(word) || word === 'then' || word === 'if' || word === 'elseif' || word === 'while'
				|| word === 'until' || word === 'not';
		};
		function scanOperators(spanStart: number, toks: readonly VbaToken[], assignIndex: number, branchAssigns: ReadonlySet<number> = new Set()): void {
			const at = (tok: VbaToken): Span => absoluteSpan({ start: spanStart, end: spanStart }, tok);
			for (let i = 0; i < toks.length; i++) {
				if (i === assignIndex || branchAssigns.has(i)) {
					continue;
				}
				const tok = toks[i];
				const word = tokenText(tok);
				const left = toks[i - 1];
				const right = toks[i + 1];
				const report = (span: Span, what: string): void => {
					push(
						'stringArithmeticCoercion',
						`Operator '${tok.rawText}' coerces ${what} to a number. This will raise Run-time error '13': Type mismatch.`,
						span,
					);
				};
				if (LOGICAL_OPERATORS.has(word) && tok.kind === 'keyword') {
					const leftString = i - 1 !== assignIndex && !branchAssigns.has(i - 1) && (i - 2 === assignIndex || branchAssigns.has(i - 2) || standsAlone(toks, i - 2)) ? nonnumericString(left) : undefined;
					const rightString = standsAlone(toks, i + 2) ? nonnumericString(right) : undefined;
					if (leftString) {
						report(at(left), leftString);
					} else if (rightString) {
						report(at(right!), rightString);
					}
					continue;
				}
				const isBinary = tok.kind === 'operator'
					? ['+', '-', '*', '/', '\\', '^', '=', '<', '>', '<=', '>=', '<>'].includes(tok.rawText)
					: word === 'mod';
				// A keyword ends an operand unless the statement's words start
				// one there: `If Not s` (issue #268).
				const leftEndsOperand = left !== undefined && (left.kind === 'identifier'
					|| (left.kind === 'keyword' && !OPERAND_STARTING_KEYWORDS.has(tokenText(left)))
					|| left.kind === 'integerLiteral' || left.kind === 'floatLiteral' || left.kind === 'stringLiteral'
					|| left.kind === 'dateLiteral' || left.rawText === ')');
				// An element operand: `v(1)` or `Split("1 b")(1)` (issue #260).
				const rightElement = tokenName(right) !== undefined && toks[i + 2]?.rawText === '('
					? elementOperandStartingAt(toks, i + 1, shapesHere(), optionBase)
					: undefined;
				const rightOperand = (): { span: Span; what: string } | undefined => {
					const what = nonnumericString(right) ?? elementString(spanStart, toks, rightElement);
					return what ? { span: rightElement ? elementSpan(spanStart, toks, rightElement) : at(right!), what } : undefined;
				};
				if (word === 'not' && !leftEndsOperand) {
					// Not binds below the comparisons: `Not s = "y"` is
					// Not (s = "y"), a Boolean (issue #361).
					if (notOperandCompares(toks, i + 1)) {
						continue;
					}
					// `Not (v)` coerces v as well.
					const inner = right?.rawText === '(' && toks[i + 3]?.rawText === ')' ? toks[i + 2] : undefined;
					const innerWhat = inner ? nonnumericString(inner) : undefined;
					const operand = innerWhat ? { span: at(inner!), what: innerWhat } : rightOperand();
					if (operand) {
						report(operand.span, operand.what);
					}
					continue;
				}
				if (!isBinary) {
					continue;
				}
				if (!leftEndsOperand) {
					// Unary `-"abc"` (a leading `+` too).
					if ((tok.rawText === '-' || tok.rawText === '+')) {
						const operand = rightOperand();
						if (operand) {
							report(operand.span, operand.what);
						}
					}
					continue;
				}
				const leftElement = left.rawText === ')' ? elementOperandEndingAt(toks, i - 1, shapesHere(), optionBase) : undefined;
				const alwaysCoerces = ['-', '*', '/', '\\', '^'].includes(tok.rawText) || word === 'mod';
				const leftWhat = nonnumericString(left) ?? elementString(spanStart, toks, leftElement);
				const leftString = leftWhat ? { span: leftElement ? elementSpan(spanStart, toks, leftElement) : at(left), what: leftWhat } : undefined;
				const rightString = rightOperand();
				if (alwaysCoerces) {
					if (leftString) {
						report(leftString.span, leftString.what);
					} else if (rightString) {
						report(rightString.span, rightString.what);
					}
					continue;
				}
				if (tok.rawText !== '+') {
					// A comparison: typed operands only (issue #268).
					if (leftString && !leftElement && typedString(left) && typedNumber(right) && !rightElement) {
						report(leftString.span, leftString.what);
					} else if (rightString && !rightElement && typedString(right) && typedNumber(left) && !leftElement) {
						report(rightString.span, rightString.what);
					} else {
						const fromLeft = unreadableAs(left, right);
						const unreadable = fromLeft ?? unreadableAs(right, left);
						if (unreadable) {
							push('stringArithmeticCoercion', `Operator '${tok.rawText}' compares ${unreadable.what} with ${unreadable.as}, which cannot read it. This will raise Run-time error '13': Type mismatch.`, fromLeft ? at(right!) : at(left));
						}
					}
					continue;
				}
				// `+`: a string against a NUMBER, True and False included:
				// `s + True` with s = "True" raises 13, and `"1" + True` is 0
				// (issue #331, measured in Excel 16.0).
				const truth = (operand: VbaToken | undefined): boolean => tokenText(operand) === 'true' || tokenText(operand) === 'false';
				if (leftString && (numeric(right) || truth(right) || typeof rightElement?.value === 'number')) {
					report(leftString.span, leftString.what);
				} else if (rightString && (numeric(left) || truth(left) || typeof leftElement?.value === 'number')) {
					report(rightString.span, rightString.what);
				} else {
					// `s + d`, `s + b`: a String a Date or Boolean cannot read
					// (issue #331, measured in Excel 16.0).
					const fromLeft = unreadableAs(left, right);
					const unreadable = fromLeft ?? unreadableAs(right, left);
					if (unreadable) {
						push('stringArithmeticCoercion', `Operator '+' adds ${unreadable.what} to ${unreadable.as}, which cannot read it. This will raise Run-time error '13': Type mismatch.`, fromLeft ? at(right!) : at(left));
					}
				}
			}
		}
		visitBlocks(member.body);
		return (stmt) => {
			if (unreachable.has(stmt)) {
				return;
			}
			known = valuesAt(stmt);
			current = stmt;
			shapes = undefined;
			const toks = statementTokens(source, stmt.span);
			const first = firstExecutableTokenIndex(toks);
			const head = tokenText(toks[first]);
			if (head === 'const') {
				return;
			}
			// A one-line If and an ElseIf line are statements; so is IIf.
			if (head === 'if' || head === 'elseif') {
				checkCondition(stmt.span.start, toks, first + 1, conditionEnd(toks, first + 1), head === 'if' ? 'If' : 'ElseIf');
			}
			for (let k = 0; k + 1 < toks.length; k++) {
				if (toks[k].rawText.length === 3 && toks[k + 1].rawText === '(' && tokenText(toks[k]) === 'iif' && isBareOrVbaQualifiedIntrinsicCall(toks, k)) {
					const close = matchParenFrom(toks, k + 1);
					const comma = toks.findIndex((tok, j) => j > k + 1 && tok.rawText === ',');
					if (close > 0 && comma > 0 && comma < close) {
						checkCondition(stmt.span.start, toks, k + 2, comma, 'IIf');
					}
				}
			}
			// A string literal in arithmetic into a numeric variable is the
			// assignment rule's: it already names the target, and one report per
			// line is enough. A local holding the string is this rule's, since
			// the assignment rule reads only literals: `x = s + 1` into a Long
			// with s holding "abc" (issue #180).
			const bare = bareAssignmentTarget(source, stmt.span);
			if (bare) {
				const targetType = env.get(bare.name.toLowerCase());
				if (targetType && nonnumericStringArithmeticOperand(targetType, bare.valueTokens, 0)) {
					return;
				}
			}
			// The assignment's own `=` stores; it compares nothing. So does the
			// `=` of an assignment in a one-line If's branch: `If L > 0 Then
			// tb = s` (issue #342).
			const branchAssigns = new Set<number>();
			for (const branch of stmt.kind === 'Statement' ? stmt.singleLineIfBranches ?? [] : []) {
				if (bareAssignmentTarget(source, branch)) {
					const at = toks.findIndex((tok) => tok.rawText === '=' && stmt.span.start + tok.start >= branch.start);
					if (at >= 0) {
						branchAssigns.add(at);
					}
				}
			}
			scanOperators(stmt.span.start, toks, bare ? toks.findIndex((tok) => tok.rawText === '=') : -1, branchAssigns);
		};
	};
}

/** The operators that convert both operands to numbers and bind loosest. */
const LOGICAL_OPERATORS: ReadonlySet<string> = new Set(['and', 'or', 'xor', 'eqv', 'imp']);

const EMPTY_SHAPES: ReadonlyMap<string, FixedArrayBound> = new Map();

const COMPARISON_OPERATORS: ReadonlySet<string> = new Set(['=', '<', '>', '<=', '>=', '<>']);

/**
 * Whether the operand of a Not starting at `from` holds a comparison: Not
 * takes everything up to the next And, Or, Xor, Eqv or Imp, a closing
 * parenthesis, a comma or Then, and a comparison inside makes it a Boolean
 * (issue #361, measured in Excel 16.0).
 */
function notOperandCompares(toks: readonly VbaToken[], from: number): boolean {
	let depth = 0;
	for (let k = from; k < toks.length; k++) {
		const tok = toks[k];
		const word = tokenText(tok);
		if (tok.rawText === '(') {
			depth++;
		} else if (tok.rawText === ')') {
			if (depth === 0) {
				return false;
			}
			depth--;
		} else if (depth === 0) {
			if (LOGICAL_OPERATORS.has(word) || word === 'then' || tok.rawText === ',' || tok.rawText === ':') {
				return false;
			}
			if ((tok.kind === 'operator' && COMPARISON_OPERATORS.has(tok.rawText)) || word === 'like' || word === 'is') {
				return true;
			}
		}
	}
	return false;
}

/** Keywords after which an operand starts, so a `Not` or a sign there is unary. */
const OPERAND_STARTING_KEYWORDS: ReadonlySet<string> = new Set([
	'if', 'elseif', 'then', 'else', 'while', 'until', 'case', 'to', 'step', 'and', 'or', 'xor', 'eqv', 'imp',
	'not', 'mod', 'like', 'is', 'call', 'set', 'let', 'return',
]);

/** A name a branch has tested non-zero, and the span the test covers. */
interface DivisionGuard {
	name: string;
	start: number;
	end: number;
}

/**
 * Which names an If condition proves non-zero on its Then arm (`SCALE_BY <> 0`,
 * `n > 0`, `Not n = 0`, a bare `n`) and which it proves zero (`n = 0`, so the
 * Else arm has the non-zero case). A constant that fails the test never
 * reaches the division: `If SCALE_BY <> 0 Then x = 10 / SCALE_BY` with
 * SCALE_BY = 0 runs clean (issue #106, measured in Excel 16.0).
 */
function divisionGuardNames(condition: readonly VbaToken[]): { nonZero: Set<string>; zero: Set<string> } {
	const nonZero = new Set<string>();
	const zero = new Set<string>();
	const words = condition.map((tok) => tokenText(tok));
	const nameAt = (index: number): string | undefined => tokenName(condition[index])?.toLowerCase();
	const isZero = (index: number): boolean => condition[index]?.kind === 'integerLiteral' && /^0+$/.test(condition[index].rawText);
	// Conjuncts each hold on the Then arm; a disjunction proves nothing.
	if (words.includes('or')) {
		return { nonZero, zero };
	}
	let start = 0;
	for (let i = 0; i <= words.length; i++) {
		if (i < words.length && words[i] !== 'and') {
			continue;
		}
		const w = words.slice(start, i);
		const n = (k: number): string | undefined => nameAt(start + k);
		const z = (k: number): boolean => isZero(start + k);
		if (w.length === 1 && n(0)) {
			nonZero.add(n(0)!);
		} else if (w.length === 3 && n(0) && z(2) && (w[1] === '<>' || w[1] === '>' || w[1] === '<')) {
			nonZero.add(n(0)!);
		} else if (w.length === 3 && n(2) && z(0) && (w[1] === '<>' || w[1] === '>' || w[1] === '<')) {
			nonZero.add(n(2)!);
		} else if (w.length === 3 && n(0) && z(2) && w[1] === '=') {
			zero.add(n(0)!);
		} else if (w.length === 4 && w[0] === 'not' && n(1) && w[2] === '=' && z(3)) {
			nonZero.add(n(1)!);
		} else if (w.length === 6 && w[0] === 'not' && w[1] === '(' && n(2) && w[3] === '=' && z(4) && w[5] === ')') {
			nonZero.add(n(2)!);
		} else if (w.length === 2 && w[0] === 'not' && n(1)) {
			zero.add(n(1)!);
		}
		start = i + 1;
	}
	return { nonZero, zero };
}

/** The guards every block If in the body establishes for its arms. */
function divisionGuardRanges(
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): DivisionGuard[] {
	const out: DivisionGuard[] = [];
	const visit = (nodes: readonly BodyNode[]): void => {
		for (const node of nodes) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (node.kind === 'IfBlock') {
				node.branches.forEach((branch: IfBranchNode, index: number) => {
					if (!branch.conditionRaw) {
						return;
					}
					const condition = rawExpressionTokens(branch.conditionRaw);
					const names = divisionGuardNames(condition);
					for (const name of names.nonZero) {
						out.push({ name, start: branch.span.start, end: branch.span.end });
					}
					const next = node.branches[index + 1];
					if (next?.branchKind === 'else') {
						for (const name of names.zero) {
							out.push({ name, start: next.span.start, end: next.span.end });
						}
					}
				});
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
			}
		}
	};
	visit(body);
	return out;
}

function divisionByZeroDivisors(
	source: string,
	span: Span,
	constants: IntegerConstantLookup,
	guards: readonly DivisionGuard[],
	fractionOf?: (lower: string) => number | undefined,
	isNull: (tok: VbaToken | undefined) => boolean = () => false,
	pastLong: (tok: VbaToken | undefined) => boolean = () => false,
): Array<{ operator: string; span: Span; message: string }> {
	const toks = statementTokens(source, span);
	const hits: Array<{ operator: string; span: Span; message: string }> = [];
	// A single-line If guards its own arms.
	const first = firstExecutableTokenIndex(toks);
	let thenIndex = -1;
	let elseIndex = -1;
	let local = { nonZero: new Set<string>(), zero: new Set<string>() };
	if (tokenText(toks[first]) === 'if') {
		thenIndex = toks.findIndex((tok, index) => index > first && tokenText(tok) === 'then');
		if (thenIndex > 0) {
			local = divisionGuardNames(toks.slice(first + 1, thenIndex));
			elseIndex = toks.findIndex((tok, index) => index > thenIndex && tokenText(tok) === 'else');
		}
	}
	for (let i = 0; i < toks.length; i++) {
		const operator = divisionByZeroOperatorLabel(toks[i]);
		if (!operator) {
			continue;
		}
		const divisor = zeroDivisorToken(source, span, toks, i + 1, constants)
			?? fractionalDivisorRoundingToZero(toks, i + 1, operator, fractionOf, constants);
		if (!divisor) {
			continue;
		}
		const divisorName = divisor.length === 1 ? tokenName(divisor[0])?.toLowerCase() : undefined;
		if (divisorName) {
			const inElse = elseIndex >= 0 && i > elseIndex;
			const inThen = thenIndex >= 0 && i > thenIndex && !inElse;
			if ((inThen && local.nonZero.has(divisorName)) || (inElse && local.zero.has(divisorName))) {
				continue;
			}
			const at = span.start + toks[i].start;
			if (guards.some((guard) => guard.name === divisorName && at >= guard.start && at < guard.end)) {
				continue;
			}
		}
		// `0 / 0` raises 6 (Overflow), not 11; `\` and `Mod` raise 11 for it
		// (issue #106, measured in Excel 16.0).
		const dividend = toks[i - 1];
		if (isNull(dividend) && toks[i - 2]?.rawText !== '.') {
			continue;
		}
		// `\` and Mod convert the dividend to a Long first: one past the Long
		// range overflows (6) before the division (issue #502).
		if (operator !== '/' && pastLong(dividend) && toks[i - 2]?.rawText !== '.') {
			continue;
		}
		const dividendZero = dividend !== undefined && (
			(dividend.kind === 'integerLiteral' && /^0+[%&^]?$/.test(dividend.rawText))
			|| (dividend.kind === 'floatLiteral' && Number(dividend.rawText.replace(/[!#@]$/, '')) === 0)
			|| (tokenName(dividend) !== undefined && constants.get(tokenName(dividend)!.toLowerCase()) === 0)
		);
		const message = operator === '/' && dividendZero
			? "Expression divides zero by zero with '/'. This will raise Run-time error '6': Overflow."
			: `Expression uses '${operator}' with a zero divisor. This will raise Run-time error '11': Division by zero.`;
		hits.push({ operator, span: absoluteTokenGroupSpan(span, divisor), message });
	}
	return hits;
}

/**
 * `\` and `Mod` round their operands to whole numbers first, with banker's
 * rounding, so a literal divisor below 0.5 - or exactly 0.5 - is zero to them:
 * `5 \ 0.4` and `5 Mod 0.5` raise 11 (issue #119, measured in Excel 16.0).
 * So is a local holding such a value here: `d = 0.4` then `10 Mod d`
 * (issue #239).
 */
function fractionalDivisorRoundingToZero(
	toks: readonly VbaToken[],
	start: number,
	operator: string,
	fractionOf?: (lower: string) => number | undefined,
	constants?: IntegerConstantLookup,
): VbaToken[] | undefined {
	if (operator === '/') {
		return undefined;
	}
	const name = tokenName(toks[start])?.toLowerCase();
	const held = name && fractionOf && isDivisorAtomBoundary(toks[start + 1]) ? fractionOf(name) : undefined;
	if (held !== undefined) {
		return held !== 0 && Math.abs(held) <= 0.5 ? [toks[start]] : undefined;
	}
	// `10 \ Val("0.4")`: a call the folder reads to a fraction (issue #703).
	if (name && toks[start + 1]?.rawText === '(' && toks[start - 1]?.rawText !== '.' && constants) {
		const end = matchParenFrom(toks, start + 1);
		const call = end > start ? toks.slice(start, end + 1) : [];
		const value = call.length > 0 && isDivisorAtomBoundary(toks[end + 1]) ? evaluateIntegerConstantExpression(call.map((tok) => tok.rawText).join(' '), constants) : undefined;
		return value !== undefined && value !== 0 && Math.abs(value) <= 0.5 ? call : undefined;
	}
	let index = start;
	const group: VbaToken[] = [];
	if (toks[index]?.kind === 'operator' && (toks[index].rawText === '-' || toks[index].rawText === '+')) {
		group.push(toks[index]);
		index++;
	}
	const literal = toks[index];
	if (literal?.kind !== 'floatLiteral' || !isDivisorAtomBoundary(toks[index + 1])) {
		return undefined;
	}
	const value = Math.abs(Number(literal.rawText.replace(/[!#@]$/, '').replace(/[dD]/g, 'E')));
	if (!Number.isFinite(value) || value > 0.5) {
		return undefined;
	}
	group.push(literal);
	return group;
}

function divisionByZeroOperatorLabel(tok: VbaToken | undefined): string | undefined {
	const text = tokenText(tok);
	if (text === '/' || text === '\\') {
		return text;
	}
	return text === 'mod' ? 'Mod' : undefined;
}

function zeroDivisorToken(
	source: string,
	span: Span,
	toks: VbaToken[],
	start: number,
	constants: IntegerConstantLookup,
): VbaToken[] | undefined {
	const first = toks[start];
	if (!first) {
		return undefined;
	}
	if (first.rawText === '(') {
		const close = matchParenFrom(toks, start);
		if (close < 0) {
			return undefined;
		}
		return zeroDivisorExpression(source, span, toks, start + 1, close, constants);
	}
	if (
		first.kind === 'operator' &&
		(first.rawText === '+' || first.rawText === '-')
	) {
		const signed = zeroDivisorAtomTokenGroup(toks, start + 1, constants);
		return signed ? [first, ...signed] : undefined;
	}
	return zeroDivisorAtomTokenGroup(toks, start, constants);
}

function zeroDivisorExpression(
	source: string,
	span: Span,
	toks: VbaToken[],
	start: number,
	endExclusive: number,
	constants: IntegerConstantLookup,
): VbaToken[] | undefined {
	if (start >= endExclusive) {
		return undefined;
	}
	const folded = foldIntegerExpressionTokens(source, span, toks, start, endExclusive, constants);
	if (folded === 0) {
		return toks.slice(start, endExclusive);
	}
	// A comparison known False is 0: `5 / (n = 2)` with n = 1 (issue #491).
	const inner = toks.slice(start, endExclusive);
	let depth = 0;
	const compares = inner.some((tok) => {
		depth += tok.rawText === '(' ? 1 : tok.rawText === ')' ? -1 : 0;
		return depth === 0 && ['=', '<>', '<', '>', '<=', '>='].includes(tok.rawText);
	});
	if (compares && conditionValue(inner, { value: (lower) => constants.get(lower) }) === false) {
		return inner;
	}
	if (toks[start]?.rawText === '(') {
		const close = matchParenFrom(toks, start);
		if (close === endExclusive - 1) {
			return zeroDivisorExpression(source, span, toks, start + 1, close, constants);
		}
	}
	if (
		endExclusive === start + 2 &&
		toks[start]?.kind === 'operator' &&
		(toks[start].rawText === '+' || toks[start].rawText === '-') &&
		isZeroDivisorAtom(toks[start + 1], constants)
	) {
		return [toks[start], toks[start + 1]];
	}
	if (endExclusive === start + 1 && isZeroDivisorAtom(toks[start], constants)) {
		return [toks[start]];
	}
	if (zeroConversionCallEnd(toks, start, constants) === endExclusive - 1) {
		return toks.slice(start, endExclusive);
	}
	return undefined;
}

function zeroDivisorAtomTokenGroup(
	toks: readonly VbaToken[],
	start: number,
	constants: IntegerConstantLookup,
): VbaToken[] | undefined {
	const first = toks[start];
	const firstName = first ? tokenName(first) : undefined;
	const member = toks[start + 2];
	const memberName = member ? tokenName(member) : undefined;
	if (firstName && toks[start + 1]?.rawText === '.' && memberName) {
		// Only treat `first.member` as the complete divisor when nothing extends
		// the member-access chain past it; otherwise `a.Zero.Foo` / `a.Zero(i)`
		// would mis-match on the inner `a.Zero == 0` lookup.
		if (!isDivisorAtomBoundary(toks[start + 3])) {
			return undefined;
		}
		return constants.get(`${firstName}.${memberName}`.toLowerCase()) === 0
			? [first, toks[start + 1], member]
			: undefined;
	}
	// A bare atom only stands alone when it is not itself a member-access head or
	// a call target (a following '.' or '(' means more of the expression follows).
	if (isZeroDivisorAtom(first, constants) && isDivisorAtomBoundary(toks[start + 1])) {
		return [first];
	}
	// `F()`, a Function of the module that returns 0 (issue #448).
	if (firstName && toks[start + 1]?.rawText === '(' && toks[start + 2]?.rawText === ')' && isDivisorAtomBoundary(toks[start + 3])
		&& toks[start - 1]?.rawText !== '.' && constants.get(`${firstName}()`) === 0) {
		return toks.slice(start, start + 3);
	}
	const close = zeroConversionCallEnd(toks, start, constants);
	if (close !== undefined && isDivisorAtomBoundary(toks[close + 1])) {
		return toks.slice(start, close + 1);
	}
	// `Int(0.9)`, `Fix(-0.9)`, `Round(0.5)`: a number made whole, 0 (issue
	// #286). `Sign1(-1)`: a Function of the module that returns 0 for these
	// arguments (issue #562).
	if (firstName && toks[start + 1]?.rawText === '(' && toks[start - 1]?.rawText !== '.') {
		const end = matchParenFrom(toks, start + 1);
		const call = end > start ? toks.slice(start, end + 1) : [];
		if (call.length > 0 && isDivisorAtomBoundary(toks[end + 1])
			&& evaluateIntegerConstantExpression(call.map((tok) => tok.rawText).join(' '), constants) === 0) {
			return call;
		}
	}
	return undefined;
}

/** Conversions to a whole-number type, which round their argument half to even. */
const WHOLE_CONVERSIONS: ReadonlySet<string> = new Set(['cbyte', 'cint', 'clng', 'clnglng', 'clngptr']);
const FRACTIONAL_CONVERSIONS: ReadonlySet<string> = new Set(['csng', 'cdbl', 'ccur', 'cdec']);

/**
 * Where a conversion of a literal that comes out 0 ends: `CLng(0)`,
 * `CDbl(0)`, `CLng(0.4)`, which rounds to 0 (issue #219, measured in Excel
 * 16.0; `10 / CDbl(0.4)` runs). A `VBA.` qualifier is allowed.
 */
function zeroConversionCallEnd(toks: readonly VbaToken[], start: number, constants?: IntegerConstantLookup): number | undefined {
	let index = start;
	if (tokenText(toks[index]) === 'vba' && toks[index + 1]?.rawText === '.') {
		index += 2;
	}
	const name = tokenText(toks[index]);
	const whole = WHOLE_CONVERSIONS.has(name);
	if ((!whole && !FRACTIONAL_CONVERSIONS.has(name)) || toks[index + 1]?.rawText !== '(' || toks[start - 1]?.rawText === '.') {
		return undefined;
	}
	const close = matchParenFrom(toks, index + 1);
	const inner = close < 0 ? [] : toks.slice(index + 2, close);
	const signed = inner.length === 2 && (inner[0].rawText === '-' || inner[0].rawText === '+');
	const literal = inner.length === 1 ? inner[0] : signed ? inner[1] : undefined;
	// `CByte(o)` with o a Boolean still False, `CLng(v)` with v Empty (issue #491).
	if (literal && constants && isZeroDivisorAtom(literal, constants)) {
		return close;
	}
	if (!literal || (literal.kind !== 'integerLiteral' && literal.kind !== 'floatLiteral')) {
		return undefined;
	}
	if (isZeroNumericLiteral(literal)) {
		return close;
	}
	const value = Math.abs(Number(literal.rawText.replace(/[!#@%&^]$/, '').replace(/[dD]/g, 'E')));
	return whole && Number.isFinite(value) && value <= 0.5 ? close : undefined;
}

/**
 * True when `tok` terminates a divisor atom: end-of-tokens, a closing paren, a
 * comma, or any operator that is NOT a member-access dot or call-opening paren.
 * Matching another '.' or '(' means the atom continues, so the group is not the
 * whole divisor.
 */
function isDivisorAtomBoundary(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return true;
	}
	return tok.rawText !== '.' && tok.rawText !== '(';
}

function isZeroDivisorAtom(
	tok: VbaToken | undefined,
	constants: IntegerConstantLookup,
): boolean {
	// False is 0 as a number: `1 \ False` raises 11 (issue #458, measured in
	// Excel 16.0).
	if (isZeroNumericLiteral(tok) || (tok?.kind === 'keyword' && (tokenText(tok) === 'false' || tokenText(tok) === 'empty'))) {
		return true;
	}
	const name = tok ? tokenName(tok) : undefined;
	return name !== undefined && constants.get(name.toLowerCase()) === 0;
}

function isZeroNumericLiteral(tok: VbaToken | undefined): boolean {
	if (!tok || (tok.kind !== 'integerLiteral' && tok.kind !== 'floatLiteral')) {
		return false;
	}
	const normalized = tok.rawText
		.replace(/[!#@%&^]$/, '')
		.replace(/[dD]/g, 'E');
	const hex = /^&[hH]([0-9A-Fa-f]+)$/.exec(normalized);
	if (hex) {
		return Number.parseInt(hex[1], 16) === 0;
	}
	const octal = /^&[oO]([0-7]+)$/.exec(normalized);
	if (octal) {
		return Number.parseInt(octal[1], 8) === 0;
	}
	if (!/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalized)) {
		return false;
	}
	return Number(normalized) === 0;
}

function absoluteTokenGroupSpan(base: Span, toks: readonly VbaToken[]): Span {
	return { start: base.start + toks[0].start, end: base.start + toks[toks.length - 1].end };
}

/**
 * Rule: when a Function is used inside an expression, its argument list must be
 * parenthesized (`x = Foo(1, 2)`). The parenless form (`Foo 1, 2`) is only a
 * call-statement form and becomes a VBE syntax error after `=`.
 */
export function checkExpressionCallParens(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): ProcedureStatementVisitor {
	const functions = expressionCallableFunctionNames(symbols, projectProcedures);
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		return (stmt) => {
			const hit = parenlessExpressionCall(source, stmt.span, functions, sourceNames);
			if (hit) {
				push(
					'expressionCallRequiresParens',
					`Function call arguments in an expression must be enclosed in parentheses: use '${hit.name}(...)'.`,
					hit.span,
				);
			}
		};
	};
}

interface ExpressionCallableFunctions {
	bare: Set<string>;
	qualified: Set<string>;
}

function expressionCallableFunctionNames(
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
): ExpressionCallableFunctions {
	const bare = new Set<string>();
	const qualified = new Set<string>();
	for (const member of symbols.root.children ?? []) {
		if (member.kind === 'function' || member.kind === 'propertyGet') {
			bare.add(member.name.toLowerCase());
		}
	}
	for (const [key, candidates] of projectProcedures ?? []) {
		if (candidates.length !== 1 || candidates[0].kind !== 'function') {
			continue;
		}
		if (key.includes('.')) {
			qualified.add(key);
		} else if (!bare.has(key)) {
			bare.add(key);
		}
	}
	return { bare, qualified };
}

function parenlessExpressionCall(
	source: string,
	span: Span,
	functions: ExpressionCallableFunctions,
	sourceNames?: SourceNameScope,
): { name: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	if (toks.length === 0 || isNonAssignmentStatementLeader(tokenText(toks[firstExecutableTokenIndex(toks)]))) {
		return undefined;
	}
	const eq = topLevelOperatorIndex(toks, '=');
	if (eq < 0) {
		return undefined;
	}

	for (let i = eq + 1; i < toks.length - 1; i++) {
		const tok = toks[i];
		const name = tokenName(tok);
		if (!name || !isExpressionCallableAt(toks, i, name, functions, sourceNames)) {
			continue;
		}
		if (i > eq + 1 && toks[i - 1].rawText === '.') {
			const qualifier = tokenName(toks[i - 2]);
			if (!qualifier || !functions.qualified.has(qualifiedProcedureKey(qualifier, name))) {
				continue; // object member calls need receiver typing before we can be precise
			}
		}
		const next = toks[i + 1];
		if (!isParenlessArgumentStart(next)) {
			continue;
		}
		const gap = source.slice(span.start + tok.end, span.start + next.start);
		if (!/\s/.test(gap)) {
			continue;
		}
		return {
			name,
			span: { start: span.start + tok.start, end: span.start + tok.end },
		};
	}
	return undefined;
}

function isExpressionCallableAt(
	toks: readonly VbaToken[],
	index: number,
	name: string,
	functions: ExpressionCallableFunctions,
	sourceNames?: SourceNameScope,
): boolean {
	if (index > 1 && toks[index - 1].rawText === '.') {
		const qualifier = tokenName(toks[index - 2]);
		return qualifier
			? functions.qualified.has(qualifiedProcedureKey(qualifier, name))
			: false;
	}
	if (index > 0 && toks[index - 1].rawText === '.') {
		return false;
	}
	if (bareCallableSourceShadowed(name, sourceNames)) {
		return false;
	}
	if (functions.bare.has(name.toLowerCase())) {
		return true;
	}
	if (runtimeCallableSourceShadowed(name, sourceNames)) {
		return false;
	}
	return resolveRuntimeFunction(name)?.kind === 'function';
}

function isParenlessArgumentStart(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return false;
	}
	switch (tok.kind) {
		case 'identifier':
		case 'bracketedIdentifier':
		case 'integerLiteral':
		case 'floatLiteral':
		case 'stringLiteral':
		case 'dateLiteral':
			return true;
		case 'keyword':
			return !isInfixExpressionKeyword(tok.rawText);
		default:
			return false;
	}
}

function isInfixExpressionKeyword(text: string): boolean {
	switch (text.toLowerCase()) {
		case 'and':
		case 'or':
		case 'xor':
		case 'eqv':
		case 'imp':
		case 'is':
		case 'mod':
			return true;
		default:
			return false;
	}
}

function isNonAssignmentStatementLeader(word: string): boolean {
	switch (word) {
		case 'if':
		case 'elseif':
		case 'for':
		case 'do':
		case 'loop':
		case 'while':
		case 'select':
		case 'case':
			return true;
		default:
			return false;
	}
}

function implicitParenthesizedBareCallableCall(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): { name: string; span: Span } | undefined {
	const call = standaloneEmptyParenthesizedCallStatement(source, span);
	if (!call || call.isMember) {
		return undefined;
	}
	const signature = callableSignatureFor(call.name, moduleSignatures, sourceNames);
	if (!signature || !callableAcceptsZeroArguments(signature)) {
		return undefined;
	}
	return {
		name: call.name,
		span: call.span,
	};
}

// A standalone `mySub2("a", "b", "C")` wraps a multi-argument list in
// parentheses without `Call`: the VBE "Expected: =" compile error. Scoped to a
// callee that resolves to a known procedure so unknown names (which could be
// array indexing or external references) stay silent:
//   - bare names bind to same-module/unique-exported project Sub/Function/Declare;
//   - `Module.Proc(...)` binds to an exported standard-module procedure through
//     its deterministic qualified key (the same resolution the argument-count
//     rule uses).
// Object member/property calls (`obj.Method(a, b)`) are deliberately deferred:
// a non-empty single-argument member form is legal (`ActiveSheet.Range("A1")`)
// and multi-argument member/default-member forms are unproven. The
// single-argument ByVal-grouping form is excluded by the >= 2 argument guard in
// the shared helper.
function implicitParenthesizedMultiArgCall(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): { name: string; span: Span } | undefined {
	const call = standaloneMultiArgParenthesizedCallStatement(source, span);
	if (!call) {
		return undefined;
	}
	if (call.isMember) {
		// `(1, 2)` is no expression, so the statement is a Syntax error
		// whatever the receiver is: `c.Calc (1, 2)` on a class, `.Calc (1, 2)`
		// in a With, `ActiveSheet.Range ("A1", "B2")` (issue #224, measured).
		return {
			name: call.qualifier ? `${call.qualifier}.${call.name}` : call.name,
			span: call.span,
		};
	}
	if (!callableSignatureFor(call.name, moduleSignatures, sourceNames)) {
		return undefined;
	}
	return {
		name: call.name,
		span: call.span,
	};
}

function implicitParenthesizedMemberCall(
	source: string,
	span: Span,
	memberCtx: MemberCompletionContext,
): { name: string; span: Span } | undefined {
	const call = standaloneEmptyParenthesizedCallStatement(source, span);
	if (!call || !call.isMember) {
		return undefined;
	}
	if (
		call.startsWithLeadingDot &&
		!resolveExactMemberCompletion(source, call.name, call.calleeEndOffset, memberCtx)
	) {
		return undefined;
	}
	return { name: call.name, span: call.span };
}
