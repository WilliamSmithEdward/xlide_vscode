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
import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import {
	type IntegerConstantLookup,
	resolveRawIntegerConstants,
} from '../../constants/integerConstantExpression';
import { tokenizeCached } from '../../lexer/tokenize';
import type { VbaToken } from '../../lexer/tokenKinds';
import type {
	ModuleNode,
	Span,
} from '../../parser/nodes';
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
	collectBodyLiteralIntegerConstants,
	collectModuleLiteralIntegerConstants,
	foldIntegerExpressionTokens,
} from '../constExpr';
import {
	bareCallableSourceShadowed,
	callableSignatureFor,
	callableTypeSignaturesFor,
	declaredTypeForSourceBinding,
	isKnownScalarType,
	normalizeType,
	resolveExactMemberCompletion,
	runtimeCallableSourceShadowed,
	scopedIntegerConstantLookup,
	type SourceDeclaredTypeResolver,
	type SourceNameScope,
	sourceNameScopeFor,
	typeEnvironmentFor,
} from '../typeInference';
import {
	absoluteSpan,
	firstExecutableTokenIndex,
	matchParenFrom,
	statementTokens,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
	type ProcedureStatementVisitor,
} from '../walker';

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
export function checkUnbalancedParens(source: string, push: PushFn): void {
	const toks = tokenizeCached(source);
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
					`Invalid operator sequence '${hit.text}'; this will fail to compile as a syntax error.`,
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

function invalidOperatorSequence(
	source: string,
	span: Span,
): { text: string; span: Span } | undefined {
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
		if (!isNonUnaryBinaryOperator(toks[i])) {
			continue;
		}
		let end = i;
		while (isNonUnaryBinaryOperator(toks[end + 1])) {
			end++;
		}
		if (end > i) {
			const first = toks[i];
			const last = toks[end];
			return {
				text: source.slice(span.start + first.start, span.start + last.end),
				span: { start: span.start + first.start, end: span.start + last.end },
			};
		}
		if (i === toks.length - 1) {
			return {
				text: toks[i].rawText,
				span: absoluteSpan(span, toks[i]),
			};
		}
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
	if (toks.length === 0 || isNonAssignmentStatementLeader(tokenText(toks[0]))) {
		return undefined;
	}
	const eq = topLevelOperatorIndex(toks, '=');
	if (eq < 0) {
		return undefined;
	}
	let depth = 0;
	for (let i = eq + 1; i + 1 < toks.length; i++) {
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
		const next = toks[i + 1];
		if (endsJuxtaposableValue(toks[i]) && isJuxtaposableValueStart(next)) {
			return { text: next.rawText, span: absoluteSpan(span, next) };
		}
	}
	return undefined;
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
): ProcedureStatementVisitor {
	const projectConstants = resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map());
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, projectConstants);
	return (member) => {
		const procedureConstants = new Map(moduleConstants);
		collectBodyLiteralIntegerConstants(member.body, procedureConstants, activity);
		const procSym = procedureSymbolFor(symbols, member);
		const constants = scopedIntegerConstantLookup(
			procedureConstants,
			symbols,
			procSym,
			projectVisibleSymbols,
		);
		return (stmt) => {
			for (const hit of divisionByZeroDivisors(source, stmt.span, constants)) {
				push(
					'divisionByZero',
					`Expression uses '${hit.operator}' with a zero divisor. This will raise Run-time error '11': Division by zero.`,
					hit.span,
				);
			}
		};
	};
}

function divisionByZeroDivisors(
	source: string,
	span: Span,
	constants: IntegerConstantLookup,
): Array<{ operator: string; span: Span }> {
	const toks = statementTokens(source, span);
	const hits: Array<{ operator: string; span: Span }> = [];
	for (let i = 0; i < toks.length; i++) {
		const operator = divisionByZeroOperatorLabel(toks[i]);
		if (!operator) {
			continue;
		}
		const divisor = zeroDivisorToken(source, span, toks, i + 1, constants);
		if (divisor) {
			hits.push({ operator, span: absoluteTokenGroupSpan(span, divisor) });
		}
	}
	return hits;
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
	return undefined;
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
	if (isZeroNumericLiteral(tok)) {
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
	if (toks.length === 0 || isNonAssignmentStatementLeader(tokenText(toks[0]))) {
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
		if (!call.qualifier || !moduleSignatures.has(qualifiedProcedureKey(call.qualifier, call.name))) {
			return undefined;
		}
		return {
			name: `${call.qualifier}.${call.name}`,
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
