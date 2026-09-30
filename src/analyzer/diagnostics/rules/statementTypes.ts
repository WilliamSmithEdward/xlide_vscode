// Rule family: statements the VBE refuses by their form or by the type of
// what they name (issue #213); statementForms.ts holds the #125 family.
// Each verdict is a full project compile in 64-bit Excel 16.0 (build 20326,
// 2026-09-30):
//
//  - for-variable-in-use: a For or For Each inside another on the same
//    control variable, "For control variable already in use".
//  - for-counter-type: a For counter that is a String, Boolean, object or
//    user-defined type, "Type mismatch". Variant, Date and the numbers take.
//  - for-each-source-type: For Each over an array of a user-defined type or
//    of fixed-length strings, fixed or dynamic.
//  - statement-before-first-case: a statement or label between Select Case
//    and its first Case; a comment is fine.
//  - return-with-value: `Return 5`, "Syntax error". A bare Return is GoSub's.
//  - with-scalar-target: With on a number, a string or a scalar variable,
//    "With object must be user-defined type, Object, or Variant".
//  - set-requires-object: Set of a user-defined type variable.
//  - paramarray-passing-mode: ByVal or ByRef on a ParamArray, "Expected:
//    identifier".
//  - udt-value-mismatch and udt-variant-coercion: a user-defined type where
//    a single value is needed, and one handed to a Variant.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { VbaToken } from '../../lexer/tokenKinds';
import type { BodyNode, ModuleNode, ProcedureNode, Span } from '../../parser/nodes';
import type { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type { VbaSymbol } from '../../symbols/symbolModel';
import type { ProjectTypeName } from '../../completion/typeCompletion';
import { procedureSymbolFor, type PushFn } from '../analysisContext';
import { isKnownScalarType, normalizeType, sourceIdentifierBinding } from '../typeInference';
import {
	absoluteSpan,
	activeModuleMembers,
	isInactiveNode,
	rawExpressionTokens,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from '../walker';

/** What a variable holds, as far as these rules care. */
type Category = 'variant' | 'number' | 'date' | 'string' | 'boolean' | 'object' | 'udt' | 'class';

const NUMBER_TYPES = new Set(['byte', 'integer', 'long', 'longlong', 'longptr', 'currency', 'single', 'double', 'decimal']);

/** Words that join two operands, so a user-defined type beside one is used as a value. */
const OPERATOR_WORDS = new Set(['and', 'or', 'xor', 'eqv', 'imp', 'mod', 'like', 'not', 'is']);

interface Context {
	source: string;
	symbols: ReturnType<typeof buildModuleSymbols>;
	projectVisibleSymbols: readonly VbaSymbol[] | undefined;
	activity: ConditionalActivityTracker | undefined;
	push: PushFn;
	udts: ReadonlySet<string>;
	enums: ReadonlySet<string>;
	/** Whether an untyped variable is a Variant, which a Deftype line changes. */
	untypedIsVariant: boolean;
}

export function checkStatementTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	projectTypes: readonly ProjectTypeName[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const udts = new Set<string>();
	const enums = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Type') {
			udts.add(member.name.toLowerCase());
		} else if (member.kind === 'Enum') {
			enums.add(member.name.toLowerCase());
		}
	}
	for (const type of projectTypes ?? []) {
		if (type.kind === 'userType') {
			udts.add(type.name.toLowerCase());
		} else if (type.kind === 'enum') {
			enums.add(type.name.toLowerCase());
		}
	}
	const ctx: Context = {
		source,
		symbols,
		projectVisibleSymbols,
		activity,
		push,
		udts,
		enums,
		untypedIsVariant: !/^[ \t]*Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Cur|Sng|Dbl|Dec|Date|Str|Obj|Var)\b/im.test(source),
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure' || member.kind === 'Declare') {
			for (const param of member.params) {
				if (param.paramArray && (param.byVal || param.byRef)) {
					push(
						'paramArrayPassingMode',
						`A ParamArray is always passed ByRef, and takes neither ByVal nor ByRef: '${param.name}'. This is a VBE compile error: Expected: identifier.`,
						param.nameSpan ?? param.span,
					);
				}
			}
		}
		if (member.kind === 'Procedure') {
			const procSym = procedureSymbolFor(symbols, member);
			walkBody(ctx, member, procSym, member.body, []);
		}
	}
}

/** The single variable a bare name binds to, or undefined. */
function variableNamed(ctx: Context, procSym: VbaSymbol | undefined, name: string): VbaSymbol | undefined {
	const binding = sourceIdentifierBinding(ctx.symbols, procSym, ctx.projectVisibleSymbols, name, 'expression');
	if (binding.scope === 'unresolved' || binding.scope === 'ambiguous' || binding.definitions.length !== 1) {
		return undefined;
	}
	const [definition] = binding.definitions;
	return definition.kind === 'localVariable' || definition.kind === 'moduleVariable' || definition.kind === 'parameter'
		? definition
		: undefined;
}

function category(ctx: Context, symbol: VbaSymbol | undefined): Category | undefined {
	if (!symbol || symbol.isArray) {
		return undefined;
	}
	const type = normalizeType(symbol.asType);
	if (!type) {
		return ctx.untypedIsVariant ? 'variant' : undefined;
	}
	if (type === 'variant') {
		return 'variant';
	}
	if (NUMBER_TYPES.has(type) || ctx.enums.has(type)) {
		return 'number';
	}
	if (type === 'date' || type === 'string' || type === 'boolean' || type === 'object') {
		return type;
	}
	if (ctx.udts.has(type)) {
		return 'udt';
	}
	// A name no scalar type or Enum claims is a class: Collection, Range, a
	// class module.
	return isKnownScalarType(type) ? undefined : 'class';
}

function walkBody(
	ctx: Context,
	member: ProcedureNode,
	procSym: VbaSymbol | undefined,
	body: readonly BodyNode[],
	counters: readonly string[],
): void {
	for (const node of body) {
		if (isInactiveNode(ctx.activity, node)) {
			continue;
		}
		let inner = counters;
		switch (node.kind) {
			case 'ForBlock':
				inner = checkFor(ctx, procSym, node, counters);
				break;
			case 'SelectBlock':
				checkSelectHead(ctx, node.body);
				break;
			case 'WithBlock':
				checkWithTarget(ctx, procSym, node.span);
				break;
			case 'IfBlock':
				for (const branch of node.branches) {
					checkCondition(ctx, procSym, branch.conditionRaw, branch.headerSpan);
				}
				break;
			case 'Statement':
			case 'Assignment':
			case 'Call':
				checkStatement(ctx, procSym, node.span);
				break;
			default:
				break;
		}
		if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			walkBody(ctx, member, procSym, (node as { body: BodyNode[] }).body, inner);
		}
	}
}

function checkFor(
	ctx: Context,
	procSym: VbaSymbol | undefined,
	node: Extract<BodyNode, { kind: 'ForBlock' }>,
	counters: readonly string[],
): readonly string[] {
	const name = node.controlVariable?.trim();
	if (!name || !/^[\p{L}_][\p{L}\p{N}_]*$/u.test(name) || !node.controlVariableSpan) {
		return counters;
	}
	const key = name.toLowerCase();
	if (counters.includes(key)) {
		ctx.push(
			'forVariableInUse',
			`'${name}' is already the control variable of an enclosing For. This is a VBE compile error: For control variable already in use.`,
			node.controlVariableSpan,
		);
	}
	const variable = variableNamed(ctx, procSym, name);
	if (!node.each) {
		const kind = category(ctx, variable);
		if (kind === 'string' || kind === 'boolean' || kind === 'object' || kind === 'udt' || kind === 'class') {
			ctx.push(
				'forCounterType',
				`A For counter must be a number, a Date or a Variant, and '${name}' is declared As ${variable!.asType}. This is a VBE compile error: Type mismatch.`,
				node.controlVariableSpan,
			);
		}
	} else if (node.sourceExpression && node.sourceExpressionSpan) {
		const sourceTokens = rawExpressionTokens(node.sourceExpression);
		const sourceName = sourceTokens.length === 1 ? tokenName(sourceTokens[0]) : undefined;
		const array = sourceName ? variableNamed(ctx, procSym, sourceName) : undefined;
		const element = normalizeType(array?.asType);
		if (array?.isArray && (array.fixedLength !== undefined || (element !== undefined && ctx.udts.has(element)))) {
			ctx.push(
				'forEachSourceType',
				`For Each cannot walk '${sourceName}', an array of ${array.fixedLength !== undefined ? 'fixed-length strings' : `the user-defined type ${array.asType}`}. This is a VBE compile error: For Each may not be used on array of user-defined type or fixed-length strings.`,
				node.sourceExpressionSpan,
			);
		}
	}
	return [...counters, key];
}

/** A statement or label between `Select Case` and its first `Case`. */
function checkSelectHead(ctx: Context, body: readonly BodyNode[]): void {
	for (const node of body) {
		if (isInactiveNode(ctx.activity, node) || node.kind === 'ConditionalDirective') {
			continue;
		}
		const toks = statementTokens(ctx.source, node.span).filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline');
		if (toks.length === 0) {
			continue;
		}
		if (tokenText(toks[0]) === 'case') {
			return;
		}
		ctx.push(
			'statementBeforeFirstCase',
			'Nothing but comments may come between Select Case and its first Case. This is a VBE compile error: Statements and labels invalid between Select Case and first Case.',
			{ start: node.span.start + toks[0].start, end: node.span.start + toks[toks.length - 1].end },
		);
		return;
	}
}

/** The `With` line's tokens, up to the end of that physical line. */
function headerTokens(source: string, span: Span): VbaToken[] {
	const lineEnd = source.slice(span.start, span.end).search(/\r|\n/);
	const toks = statementTokens(source, lineEnd < 0 ? span : { start: span.start, end: span.start + lineEnd });
	const end = toks.findIndex((tok) => tok.kind === 'comment' || tok.kind === 'colon');
	return end < 0 ? toks : toks.slice(0, end);
}

function checkWithTarget(ctx: Context, procSym: VbaSymbol | undefined, span: Span): void {
	const toks = headerTokens(ctx.source, span);
	if (tokenText(toks[0]) !== 'with' || toks.length !== 2) {
		return;
	}
	const target = toks[1];
	const literal = target.kind === 'stringLiteral' || target.kind === 'integerLiteral' || target.kind === 'floatLiteral' || target.kind === 'dateLiteral';
	const name = literal ? undefined : tokenName(target);
	const kind = name ? category(ctx, variableNamed(ctx, procSym, name)) : undefined;
	if (literal || kind === 'number' || kind === 'string' || kind === 'boolean' || kind === 'date') {
		ctx.push(
			'withScalarTarget',
			`With needs an object, a user-defined type or a Variant, and '${target.rawText}' is ${literal ? 'a literal' : 'a scalar'}. This is a VBE compile error: With object must be user-defined type, Object, or Variant.`,
			absoluteSpan(span, target),
		);
	}
}

function checkCondition(ctx: Context, procSym: VbaSymbol | undefined, conditionRaw: string | undefined, headerSpan: Span): void {
	const toks = conditionRaw ? rawExpressionTokens(conditionRaw) : [];
	const name = toks.length === 1 ? tokenName(toks[0]) : undefined;
	if (name && category(ctx, variableNamed(ctx, procSym, name)) === 'udt') {
		const at = ctx.source.indexOf(conditionRaw!.trim(), headerSpan.start);
		udtMismatch(ctx, name, at < 0 ? headerSpan : { start: at, end: at + name.length });
	}
}

function udtMismatch(ctx: Context, name: string, span: Span): void {
	ctx.push(
		'udtValueMismatch',
		`'${name}' is a user-defined type, which cannot be used as a single value here. This is a VBE compile error: Type mismatch.`,
		span,
	);
}

function udtCoercion(ctx: Context, name: string, span: Span): void {
	ctx.push(
		'udtVariantCoercion',
		`'${name}' is a user-defined type, and a Variant cannot hold one declared in a standard module or a private class. This is a VBE compile error: Only user-defined types defined in public object modules can be coerced to or from a variant or passed to late-bound functions.`,
		span,
	);
}

function checkStatement(ctx: Context, procSym: VbaSymbol | undefined, span: Span): void {
	const toks = statementTokensAfterLeadingLabel(ctx.source, span).filter((tok) => tok.kind !== 'comment');
	if (toks.length === 0) {
		return;
	}
	const first = tokenText(toks[0]);
	if (first === 'return' && toks.length > 1 && toks[1].kind !== 'colon') {
		ctx.push(
			'returnWithValue',
			'Return takes no value in VBA: it goes back from a GoSub. Assign the function name instead. This is a VBE compile error: Syntax error.',
			{ start: span.start + toks[0].start, end: span.start + toks[toks.length - 1].end },
		);
		return;
	}
	const udtAt = (i: number): string | undefined => {
		const tok = toks[i];
		const name = tok ? tokenName(tok) : undefined;
		if (!name || toks[i - 1]?.rawText === '.' || toks[i + 1]?.rawText === '.' || toks[i + 1]?.rawText === '(') {
			return undefined;
		}
		return category(ctx, variableNamed(ctx, procSym, name)) === 'udt' ? name : undefined;
	};
	const at = (i: number): Span => absoluteSpan(span, toks[i]);

	if (first === 'set') {
		const name = udtAt(1);
		if (name && toks[2]?.rawText === '=') {
			ctx.push(
				'setRequiresObject',
				`'${name}' is a user-defined type, not an object, so Set cannot assign it. This is a VBE compile error: Object required.`,
				at(1),
			);
		}
		return;
	}
	if (first === 'debug' && toks[1]?.rawText === '.' && tokenText(toks[2]) === 'print') {
		for (let i = 3; i < toks.length; i++) {
			const name = udtAt(i);
			if (name && isArgumentBoundary(toks[i - 1]) && isArgumentBoundary(toks[i + 1])) {
				udtMismatch(ctx, name, at(i));
			}
		}
	}
	if (first === 'msgbox') {
		const start = toks[1]?.rawText === '(' ? 2 : 1;
		const name = udtAt(start);
		if (name && isArgumentBoundary(toks[start + 1])) {
			udtCoercion(ctx, name, at(start));
		}
		return;
	}

	// `a = b`: the target and the value's kinds, when each is a bare name or a
	// literal.
	const eq = first === 'let' ? 2 : 1;
	if (toks[eq]?.rawText === '=' && toks.length === eq + 2 && tokenName(toks[eq - 1])) {
		const targetName = tokenName(toks[eq - 1])!;
		const target = variableNamed(ctx, procSym, targetName);
		const targetKind = category(ctx, target);
		const value = toks[eq + 1];
		const valueName = tokenName(value);
		const valueSymbol = valueName ? variableNamed(ctx, procSym, valueName) : undefined;
		const valueKind = category(ctx, valueSymbol);
		const literal = value.kind === 'stringLiteral' || value.kind === 'integerLiteral' || value.kind === 'floatLiteral' || value.kind === 'dateLiteral';
		if (targetKind === 'udt') {
			const otherType = valueKind === 'udt' && normalizeType(valueSymbol!.asType) !== normalizeType(target!.asType);
			if (literal || otherType || valueKind === 'number' || valueKind === 'string' || valueKind === 'boolean' || valueKind === 'date') {
				udtMismatch(ctx, targetName, at(eq - 1));
			}
		} else if (valueKind === 'udt') {
			if (targetKind === 'variant') {
				udtCoercion(ctx, valueName!, at(eq + 1));
			} else if (targetKind === 'number' || targetKind === 'string' || targetKind === 'boolean' || targetKind === 'date') {
				udtMismatch(ctx, valueName!, at(eq + 1));
			}
		}
		return;
	}

	// A user-defined type beside an operator is used as a value: `t + 1`,
	// `t = u` in an expression. The assignment's own `=` is not one, however
	// long its target: `parents(depth) = cInfo`, `.Items(0).Id = rec`.
	const assignmentEq = startsAssignment(toks) ? firstTopLevelEquals(toks) : -1;
	for (let i = 0; i < toks.length; i++) {
		const name = udtAt(i);
		if (!name) {
			continue;
		}
		const before = toks[i - 1];
		const after = toks[i + 1];
		const operatorBefore = before && i - 1 !== assignmentEq && isOperator(before);
		const operatorAfter = after && i + 1 !== assignmentEq && isOperator(after);
		if (operatorBefore || operatorAfter) {
			udtMismatch(ctx, name, at(i));
		}
	}
}

/**
 * Whether a statement is an assignment by its first token: a name, a
 * `.member`, Let, Set, or LSet and RSet, which copy one user-defined type
 * into another.
 */
function startsAssignment(toks: readonly VbaToken[]): boolean {
	const first = toks[0];
	if (!first) {
		return false;
	}
	const word = tokenText(first);
	return first.rawText === '.' || ['let', 'set', 'lset', 'rset'].includes(word)
		|| first.kind === 'identifier' || first.kind === 'bracketedIdentifier';
}

function firstTopLevelEquals(toks: readonly VbaToken[]): number {
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		depth += raw === '(' ? 1 : raw === ')' ? -1 : 0;
		if (depth === 0 && raw === '=') {
			return i;
		}
	}
	return -1;
}

function isOperator(tok: VbaToken): boolean {
	return tok.kind === 'operator' ? tok.rawText !== '.' && tok.rawText !== '!' : OPERATOR_WORDS.has(tokenText(tok));
}

/** Whether a token ends or begins a Print or MsgBox argument. */
function isArgumentBoundary(tok: VbaToken | undefined): boolean {
	return !tok || tok.rawText === ',' || tok.rawText === ';' || tok.rawText === ')' || tok.rawText === '(' || tok.kind === 'colon'
		|| tokenText(tok) === 'print';
}
