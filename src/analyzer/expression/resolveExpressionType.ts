// The declared type of the expression covering a span.
//
// Extract Variable needs two facts a refactoring cannot derive on its own:
// what type a selected expression has, and whether assigning it needs `Set`.
// VBA assigns an object with `Set` and a value without, and there is no form
// that works for both, so getting the second wrong emits a line that does not
// compile.
//
// Both facts live in the binder already. `resolveHover` answers for the
// IDENTIFIER at an offset and returns a display string; a consumer parsing the
// type back out of prose the analyzer is free to reword has no contract.
// `typeInference` has the real machinery, but reached through a binding rather
// than a span. This composes them for a span
// (github.com/WilliamSmithEdward/xlide_vscode/issues/61).

import type { MemberCompletionContext } from '../completion/memberAccess';
import type { HostObjectModel } from '../host/excelObjectModel';
import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { parseExpression } from '../parser/parseExpression';
import type { ExprNode, ProcedureNode, Span } from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type { ModuleSymbolKind, VbaProjectClassMembers, VbaSymbol } from '../symbols/symbolModel';
import { procedureSymbolFor } from '../diagnostics/analysisContext';
import { resolveRuntimeObject } from '../runtime/vbaRuntime';
import {
	buildModuleTypeSignatures,
	declaredValueTypeForQualifiedSourceBinding,
	declaredValueTypeForSourceBinding,
	inferExpressionType,
	isKnownObjectAssignmentType,
	isKnownScalarType,
	normalizeType,
	sourceNameScopeFor,
	typeEnvironmentFor,
} from '../diagnostics/typeInference';

export interface ExpressionTypeContext {
	moduleName?: string;
	moduleKind?: ModuleSymbolKind;
	/** Host object model to resolve members against. Defaults to the caller's. */
	model?: HostObjectModel;
	/** Source-declared project object members and UDT fields, keyed by type. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/** Module-level symbols visible from other modules in the project. */
	projectVisibleSymbols?: readonly VbaSymbol[];
	/** Member-resolution context, when the caller already has one built. */
	memberContext?: MemberCompletionContext;
}

export interface ExpressionTypeInfo {
	/** `Long`, `Collection`, `Excel.Range`, `Variant`. */
	type: string;
	/** Whether assigning this expression needs `Set`. */
	isObject: boolean;
	/**
	 * Whether the span covers a WHOLE expression. A selection of `1 +` is not
	 * one, and a refactoring must decline rather than emit something that does
	 * not parse. Asked separately because a consumer deciding it outside would
	 * be a second implementation of VBA's expression grammar.
	 */
	complete: boolean;
}


// Binding a module is the expensive half of answering for a span, and a caller
// asking about several spans in one module would otherwise pay it per span -
// which is quadratic over a module's assignments
// (github.com/WilliamSmithEdward/xlide_vscode/issues/62). Value-keyed like
// `tokenizeCached`, and sized the same, so a project pass touching sibling
// modules does not evict the active one. Callers must not mutate what they get.
const BOUND_MODULE_CACHE_MAX = 8;
interface BoundModule {
	source: string;
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	module: ReturnType<typeof parseModule>;
	symbols: ReturnType<typeof buildModuleSymbols>;
}
const boundModuleCache: BoundModule[] = [];

function boundModule(
	source: string,
	moduleName: string,
	moduleKind: ModuleSymbolKind,
): BoundModule {
	for (let i = 0; i < boundModuleCache.length; i += 1) {
		const hit = boundModuleCache[i];
		if (hit.source === source && hit.moduleName === moduleName && hit.moduleKind === moduleKind) {
			// Adopt the caller's instance so later lookups settle on the pointer
			// rather than comparing the whole module again.
			hit.source = source;
			if (i > 0) {
				boundModuleCache.splice(i, 1);
				boundModuleCache.unshift(hit);
			}
			return hit;
		}
	}
	const module = parseModule(source);
	const entry: BoundModule = {
		source,
		moduleName,
		moduleKind,
		module,
		symbols: buildModuleSymbols(moduleName, moduleKind, source, { parsedModule: module }),
	};
	boundModuleCache.unshift(entry);
	if (boundModuleCache.length > BOUND_MODULE_CACHE_MAX) {
		boundModuleCache.pop();
	}
	return entry;
}

/** Tokens of `span`, with comments and line breaks dropped. */
function expressionTokens(source: string, span: Span): VbaToken[] {
	return tokenize(source.slice(span.start, span.end))
		.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
}

/** The procedure whose body contains `span`, if any. */
function enclosingProcedure(
	members: readonly { kind: string; span: Span }[],
	span: Span,
): ProcedureNode | undefined {
	for (const member of members) {
		if (member.kind === 'Procedure' && member.span.start <= span.start && span.end <= member.span.end) {
			return member as ProcedureNode;
		}
	}
	return undefined;
}

/**
 * Whether the expression is a whole-number literal, through any parentheses and
 * a leading sign.
 *
 * VBA types one as Integer or Long, never Double. The shared inference widens
 * every numeric literal to Double, which is right for CHECKING compatibility -
 * a Double accepts any of them - and wrong for a caller about to write the type
 * into a `Dim` (github.com/WilliamSmithEdward/xlide_vscode/issues/64). Asked of
 * the parse rather than the source text, so `10&`, `&H10` and `(10)` all answer
 * the same.
 */
function isWholeNumberLiteral(expr: ExprNode | null | undefined): boolean {
	if (!expr) {
		return false;
	}
	if (expr.exprKind === 'ParenExpr') {
		return isWholeNumberLiteral(expr.inner);
	}
	if (expr.exprKind === 'UnaryExpr' && (expr.operator === '-' || expr.operator === '+')) {
		return isWholeNumberLiteral(expr.operand);
	}
	return expr.exprKind === 'LiteralExpr' && expr.literalKind === 'integer';
}

/**
 * The declared type of the expression `span` covers, or undefined when the span
 * holds no expression at all.
 *
 * `Variant` is a real answer, not a failure: it is what VBA gives a name
 * nothing narrows, and what the developer would have typed themselves.
 */
export function resolveExpressionType(
	source: string,
	span: Span,
	ctx: ExpressionTypeContext = {},
): ExpressionTypeInfo | undefined {
	const tokens = expressionTokens(source, span);
	if (tokens.length === 0) {
		return undefined;
	}
	const parsed = parseExpression(tokens);
	const complete = !!parsed.expr
		&& parsed.diagnostics.length === 0
		&& parsed.endIndex === tokens.length;

	const { module, symbols } = boundModule(
		source,
		ctx.moduleName ?? 'Module',
		ctx.moduleKind ?? 'standard',
	);
	const proc = enclosingProcedure(module.members, span);
	const procSym = proc ? procedureSymbolFor(symbols, proc) : undefined;
	const env = proc ? typeEnvironmentFor(symbols, proc) : new Map<string, string>();
	const sourceNames = proc
		? sourceNameScopeFor(symbols, proc, ctx.projectVisibleSymbols)
		: undefined;
	const memberCtx: MemberCompletionContext = ctx.memberContext ?? {
		model: ctx.model,
		projectClassMembers: ctx.projectClassMembers,
	};

	const inferred = inferExpressionType(
		tokens,
		span.start,
		env,
		buildModuleTypeSignatures(symbols),
		sourceNames,
		source,
		memberCtx,
		(name) => declaredValueTypeForSourceBinding(symbols, procSym, ctx.projectVisibleSymbols, name),
		(qualifier, name) => declaredValueTypeForQualifiedSourceBinding(
			symbols,
			ctx.projectVisibleSymbols,
			qualifier,
			name,
		),
	);

	// Nothing narrowed it. Variant is the honest answer and the compiling one.
	const type = isWholeNumberLiteral(parsed.expr)
		? 'Long'
		: inferred?.type?.trim() || 'Variant';
	// `New T` yields a reference whatever T is, so the expression settles this
	// even when the type table does not carry the class.
	const isObject = isNewExpression(tokens) || needsSetAssignment(type, memberCtx);
	return { type, isObject, complete };
}

/** Whether the expression is `New SomeClass`, which is always a reference. */
function isNewExpression(tokens: readonly VbaToken[]): boolean {
	return tokens[0]?.rawText.toLowerCase() === 'new';
}

/**
 * VBA's own creatable class. It belongs to no host model and to no project, so
 * the assignment table cannot reach it, yet `Dim c As Collection` genuinely
 * needs `Set`. Kept here rather than in the shared table because widening that
 * table changes what `set-required` reports on existing code.
 */
const VBA_INTRINSIC_CLASSES = new Set(['collection']);

/**
 * Whether assigning a value of this type needs `Set`.
 *
 * Object types do; scalars, `Variant` and user-defined `Type`s do not - a UDT
 * is assigned with a plain `=` like a value, which is why "not a scalar" alone
 * is the wrong test. An unrecognised name answers false, so a refactoring that
 * cannot tell emits the plain assignment rather than a `Set` that will not
 * compile against a value.
 */
function needsSetAssignment(type: string, memberCtx: MemberCompletionContext): boolean {
	const normalized = normalizeType(type);
	if (!normalized || normalized === 'variant' || isKnownScalarType(normalized)) {
		return false;
	}
	if (isKnownObjectAssignmentType(type, memberCtx)) {
		return true;
	}
	return VBA_INTRINSIC_CLASSES.has(normalized) || resolveRuntimeObject(normalized) !== undefined;
}
