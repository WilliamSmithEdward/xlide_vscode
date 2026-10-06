// Expression type inference and callable signature tables for diagnostics.
//
// Extracted verbatim from `analyzeModule.ts`: the per-module/project callable
// signature tables, procedure type environments and source-name scopes, the
// expression/atomic-expression type-inference engine, runtime signature
// parsing, expression-level call extraction (expressionCalls and the member
// call binders), argument-type validation, and the type-compatibility tables.
// Pure analysis: the only diagnostics emitted here flow through the PushFn
// passed by the argument-type rules.

import { untouchedModuleVariablesIn, writtenNamesIn } from './moduleState';
import type { VbaToken } from '../lexer/tokenKinds';
import type { HostMember, HostObjectModel } from '../host/excelObjectModel';
import { IDENT_RE, matchParenFrom, splitTopLevelTokenGroups } from '../lexer/tokenHelpers';
import { parseModule } from '../parser/parseModule';
import { HOST_DEFAULT_MEMBERS } from '../host/hostDefaultMembers';
import {
	bankersRound,
	evaluateIntegerConstantExpression,
	parseDecimalIntegerLiteral,
	parseVbaIntegerLiteral,
	type IntegerConstantLookup,
} from '../constants/integerConstantExpression';
import {
	getHostMembers,
	getHostType,
	resolveHostAlias,
	resolveHostConstant,
	resolveHostGlobal,
	resolveHostGlobalMember,
} from '../host/hostModel';
import { hostTypeResolvesWhenCompiling } from '../host/typeExtensibility';
import {
	resolveRuntimeConstant,
	resolveRuntimeFunction,
	resolveRuntimeObject,
	type VbaRuntimeFunction,
} from '../runtime/vbaRuntime';
import { standaloneEmptyParenthesizedCallStatement } from '../call/callContext';
import type { BodyNode, IfBlockNode, ProcedureNode, Span } from '../parser/nodes';
import { isLeafStatement } from '../parser/nodes';
import type { ConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
	VbaProjectClassMembers,
	VbaSymbol,
} from '../symbols/symbolModel';
import {
	isBareCallableKind,
	isProcedureKind,
	procedureParamsFromSymbol,
	qualifiedProcedureKey,
} from '../symbols/symbolModel';
import {
	resolveBareIdentifierBinding,
	sourceIdentifierNames,
	type BareIdentifierContext,
	type BareIdentifierResolution,
} from '../symbols/nameResolution';
import {
	resolveMemberCompletionNamed,
	type MemberCompletion,
	type MemberCompletionContext,
	isExplicitElementAccessor,
	memberTakesOwnArguments,
} from '../completion/memberAccess';
import { procedureSymbolFor, type PushFn } from './analysisContext';
import { EMPTY_COLLECTION, identityAssignment, OBJECT_NOTHING, setCallEffects, VARIANT_EMPTY, setDeclaredFacts, straightLineAssignments, straightLineDeadBranches, straightLineExit, straightLineUnreachable, type CallEffects, type DeclaredFacts, type ReachingAssignments } from './straightLineValues';
import { isInvalidBooleanString, isInvalidDateString, isInvalidNumericString, numericStringVerdict } from './stringConversion';
import {
	callableAcceptsZeroArguments,
	emptyArgSplit,
	isNamedSlot,
	splitArgSlots,
	type CallArguments,
	type CallableParamType,
	type CallableTypeSignature,
	type InferredArgumentType,
} from './callExtraction';
import {
	collectBodyLiteralIntegerConstants,
	collectModuleLiteralIntegerConstants,
	externalIntegerConstantValue,
	numericExternalConstantValue,
} from './constExpr';
import { operatorYieldsNull } from './nullOperators';
import {
	bareAssignmentTarget,
	blockFooterLineSpan,
	blockHeaderLineSpan,
	firstExecutableTokenIndex,
	rawExpressionTokens,
	statementAndBranchSpans,
	statementTokens,
	statementTokensAfterLeadingLabel,
	stripHeaderBrackets,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
} from './walker';

export function resolveExactMemberCompletion(
	source: string,
	memberName: string,
	memberEndOffset: number,
	memberCtx: MemberCompletionContext,
): MemberCompletion | undefined {
	return resolveMemberCompletionNamed(source, memberEndOffset, memberName, memberCtx);
}
// Signature tables and per-procedure environments used to be rebuilt by every
// rule (7+ call sites each iterating all module symbols plus all project
// procedures; audit #5). They are pure functions of the per-pass
// buildModuleSymbols result (plus the per-pass projectProcedures /
// projectVisibleSymbols identities), so memoize them per pass with value-keyed
// WeakMaps, following the procedureSymbolFor precedent. Results are shared:
// callers must not mutate the returned maps (none do - the engine treats all
// derived tables as read-only).
const MODULE_TYPE_SIGNATURES = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	Map<string, CallableTypeSignature>
>();
const SAME_MODULE_CALLABLE_SIGNATURES = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	Map<string, CallableTypeSignature[]>
>();
const CALLABLE_TYPE_SIGNATURES = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	{
		projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined;
		result: Map<string, CallableTypeSignature>;
	}
>();
const UNIQUE_PROJECT_TYPE_SIGNATURES = new WeakMap<
	ReadonlyMap<string, readonly VbaProcedureSignature[]>,
	Map<string, CallableTypeSignature>
>();
const EMPTY_PROJECT_TYPE_SIGNATURES = new Map<string, CallableTypeSignature>();

export function buildModuleTypeSignatures(
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, CallableTypeSignature> {
	const cached = MODULE_TYPE_SIGNATURES.get(symbols);
	if (cached) {
		return cached;
	}
	const out = new Map<string, CallableTypeSignature>();
	for (const symbol of symbols.root.children ?? []) {
		if (isProcedureKind(symbol.kind) || symbol.kind === 'declare') {
			out.set(symbol.name.toLowerCase(), moduleSignatureFromSymbol(symbols, symbol));
		}
	}
	MODULE_TYPE_SIGNATURES.set(symbols, out);
	return out;
}

/** A procedure's signature, its untyped parameters and result typed by the module's DefType lines (issue #285). */
function moduleSignatureFromSymbol(symbols: ReturnType<typeof buildModuleSymbols>, symbol: VbaSymbol): CallableTypeSignature {
	const declared = callableTypeSignatureFromSymbol(symbol);
	if (!symbols.defTypes || symbol.kind === 'declare') {
		return declared;
	}
	return {
		...declared,
		params: declared.params.map((p) => (p.type || p.paramArray ? p : { ...p, type: defTypeOf(symbols, p.name) })),
		returnType: declared.returnType ?? (symbol.kind === 'function' || symbol.kind === 'propertyGet' ? defTypeOf(symbols, symbol.name) : undefined),
	};
}

export function sameModuleCallableSignatures(
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, CallableTypeSignature[]> {
	const cached = SAME_MODULE_CALLABLE_SIGNATURES.get(symbols);
	if (cached) {
		return cached;
	}
	const out = new Map<string, CallableTypeSignature[]>();
	for (const symbol of symbols.root.children ?? []) {
		if (!isBareCallableKind(symbol.kind)) {
			continue;
		}
		const signature = moduleSignatureFromSymbol(symbols, symbol);
		const key = signature.name.toLowerCase();
		const arr = out.get(key);
		if (arr) {
			arr.push(signature);
		} else {
			out.set(key, [signature]);
		}
	}
	SAME_MODULE_CALLABLE_SIGNATURES.set(symbols, out);
	return out;
}

export function callableTypeSignatureFromSymbol(symbol: VbaSymbol): CallableTypeSignature {
	return {
		name: symbol.name,
		params: procedureParamsFromSymbol(symbol, { includePassing: true }).map((p) => ({
			name: stripHeaderBrackets(p.name),
			type: p.type,
			optional: p.optional,
			paramArray: p.paramArray,
			isArray: p.isArray,
			byRef: isByRefProcedureParam(p),
		})),
		returnType: symbol.asType,
		...(symbol.kind === 'function' || symbol.kind === 'propertyGet' ? { valued: true } : {}),
	};
}

export function callableTypeSignaturesFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
): Map<string, CallableTypeSignature> {
	const cached = CALLABLE_TYPE_SIGNATURES.get(symbols);
	if (cached && cached.projectProcedures === projectProcedures) {
		return cached.result;
	}
	// Copy: buildModuleTypeSignatures' result is memoized and must stay pure.
	const out = new Map(buildModuleTypeSignatures(symbols));
	for (const [lower, sig] of uniqueProjectTypeSignatures(projectProcedures)) {
		if (!out.has(lower)) {
			out.set(lower, sig);
		}
	}
	CALLABLE_TYPE_SIGNATURES.set(symbols, { projectProcedures, result: out });
	return out;
}

export function uniqueProjectTypeSignatures(
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
): Map<string, CallableTypeSignature> {
	if (!projectProcedures) {
		return EMPTY_PROJECT_TYPE_SIGNATURES;
	}
	const cached = UNIQUE_PROJECT_TYPE_SIGNATURES.get(projectProcedures);
	if (cached) {
		return cached;
	}
	const out = new Map<string, CallableTypeSignature>();
	for (const [lower, candidates] of projectProcedures) {
		if (candidates.length !== 1) {
			continue;
		}
		const candidate = candidates[0];
		out.set(lower, {
			name: candidate.name,
			params: candidate.params.map((p) => ({
				name: p.name,
				type: p.type,
				optional: p.optional,
				paramArray: p.paramArray,
				isArray: p.isArray,
				byRef: isByRefProcedureParam(p),
			})),
			returnType: candidate.returnType,
			...(candidate.kind === 'function' ? { valued: true } : {}),
		});
	}
	UNIQUE_PROJECT_TYPE_SIGNATURES.set(projectProcedures, out);
	return out;
}

export function isByRefProcedureParam(param: { byRef?: boolean; byVal?: boolean; paramArray?: boolean }): boolean {
	if (param.paramArray) {
		return false;
	}
	return param.byRef === true || param.byVal !== true;
}

// Per-procedure environments, memoized per pass (audit #5): every type rule
// used to rebuild these for each procedure it visited.
const TYPE_ENVIRONMENTS = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	WeakMap<ProcedureNode, ReadonlyMap<string, string>>
>();
const DECLARATION_SHAPE_ENVIRONMENTS = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	WeakMap<ProcedureNode, ReadonlyMap<string, DeclaredValueShape>>
>();
const SOURCE_NAME_SCOPES = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	WeakMap<
		ProcedureNode,
		{ projectVisibleSymbols: readonly VbaSymbol[] | undefined; result: SourceNameScope }
	>
>();

function perProcedureCache<V>(
	store: WeakMap<ReturnType<typeof buildModuleSymbols>, WeakMap<ProcedureNode, V>>,
	symbols: ReturnType<typeof buildModuleSymbols>,
): WeakMap<ProcedureNode, V> {
	let byProc = store.get(symbols);
	if (!byProc) {
		byProc = new WeakMap<ProcedureNode, V>();
		store.set(symbols, byProc);
	}
	return byProc;
}

/**
 * A procedure's view of a module-level table: the procedure's own entries
 * over the module's, read through without copying. The per-procedure
 * environments used to copy the whole module table for every procedure,
 * which on a module with a thousand procedures and hundreds of module-level
 * names was 6% of the analysis pass in four places (issue #139). Iteration
 * yields the module entries the overlay does not shadow, then the overlay.
 */
class LayeredMap<V> implements ReadonlyMap<string, V> {
	constructor(
		private readonly base: ReadonlyMap<string, V>,
		private readonly overlay: ReadonlyMap<string, V>,
	) {}

	get(key: string): V | undefined {
		return this.overlay.has(key) ? this.overlay.get(key) : this.base.get(key);
	}

	has(key: string): boolean {
		return this.overlay.has(key) || this.base.has(key);
	}

	get size(): number {
		let shadowed = 0;
		for (const key of this.overlay.keys()) {
			if (this.base.has(key)) {
				shadowed++;
			}
		}
		return this.base.size + this.overlay.size - shadowed;
	}

	*entries(): MapIterator<[string, V]> {
		for (const entry of this.base.entries()) {
			if (!this.overlay.has(entry[0])) {
				yield entry;
			}
		}
		yield* this.overlay.entries();
	}

	*keys(): MapIterator<string> {
		for (const [key] of this.entries()) {
			yield key;
		}
	}

	*values(): MapIterator<V> {
		for (const [, value] of this.entries()) {
			yield value;
		}
	}

	forEach(callback: (value: V, key: string, map: ReadonlyMap<string, V>) => void, thisArg?: unknown): void {
		for (const [key, value] of this.entries()) {
			callback.call(thisArg, value, key, this);
		}
	}

	[Symbol.iterator](): MapIterator<[string, V]> {
		return this.entries();
	}
}

/**
 * The membership counterpart of {@link LayeredMap}: a procedure's names over
 * the module's. Callers only ask `has`, so that is all it provides.
 */
class LayeredSet {
	constructor(
		private readonly base: ReadonlySet<string> | ReadonlyMap<string, unknown>,
		private readonly overlay: ReadonlySet<string>,
	) {}

	has(key: string): boolean {
		return this.overlay.has(key) || this.base.has(key);
	}
}

// Module-level portion of the type environment, cached per symbols instance:
// environments are built per procedure, so re-walking (and re-lowercasing)
// every module-level declaration for each procedure is O(procedures x
// declarations) on large modules. Procedure entries overwrite module entries in
// the clone exactly as they did in the single-pass build.
const TYPE_ENV_MODULE_BASE = new WeakMap<ReturnType<typeof buildModuleSymbols>, Map<string, string>>();

function typeEnvModuleBase(symbols: ReturnType<typeof buildModuleSymbols>): Map<string, string> {
	const cached = TYPE_ENV_MODULE_BASE.get(symbols);
	if (cached) {
		return cached;
	}
	const base = new Map<string, string>();
	for (const sym of symbols.root.children ?? []) {
		const type = sym.asType ?? (sym.kind === 'moduleVariable' ? defTypeOf(symbols, sym.name) : undefined);
		if (type && !isProcedureKind(sym.kind)) {
			base.set(sym.name.toLowerCase(), type);
		}
	}
	TYPE_ENV_MODULE_BASE.set(symbols, base);
	return base;
}

/**
 * The type a DefType line gives a name declared with no type and no type
 * character: `DefInt A-Z` then `Dim i` is an Integer (issue #285). A Variant
 * is no type to report on, so it gives undefined, as no DefType does.
 */
export function defTypeOf(symbols: ReturnType<typeof buildModuleSymbols>, name: string): string | undefined {
	const type = /^[A-Za-z]\w*$/.test(name) ? symbols.defTypes?.get(name[0].toLowerCase()) : undefined;
	return type === 'Variant' || type === 'Decimal' ? undefined : type;
}

/** Resolve an implicit value type only in the module that owns its declaration. */
function effectiveDeclaredValueType(symbols: ReturnType<typeof buildModuleSymbols>, symbol: VbaSymbol): string | undefined {
	if (symbol.asType || symbol.paramArray || symbol.moduleName.toLowerCase() !== symbols.moduleName.toLowerCase()) {
		return symbol.asType;
	}
	return symbol.kind === 'localVariable' || symbol.kind === 'parameter' || symbol.kind === 'moduleVariable'
		|| symbol.kind === 'function' || symbol.kind === 'propertyGet'
		? defTypeOf(symbols, symbol.name)
		: undefined;
}

export function typeEnvironmentFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): ReadonlyMap<string, string> {
	const cache = perProcedureCache(TYPE_ENVIRONMENTS, symbols);
	const cached = cache.get(proc);
	if (cached) {
		return cached;
	}
	const own = new Map<string, string>();
	const procSym = procedureSymbolFor(symbols, proc);
	const returnType = returnAssignmentTypeFor(proc)
		?? ((proc.procKind === 'Function' || proc.procKind === 'PropertyGet') && !proc.typeSuffix ? defTypeOf(symbols, proc.name) : undefined);
	if (returnType) {
		own.set(proc.name.toLowerCase(), returnType);
	}
	for (const child of procSym?.children ?? []) {
		const type = effectiveDeclaredValueType(symbols, child);
		if (type) {
			own.set(child.name.toLowerCase(), type);
		}
	}
	// A name the procedure assigns with no declaration is a local the DefType
	// types (issue #285), unless it is a host global such as StatusBar or a
	// VBA function such as Mid.
	for (const name of symbols.implicitLocals?.get(proc.span.start) ?? []) {
		const type = defTypeOf(symbols, name);
		if (type && !own.has(name) && !resolveHostGlobal(name, undefined) && !resolveHostGlobalMember(name, undefined) && !resolveRuntimeFunction(name)) {
			own.set(name, type);
		}
	}
	const out = new LayeredMap(typeEnvModuleBase(symbols), own);
	cache.set(proc, out);
	return out;
}

export interface DeclaredValueShape {
	asType?: string;
	isArray: boolean;
	isFixedArray: boolean;
}

// Mirror of typeEnvModuleBase for declaration shapes. The shape objects are
// treated as read-only by every consumer, so clones share them safely.
const DECLARATION_SHAPE_MODULE_BASE = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	Map<string, DeclaredValueShape>
>();

function declarationShapeModuleBase(
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, DeclaredValueShape> {
	const cached = DECLARATION_SHAPE_MODULE_BASE.get(symbols);
	if (cached) {
		return cached;
	}
	const base = new Map<string, DeclaredValueShape>();
	for (const sym of symbols.root.children ?? []) {
		if (isValueDeclarationSymbol(sym)) {
			base.set(sym.name.toLowerCase(), {
				asType: effectiveDeclaredValueType(symbols, sym),
				isArray: sym.isArray === true,
				isFixedArray: sym.arrayBounds !== undefined,
			});
		}
	}
	DECLARATION_SHAPE_MODULE_BASE.set(symbols, base);
	return base;
}

export function declarationShapeEnvironmentFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): ReadonlyMap<string, DeclaredValueShape> {
	const cache = perProcedureCache(DECLARATION_SHAPE_ENVIRONMENTS, symbols);
	const cached = cache.get(proc);
	if (cached) {
		return cached;
	}
	const own = new Map<string, DeclaredValueShape>();
	const procSym = procedureSymbolFor(symbols, proc);
	const returnType = returnAssignmentTypeFor(proc)
		?? ((proc.procKind === 'Function' || proc.procKind === 'PropertyGet') && !proc.typeSuffix ? defTypeOf(symbols, proc.name) : undefined);
	if (returnType) {
		own.set(proc.name.toLowerCase(), {
			asType: returnType,
			isArray: returnAssignmentIsArray(proc),
			isFixedArray: false,
		});
	}
	for (const child of procSym?.children ?? []) {
		if (isValueDeclarationSymbol(child)) {
			own.set(child.name.toLowerCase(), {
				asType: effectiveDeclaredValueType(symbols, child),
				isArray: child.isArray === true,
				isFixedArray: child.arrayBounds !== undefined,
			});
		}
	}
	const out = new LayeredMap(declarationShapeModuleBase(symbols), own);
	cache.set(proc, out);
	return out;
}

export function isValueDeclarationSymbol(sym: VbaSymbol): boolean {
	return (
		sym.kind === 'parameter' ||
		sym.kind === 'localVariable' ||
		sym.kind === 'moduleVariable' ||
		sym.kind === 'constant'
	);
}

export interface SourceDeclaredType {
	resolved: boolean;
	asType?: string;
	/** What the resolved binding is: a variable, a constant, a parameter, a procedure. */
	kind?: VbaSymbol['kind'];
	/** Whether the binding is an array, whose element type `asType` then is. */
	isArray?: boolean;
	/** A Const's value, when it is one string literal: `Const K = "abc"` (issue #255). */
	stringValue?: string;
}

/** The value of a Const that is one string literal, or undefined. */
export function constantStringValue(symbol: VbaSymbol): string | undefined {
	if (symbol.kind !== 'constant' || symbol.defaultRaw === undefined) {
		return undefined;
	}
	const toks = rawExpressionTokens(symbol.defaultRaw);
	return toks.length === 1 && toks[0].kind === 'stringLiteral' ? stringLiteralValue(toks[0].rawText) : undefined;
}

/**
 * The String Consts a procedure sees, each one string literal, by lowercased
 * name (issue #255). A local or parameter of the same name hides a module's.
 */
export function stringConstantsInScope(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode): Map<string, string> {
	const children = procedureSymbolFor(symbols, proc)?.children ?? [];
	const out = new Map<string, string>();
	for (const symbol of [...(symbols.root.children ?? []), ...children]) {
		const value = constantStringValue(symbol);
		if (value !== undefined) {
			out.set(symbol.name.toLowerCase(), value);
		} else if (children.includes(symbol)) {
			out.delete(symbol.name.toLowerCase());
		}
	}
	return out;
}

export type SourceDeclaredTypeResolver = (name: string) => SourceDeclaredType;
export type SourceQualifiedDeclaredTypeResolver = (qualifier: string, name: string) => SourceDeclaredType;

export interface SourceDeclaredShape {
	resolved: boolean;
	shape?: DeclaredValueShape;
}

export function declaredTypeForSourceBinding(
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	name: string,
	context: BareIdentifierContext,
): SourceDeclaredType {
	const binding = sourceIdentifierBinding(
		symbols,
		procSym,
		projectVisibleSymbols,
		name,
		context,
	);
	if (binding.scope === 'unresolved' || binding.scope === 'ambiguous') {
		return { resolved: binding.scope === 'ambiguous' };
	}
	const typed = binding.definitions.find((definition) => effectiveDeclaredValueType(symbols, definition));
	return { resolved: true, asType: typed ? effectiveDeclaredValueType(symbols, typed) : undefined };
}

export function declaredValueTypeForSourceBinding(
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	name: string,
): SourceDeclaredType {
	const binding = sourceIdentifierBinding(
		symbols,
		procSym,
		projectVisibleSymbols,
		name,
		'expression',
	);
	if (binding.scope === 'unresolved' || binding.scope === 'ambiguous') {
		return { resolved: binding.scope === 'ambiguous' };
	}
	const valueDefinitions = binding.definitions.filter(isValueDeclarationSymbol);
	if (valueDefinitions.length === 0) {
		return { resolved: false };
	}
	const typed = valueDefinitions.find((definition) => effectiveDeclaredValueType(symbols, definition));
	const chosen = typed ?? valueDefinitions[0];
	const stringValue = valueDefinitions.length === 1 ? constantStringValue(chosen) : undefined;
	return { resolved: true, asType: effectiveDeclaredValueType(symbols, chosen), kind: chosen.kind, isArray: chosen.isArray === true, ...(stringValue !== undefined ? { stringValue } : {}) };
}

export function declaredValueTypeForQualifiedSourceBinding(
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	qualifier: string,
	name: string,
): SourceDeclaredType {
	const qualifierLower = qualifier.toLowerCase();
	const nameLower = name.toLowerCase();
	const candidates = [
		...(symbols.moduleName.toLowerCase() === qualifierLower
			? symbols.root.children ?? []
			: []),
		...(projectVisibleSymbols ?? []).filter(
			(symbol) => symbol.moduleName.toLowerCase() === qualifierLower,
		),
	];
	if (candidates.length === 0) {
		return { resolved: false };
	}
	const matchingValues = candidates.filter(
		(symbol) =>
			symbol.name.toLowerCase() === nameLower &&
			isValueDeclarationSymbol(symbol),
	);
	if (matchingValues.length === 0) {
		return { resolved: true };
	}
	const typed = matchingValues.find((definition) => effectiveDeclaredValueType(symbols, definition));
	return { resolved: true, asType: typed ? effectiveDeclaredValueType(symbols, typed) : undefined };
}

/**
 * The two lookups an expression-typing pass needs, closed over one procedure's
 * scope: a bare name's declared value type, and a `Qualifier.Name` member's.
 */
export function sourceBindingTypeResolvers(
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
): {
	resolveExpressionType: (name: string) => SourceDeclaredType;
	resolveQualifiedExpressionType: (qualifier: string, name: string) => SourceDeclaredType;
} {
	return {
		resolveExpressionType: (name) =>
			declaredValueTypeForSourceBinding(symbols, procSym, projectVisibleSymbols, name),
		resolveQualifiedExpressionType: (qualifier, name) => {
			const qualified = declaredValueTypeForQualifiedSourceBinding(symbols, projectVisibleSymbols, qualifier, name);
			return qualified.resolved
				? qualified
				: typeFieldDeclaredType(symbols, procSym, projectVisibleSymbols, qualifier, name) ?? qualified;
		},
	};
}

/**
 * The declared type of a field of a user-defined type a variable holds:
 * `t.i` with `Dim t As T1` and `i As Integer` in T1 (issue #369). An array
 * field is left to the shape rules.
 */
export function typeFieldDeclaredType(
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	variable: string,
	field: string,
): SourceDeclaredType | undefined {
	const holder = declaredValueTypeForSourceBinding(symbols, procSym, projectVisibleSymbols, variable);
	const typeName = holder.resolved ? holder.asType?.split('.').pop()?.trim().toLowerCase() : undefined;
	if (!typeName || holder.kind === 'constant') {
		return undefined;
	}
	const userType = [...(symbols.root.children ?? []), ...(projectVisibleSymbols ?? [])]
		.find((symbol) => symbol.kind === 'type' && symbol.name.toLowerCase() === typeName);
	const member = userType?.children?.find((child) => child.kind === 'typeField' && child.name.toLowerCase() === field.toLowerCase());
	return member && !member.isArray ? { resolved: true, asType: member.asType } : undefined;
}

export function declaredShapeForSourceBinding(
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	name: string,
	context: BareIdentifierContext,
): SourceDeclaredShape {
	const binding = sourceIdentifierBinding(
		symbols,
		procSym,
		projectVisibleSymbols,
		name,
		context,
	);
	if (binding.scope === 'unresolved' || binding.scope === 'ambiguous') {
		return { resolved: binding.scope === 'ambiguous' };
	}
	const shaped = binding.definitions.find(
		(definition) => effectiveDeclaredValueType(symbols, definition) || definition.isArray === true,
	);
	return {
		resolved: true,
		shape: {
			asType: shaped ? effectiveDeclaredValueType(symbols, shaped) : undefined,
			isArray: shaped?.isArray === true,
			isFixedArray: shaped?.arrayBounds !== undefined,
		},
	};
}

export interface SourceNameScope {
	/**
	 * Non-callable names visible at the current expression/call site. These block
	 * bare callable resolution before same-module, project, or runtime signatures.
	 */
	callableShadows: { has(lowerName: string): boolean };
	/**
	 * Any source-backed identifier visible in the current procedure. These block
	 * runtime fallback once source/project callable signatures have not resolved.
	 */
	runtimeShadows: { has(lowerName: string): boolean };
}

export function sourceNameScopeFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	projectVisibleSymbols?: readonly VbaSymbol[],
): SourceNameScope {
	const cache = perProcedureCache(SOURCE_NAME_SCOPES, symbols);
	const cached = cache.get(proc);
	if (cached && cached.projectVisibleSymbols === projectVisibleSymbols) {
		return cached.result;
	}
	const own = new Set<string>();
	const procSym = procedureSymbolFor(symbols, proc);
	const runtimeShadows = sourceIdentifierNames({
		currentModule: symbols,
		enclosingProcedure: procSym,
		projectVisibleSymbols,
	});
	for (const child of procSym?.children ?? []) {
		const lower = child.name.toLowerCase();
		if (isNonCallableSymbol(child)) {
			own.add(lower);
		}
	}
	const callableShadows = new LayeredSet(moduleNonCallableSymbols(symbols), own);
	const result: SourceNameScope = { callableShadows, runtimeShadows };
	cache.set(proc, { projectVisibleSymbols, result });
	return result;
}

export function sourceIdentifierBinding(
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	name: string,
	context: BareIdentifierContext,
): BareIdentifierResolution {
	return resolveBareIdentifierBinding({
		currentModule: symbols,
		enclosingProcedure: procSym,
		projectVisibleSymbols,
		name,
		context,
	});
}

export function sourceIdentifierBound(
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	name: string,
	context: BareIdentifierContext,
): boolean {
	return sourceIdentifierBinding(
		symbols,
		procSym,
		projectVisibleSymbols,
		name,
		context,
	).scope !== 'unresolved';
}

/**
 * The integer-constant lookup for one procedure: its own `Const`s over the
 * module's, resolved the way names resolve from inside that procedure.
 */
export function procedureIntegerConstantLookup(
	member: ProcedureNode,
	moduleConstants: ReadonlyMap<string, number | undefined>,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	model?: HostObjectModel,
): IntegerConstantLookup {
	const own = new Map<string, number | undefined>();
	collectBodyLiteralIntegerConstants(member.body, own, activity);
	return scopedIntegerConstantLookup(
		own.size === 0 ? moduleConstants : new LayeredMap(moduleConstants, own),
		symbols,
		procedureSymbolFor(symbols, member),
		projectVisibleSymbols,
		model,
	);
}

/**
 * `constants`, then the whole-number value a local holds at the statement
 * (issue #238): `zz = 12` then `arr(zz)`, `zz = -1` then `ReDim a(zz)`.
 */
export function withKnownLocals(
	constants: IntegerConstantLookup,
	known: ReadonlyMap<string, KnownLocalValue>,
): IntegerConstantLookup {
	return {
		get: (name) => {
			const constant = constants.get(name);
			if (constant !== undefined) {
				return constant;
			}
			const local = known.get(name.toLowerCase());
			return local?.kind === 'number' && Number.isInteger(local.value) ? (local.value as number) : undefined;
		},
	};
}

export function scopedIntegerConstantLookup(
	constants: ReadonlyMap<string, number | undefined>,
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	model?: HostObjectModel,
): IntegerConstantLookup {
	return {
		get(name: string): number | undefined {
			const key = name.toLowerCase();
			if (key.includes('.')) {
				if (constants.has(key)) {
					return constants.get(key);
				}
				return externalIntegerConstantValue(key, model);
			}
			const binding = sourceIdentifierBinding(
				symbols,
				procSym,
				projectVisibleSymbols,
				name,
				'expression',
			);
			if (binding.scope === 'unresolved') {
				if (constants.has(key)) {
					return constants.get(key);
				}
				return externalIntegerConstantValue(key, model);
			}
			if (
				binding.scope === 'ambiguous' ||
				binding.definitions.some((definition) => !isIntegerConstantBindingSymbol(definition))
			) {
				return undefined;
			}
			return constants.get(key);
		},
	};
}

export function inferBareExternalConstantExpressionType(
	name: string,
	span: Span,
	sourceNames?: SourceNameScope,
	model?: HostObjectModel,
): InferredArgumentType | undefined {
	if (runtimeCallableSourceShadowed(name, sourceNames)) {
		return undefined;
	}
	const candidates = [
		inferredExternalConstant(name, resolveRuntimeConstant(name)),
		inferredExternalConstant(name, resolveHostConstant(name, model)),
	].filter((candidate): candidate is InferredArgumentType => candidate !== undefined);
	if (candidates.length !== 1) {
		return undefined;
	}
	return { ...candidates[0], span };
}

export function inferBareExternalObjectExpressionType(
	name: string,
	span: Span,
	sourceNames?: SourceNameScope,
	memberCtx?: MemberCompletionContext,
): InferredArgumentType | undefined {
	if (runtimeCallableSourceShadowed(name, sourceNames)) {
		return undefined;
	}
	const hostType = resolveHostGlobal(name, memberCtx?.model);
	if (hostType) {
		return { type: hostType, label: `${name} As ${hostType}`, span };
	}
	const runtimeObject = resolveRuntimeObject(name);
	if (runtimeObject) {
		return { type: runtimeObject.type, label: `${name} As ${runtimeObject.type}`, span };
	}
	return undefined;
}

const HOST_CONSTANT_QUALIFIERS = new Set(['excel', 'word', 'powerpoint', 'access', 'office']);

export function inferQualifiedExternalConstantExpressionType(
	qualifier: string,
	name: string,
	span: Span,
	model?: HostObjectModel,
): InferredArgumentType | undefined {
	const lower = qualifier.toLowerCase();
	if (lower === 'vba') {
		const inferred = inferredExternalConstant(`${qualifier}.${name}`, resolveRuntimeConstant(name));
		return inferred ? { ...inferred, span } : undefined;
	}
	// Resolved against the CURRENT host's model: `Word.wdRed` answers in a
	// Word module, misses in an Excel one (issue #24).
	if (HOST_CONSTANT_QUALIFIERS.has(lower)) {
		const inferred = inferredExternalConstant(`${qualifier}.${name}`, resolveHostConstant(name, model));
		return inferred ? { ...inferred, span } : undefined;
	}
	return undefined;
}

export function inferredExternalConstant(
	displayName: string,
	constant: { type?: string; value?: string | number } | undefined,
): InferredArgumentType | undefined {
	if (!constant) {
		return undefined;
	}
	const declaredType = constant.type;
	const normalizedDeclared = normalizeType(declaredType);
	const numericValue = numericExternalConstantValue(constant.value);
	if (numericValue !== undefined) {
		return {
			type: 'Long',
			label: `${displayName} As ${declaredType ?? 'Long'}`,
			span: { start: 0, end: 0 },
			numericValue,
			numericText: displayName,
			// Mark this as a named-constant origin so overflow diagnostics phrase
			// it as a constant value rather than a "numeric literal".
			numericConstantName: displayName,
		};
	}
	if (normalizedDeclared === 'string') {
		return {
			type: 'String',
			label: `${displayName} As ${declaredType ?? 'String'}`,
			span: { start: 0, end: 0 },
		};
	}
	return undefined;
}

export function isIntegerConstantBindingSymbol(symbol: VbaSymbol): boolean {
	return symbol.kind === 'constant' || symbol.kind === 'enumMember';
}

export function bareCallableSourceShadowed(
	name: string,
	sourceNames: SourceNameScope | undefined,
): boolean {
	return sourceNames?.callableShadows.has(name.toLowerCase()) === true;
}

export function runtimeCallableSourceShadowed(
	name: string,
	sourceNames: SourceNameScope | undefined,
): boolean {
	// A Public Function InStr compiles, and `InStr(0, "abc", "a")` still calls
	// VBA's, which raises 5 (issue #280, measured in Excel 16.0).
	const lower = name.toLowerCase();
	return lower !== 'instr' && sourceNames?.runtimeShadows.has(lower) === true;
}

export function returnAssignmentTypeFor(proc: ProcedureNode): string | undefined {
	return (proc.procKind === 'Function' || proc.procKind === 'PropertyGet')
		? proc.returnType
		: undefined;
}

export function returnAssignmentIsArray(proc: ProcedureNode): boolean {
	return /\(\s*\)\s*$/i.test(returnAssignmentTypeFor(proc) ?? '');
}

export function expressionCalls(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallArguments[] {
	const toks = statementTokens(source, span);
	const out: CallArguments[] = [];
	const firstExecutable = firstExecutableTokenIndex(toks);
	const statementHead = tokenText(toks[firstExecutable]);
	for (let i = 0; i < toks.length - 1; i++) {
		const callName = parenthesizedCallNameAt(toks, i);
		if (!callName) {
			continue;
		}
		const { name, parenIndex, nameEndIndex } = callName;
		// `Take (i)` as a statement passes `(i)` as its one argument: the space
		// makes the parentheses part of the argument, which is then a copy
		// (issue #111). `Take(i)` glued, and `x = Take (i)` inside an
		// expression, are calls with a list.
		const argumentsParenthesized = i === firstExecutable
			&& toks[parenIndex].start > toks[nameEndIndex].end;
		// `ReDim Three(2)` inside Function Three sizes the return array, and a
		// ReDim target anywhere is a variable, never a call (issue #115).
		if (statementHead === 'redim' && isRedimTargetAt(toks, i, firstExecutable)) {
			continue;
		}
		const qualifier =
			i >= 2 && toks[i - 1].rawText === '.'
				? tokenName(toks[i - 2])
				: undefined;
		const lookupKey = qualifier ? qualifiedProcedureKey(qualifier, name) : undefined;
		if (qualifier && !moduleSignatures.has(lookupKey!)) {
			continue; // host/member calls need receiver binding before checking
		}
		if (!qualifier && i > 0 && toks[i - 1].rawText === '.') {
			continue;
		}
		if (
			lookupKey
				? !moduleSignatures.has(lookupKey)
				: !callableSignatureFor(name, moduleSignatures, sourceNames)
		) {
			continue;
		}
		const close = matchParenFrom(toks, parenIndex);
		if (close < 0) {
			continue;
		}
		const inner = toks.slice(parenIndex + 1, close);
		const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
		out.push({
			name,
			qualifier,
			lookupKey,
			nameSpan: { start: span.start + toks[i].start, end: span.start + toks[nameEndIndex].end },
			slots: split.slots,
			slotSpans: split.spans,
			sliceStart: span.start,
			...(argumentsParenthesized ? { argumentsParenthesized: true } : {}),
		});
	}
	return out;
}

/**
 * Whether the name at `index` is a target of the ReDim statement the tokens
 * spell: at depth 0, right after `ReDim`, `Preserve`, or a separating comma.
 */
function isRedimTargetAt(toks: readonly VbaToken[], index: number, firstExecutable: number): boolean {
	let depth = 0;
	for (let k = firstExecutable + 1; k < index; k++) {
		const raw = toks[k].rawText;
		if (raw === '(') {
			depth++;
		} else if (raw === ')') {
			depth--;
		}
	}
	if (depth !== 0) {
		return false;
	}
	const prev = tokenText(toks[index - 1]);
	return prev === 'redim' || prev === 'preserve' || prev === ',';
}

export interface ParenthesizedCallName {
	name: string;
	parenIndex: number;
	nameEndIndex: number;
}

export function parenthesizedCallNameAt(
	toks: readonly VbaToken[],
	nameIndex: number,
): ParenthesizedCallName | undefined {
	const baseName = tokenName(toks[nameIndex]);
	if (!baseName) {
		return undefined;
	}
	const suffix = toks[nameIndex + 1];
	if (
		suffix?.rawText === '$' &&
		toks[nameIndex].end === suffix.start &&
		toks[nameIndex + 2]?.rawText === '(' &&
		suffix.end === toks[nameIndex + 2].start
	) {
		return { name: `${baseName}$`, parenIndex: nameIndex + 2, nameEndIndex: nameIndex + 1 };
	}
	if (toks[nameIndex + 1]?.rawText === '(') {
		return { name: baseName, parenIndex: nameIndex + 1, nameEndIndex: nameIndex };
	}
	return undefined;
}

export interface BoundMemberCall {
	call: CallArguments;
	signature: CallableTypeSignature;
}

export function memberExpressionCalls(
	source: string,
	span: Span,
	memberCtx: MemberCompletionContext,
): BoundMemberCall[] {
	const toks = statementTokens(source, span);
	const standaloneEmptyCall = standaloneEmptyParenthesizedCallStatement(source, span);
	const out: BoundMemberCall[] = [];
	for (let i = 1; i < toks.length - 1; i++) {
		const name = tokenName(toks[i]);
		if (!name || toks[i - 1]?.rawText !== '.' || toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const member = resolveExactMemberCompletion(
			source,
			name,
			span.start + toks[i].end,
			memberCtx,
		);
		if (!member?.signature) {
			continue;
		}
		const inner = toks.slice(i + 2, close);
		const callSpan = {
			start: span.start + toks[i].start,
			end: span.start + toks[close].end,
		};
		if (
			standaloneEmptyCall?.isMember &&
			standaloneEmptyCall.span.start === callSpan.start &&
			standaloneEmptyCall.span.end === callSpan.end
		) {
			continue;
		}
		const parsed = parseRuntimeDisplaySignature(member.name, member.signature);
		if (isPropertyResultIndexing(member, parsed, inner)) {
			continue;
		}
		// A Function of a project class gives a value, and `k.Items(1)` may
		// index it (issue #609).
		const signature = member.kind === 'method' && !member.sub && (member.definitions?.length ?? 0) > 0
			? { ...parsed, valued: true, returnType: member.returns ?? parsed.returnType }
			: parsed;
		const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
		out.push({
			signature,
			call: {
				name: member.name,
				nameSpan: { start: callSpan.start, end: span.start + toks[i].end },
				slots: split.slots,
				slotSpans: split.spans,
				sliceStart: span.start,
			},
		});
	}
	return out;
}

export function memberStatementCalls(
	source: string,
	span: Span,
	memberCtx: MemberCompletionContext,
): BoundMemberCall[] {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (toks.length === 0 || topLevelOperatorIndex(toks, '=') >= 0) {
		return [];
	}
	const explicitCall = tokenText(toks[0]) === 'call';
	const chainStart = explicitCall ? 1 : 0;
	if (!tokenName(toks[chainStart]) && toks[chainStart]?.rawText !== '.') {
		return [];
	}
	const out: BoundMemberCall[] = [];
	const firstMemberIndex = toks[chainStart]?.rawText === '.' ? chainStart + 1 : chainStart + 2;
	for (let i = firstMemberIndex; i < toks.length; i++) {
		const name = tokenName(toks[i]);
		if (!name || toks[i - 1]?.rawText !== '.') {
			continue;
		}
		if (!isMemberStatementChainThrough(toks, chainStart, i)) {
			continue;
		}
		const next = toks[i + 1];
		if (next?.rawText === '(') {
			continue; // parenthesized member calls are handled by memberExpressionCalls
		}
		if (explicitCall && next) {
			continue; // Call p.Save arg is a call-requires-parens syntax error
		}
		if (next) {
			const gap = source.slice(span.start + toks[i].end, span.start + next.start);
			if (!/\s/.test(gap) || !isMemberParenlessArgumentStart(next)) {
				continue;
			}
		}
		const member = resolveExactMemberCompletion(
			source,
			name,
			span.start + toks[i].end,
			memberCtx,
		);
		if (!member?.signature) {
			continue;
		}
		const argToks = toks.slice(i + 1);
		const split = argToks.length === 0 ? emptyArgSplit() : splitArgSlots(argToks, span.start);
		out.push({
			signature: parseRuntimeDisplaySignature(member.name, member.signature),
			call: {
				name: member.name,
				nameSpan: { start: span.start + toks[i].start, end: span.start + toks[i].end },
				explicitCall,
				slots: split.slots,
				slotSpans: split.spans,
				sliceStart: span.start,
			},
		});
		break;
	}
	return out;
}

export function isPropertyResultIndexing(
	member: MemberCompletion,
	signature: CallableTypeSignature,
	inner: readonly VbaToken[],
): boolean {
	return member.kind === 'property' &&
		signature.params.length === 0 &&
		inner.length > 0;
}

export function isMemberStatementChainThrough(
	toks: readonly VbaToken[],
	startIdx: number,
	memberIdx: number,
): boolean {
	if (toks[startIdx]?.rawText === '.') {
		if (!tokenName(toks[startIdx + 1])) {
			return false;
		}
		if (startIdx + 1 === memberIdx) {
			return true;
		}
		return isMemberStatementChainThrough(toks, startIdx + 1, memberIdx);
	}
	if (!tokenName(toks[startIdx])) {
		return false;
	}
	// A keyword spaced off a dot is a statement's keyword before a With
	// member, not a receiver: `If .Count = 0 Then .Add 1` (issue #198).
	const next = toks[startIdx + 1];
	if (toks[startIdx].kind === 'keyword' && next?.rawText === '.' && next.start > toks[startIdx].end) {
		return false;
	}
	let i = startIdx + 1;
	while (i < toks.length) {
		const raw = toks[i]?.rawText;
		if (raw === '(') {
			const close = matchParenFrom(toks, i);
			if (close < 0 || close >= memberIdx) {
				return false;
			}
			i = close + 1;
			continue;
		}
		if (raw !== '.') {
			return false;
		}
		const nameIdx = i + 1;
		if (!tokenName(toks[nameIdx])) {
			return false;
		}
		if (nameIdx === memberIdx) {
			return true;
		}
		i = nameIdx + 1;
	}
	return false;
}

export function isMemberParenlessArgumentStart(tok: VbaToken): boolean {
	if (
		tok.kind === 'identifier' ||
		tok.kind === 'keyword' ||
		tok.kind === 'bracketedIdentifier' ||
		tok.kind === 'stringLiteral' ||
		tok.kind === 'dateLiteral' ||
		tok.kind === 'integerLiteral' ||
		tok.kind === 'floatLiteral'
	) {
		return true;
	}
	return tok.rawText === ',' || tok.rawText === '+' || tok.rawText === '-';
}

export interface ArgumentObjectQueries {
	resolveType: ReturnType<typeof createObjectAssignmentTypeResolver>;
	shareInterfaces: ReturnType<typeof createProjectInterfaceSharingLookup>;
	implementsType: ReturnType<typeof createObjectTypeImplementationLookup>;
	needsIndex?: ReturnType<typeof createObjectDefaultQueries>['needsIndex'];
}

export function validateArgumentTypes(
	call: CallArguments,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
	source: string,
	memberCtx: MemberCompletionContext,
	push: PushFn,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
	heldClassOf?: (lower: string) => string | undefined,
	heldNull?: (lower: string) => boolean,
	heldNumber?: (lower: string) => number | string | undefined,
	objectQueries?: ArgumentObjectQueries,
): void {
	const sig = callableSignatureForCall(call, moduleSignatures, sourceNames);
	if (!sig || sig.params.length === 0) {
		return;
	}
	validateArgumentTypesForSignature(
		sig,
		call,
		env,
		moduleSignatures,
		sourceNames,
		source,
		memberCtx,
		push,
		resolveExpressionType,
		resolveQualifiedExpressionType,
		heldClassOf,
		heldNull,
		heldNumber,
		objectQueries,
	);
}

export function validateArgumentTypesForSignature(
	sig: CallableTypeSignature,
	call: CallArguments,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
	source: string,
	memberCtx: MemberCompletionContext,
	push: PushFn,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
	heldClassOf?: (lower: string) => string | undefined,
	heldNull?: (lower: string) => boolean,
	heldNumber?: (lower: string) => number | string | undefined,
	objectQueries?: ArgumentObjectQueries,
): void {
	if (sig.params.length === 0) {
		return;
	}
	const paramsByName = new Map(
		sig.params.map((p) => [stripHeaderBrackets(p.name).toLowerCase(), p]),
	);
	let positionalIndex = 0;
	for (let i = 0; i < call.slots.length; i++) {
		const named = namedArgumentSlot(call.slots[i]);
		let param: CallableParamType | undefined;
		let valueSlot = call.slots[i];
		if (named) {
			param = paramsByName.get(named.name.toLowerCase());
			valueSlot = named.value;
		} else {
			param = sig.params[Math.min(positionalIndex, sig.params.length - 1)];
			if (!param || (positionalIndex >= sig.params.length && !param.paramArray)) {
				continue;
			}
			positionalIndex++;
		}
		if (!param) {
			continue;
		}
		const expected = param.type;
		if (!expected) {
			continue;
		}
		const byRefMismatch = call.argumentsParenthesized
			? undefined
			: byRefVariableTypeMismatch(
				param,
				valueSlot,
				call.sliceStart,
				env,
				resolveExpressionType,
				resolveQualifiedExpressionType,
			);
		if (byRefMismatch) {
			push(
				'byRefArgumentTypeMismatch',
				`ByRef argument '${param.name}' of '${sig.name}' expects ${expected}, but '${byRefMismatch.name}' is declared as ${byRefMismatch.actual}. This is a VBE compile error: ByRef argument type mismatch.`,
				byRefMismatch.span,
			);
			continue;
		}
		// An array parameter takes an array variable; anything else is
		// argument-shape-mismatch's, not a value to convert (issue #410).
		if (param.isArray) {
			continue;
		}
		// `TakeL(1 + Null)`: an operator on Null gives Null, which a typed
		// parameter refuses (issue #324, measured in Excel 16.0).
		const nullSlot = valueSlot.filter((tok) => tok.kind !== 'comment');
		const scalarExpected = normalizeType(expected);
		if (nullSlot.length > 1 && scalarExpected && scalarExpected !== 'variant' && isKnownScalarType(scalarExpected)
			&& operatorYieldsNull(nullSlot, (tok) => tokenText(tok) === 'null' || heldNull?.(tokenName(tok)?.toLowerCase() ?? '') === true)) {
			push(
				'argumentTypeMismatch',
				`Argument '${param.name}' of '${sig.name}' expects ${expected}, but '${nullSlot.map((tok) => tok.rawText).join(' ').replace(/ ?([()]) ?/g, '$1')}' is Null: an operator on Null gives Null. Null cannot be coerced to this scalar type. This will raise Run-time error '94': Invalid use of Null.`,
				{ start: call.sliceStart + nullSlot[0].start, end: call.sliceStart + nullSlot[nullSlot.length - 1].end },
			);
			continue;
		}
		const stringArithmetic = nonnumericStringArithmeticOperand(
			expected,
			valueSlot,
			call.sliceStart,
		);
		if (stringArithmetic) {
			push(
				'stringArithmeticCoercion',
				`Argument '${param.name}' of '${sig.name}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
				stringArithmetic.span,
			);
			continue;
		}
		let actual = inferArgumentType(
			valueSlot,
			call.sliceStart,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
		const kindProblem = objectValueArgumentProblem(expected, valueSlot, actual, memberCtx, sourceNames, (name) =>
			env.has(name.toLowerCase()) || resolveExpressionType?.(name).resolved === true, heldClassOf,
			param.byRef === false || call.argumentsParenthesized === true, env, objectQueries);
		if (kindProblem) {
			push(
				kindProblem.rule,
				`Argument '${param.name}' of '${sig.name}' expects ${expected}, but got ${kindProblem.what}. ${kindProblem.reason}`,
				{ start: call.sliceStart + kindProblem.tokens[0].start, end: call.sliceStart + kindProblem.tokens[kindProblem.tokens.length - 1].end },
			);
			continue;
		}
		// A Variant local a straight line has just set to Null is Null here,
		// for both checks below, whether or not it was given a type, and so
		// is an element `a(0)` or `a(i)` (issue #332, measured in Excel 16.0).
		const element = valueSlot.length === 4 && valueSlot[1].rawText === '(' && valueSlot[3].rawText === ')' ? tokenName(valueSlot[0]) : undefined;
		const heldName = valueSlot.length === 1 ? tokenName(valueSlot[0])?.toLowerCase()
			: element !== undefined ? `${element.toLowerCase()}(${valueSlot[2].rawText.toLowerCase()})` : undefined;
		// Mid hands a Null string back without reading its Length: `Mid(n0,
		// 1, n2)` with both Null runs (issue #664, measured in Excel 16.0).
		const first = call.slots[0]?.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline') ?? [];
		const nullString = sig.name.toLowerCase() === 'mid' && param.name.toLowerCase() === 'length' && first.length === 1
			&& (tokenText(first[0]) === 'null' || heldNull?.(tokenName(first[0])?.toLowerCase() ?? '') === true);
		const heldNullHere = !nullString && heldName !== undefined && heldNull?.(heldName) === true && [undefined, 'variant'].includes(normalizeType(actual?.type));
		if (heldNullHere) {
			const last = valueSlot[valueSlot.length - 1];
			const span = { start: call.sliceStart + valueSlot[0].start, end: call.sliceStart + last.end };
			const text = valueSlot.map((tok) => tok.rawText).join('');
			actual = { ...(actual ?? { span }), type: 'Null', label: `'${text}', which holds Null here` };
		}
		if (!actual) {
			continue;
		}
		// A local known to hold a number that is not whole is range-checked as
		// that number: `Dim c As Currency: c = 922337203685477.5807@: Space(c)`
		// overflows (issue #332). A whole one past the Long range is
		// runtime-argument-value's (issue #336).
		const heldValue = heldName !== undefined && actual.numericValue === undefined && actual.floatValue === undefined ? heldNumber?.(heldName) : undefined;
		if (typeof heldValue === 'number' && !Number.isInteger(heldValue)) {
			actual = { ...actual, heldBy: valueSlot[0].rawText, floatValue: heldValue };
		}
		// A local's whole number or String passed by value to the project's own
		// procedure converts as a literal does: `v = -3` then `S v` with
		// `ByVal p As Byte` raises 6, and "abc" into an Integer 13 (issue #558,
		// measured in Excel 16.0). `S (v)` passes a copy whatever p is.
		const inParens = valueSlot.length === 3 && valueSlot[0].rawText === '(' && valueSlot[2].rawText === ')' ? tokenName(valueSlot[1])?.toLowerCase() : undefined;
		const copiedName = (param.byRef === false || call.argumentsParenthesized === true) && heldName !== undefined ? heldName : inParens;
		const ownProcedure = moduleSignatures.get(call.lookupKey ?? call.name.toLowerCase()) === sig;
		// A Boolean's True is no -1 here: into a Byte it is 255 (issue #664,
		// measured in Excel 16.0, as #624 measured for a Let).
		const boolean = copiedName !== undefined && normalizeType(env.get(copiedName)) === 'boolean';
		const copied = ownProcedure && copiedName !== undefined && !boolean && actual.numericValue === undefined && actual.floatValue === undefined && actual.stringValue === undefined
			? heldNumber?.(copiedName) : undefined;
		const holder = inParens !== undefined && copiedName === inParens ? valueSlot[1].rawText : valueSlot[0].rawText;
		if (typeof copied === 'number') {
			actual = Number.isInteger(copied) ? { ...actual, heldBy: holder, numericValue: copied } : { ...actual, heldBy: holder, floatValue: copied };
		} else if (typeof copied === 'string') {
			actual = { ...actual, heldBy: holder, type: 'String', stringValue: copied };
		}
		// A Variant parameter the function still refuses Null for: CStr(Null),
		// Chr(Null), Asc(Null) raise 94 where Left(Null, 1) hands Null back
		// (issue #104).
		if (param.nullRaises && normalizeType(actual.type) === 'null') {
			push(
				'argumentTypeMismatch',
				`Argument '${param.name}' of '${sig.name}' cannot be Null${heldNullHere ? `, and ${actual.label}` : ''}. This will raise Run-time error '94': Invalid use of Null.`,
				actual.span,
			);
			continue;
		}
		const reason = incompatibilityReason(expected, actual);
		if (!reason) {
			continue;
		}
		const rule =
			normalizeType(expected) === 'object'
				? 'argumentObjectTypeMismatch'
				: 'argumentTypeMismatch';
		push(
			rule,
			`Argument '${param.name}' of '${sig.name}' expects ${expected}, but got ${actual.label}. ${reason}`,
			actual.span,
		);
	}
}

/** The workbooks a Worksheets or Charts property is read from: `ThisWorkbook.Worksheets`. */
const SHEETS_RECEIVERS: ReadonlySet<string> = new Set(['application', 'thisworkbook', 'activeworkbook']);

/**
 * Excel's Worksheets and Charts properties, bare, on Application or a
 * workbook, or given an Array of names, hand back a Sheets object (issue
 * #404, measured in Excel 16.0): TypeName(Worksheets) is "Sheets", and Set
 * into a variable As Worksheets or As Charts raises 13. The model types the
 * property by what it holds, so `Worksheets(1)` stays a Worksheet; this
 * names the property when the target is one of those two collections.
 */
export function sheetsFromCollectionProperty(
	value: readonly VbaToken[],
	expected: string | undefined,
	sourceNames: SourceNameScope,
	memberCtx: MemberCompletionContext,
): { text: string; collection: 'Worksheets' | 'Charts' } | undefined {
	const target = resolveHostAlias(expected ?? '', memberCtx.model)?.toLowerCase();
	const collection = target === 'excel.worksheets' ? 'Worksheets' : target === 'excel.charts' ? 'Charts' : undefined;
	if (!collection || value.length === 0) {
		return undefined;
	}
	// `Worksheets(Array("Sheet1"))` is a Sheets object too; `Worksheets(1)` is a sheet.
	let end = value.length - 1;
	if (value[end].rawText === ')') {
		const open = value.findIndex((tok, i) => tok.rawText === '(' && matchParenFrom(value, i) === end);
		if (open < 1 || tokenText(value[open + 1]) !== 'array' || value[open + 2]?.rawText !== '(' || matchParenFrom(value, open + 2) !== end - 1) {
			return undefined;
		}
		end = open - 1;
	}
	const property = tokenText(value[end]);
	if (property !== 'worksheets' && property !== 'charts') {
		return undefined;
	}
	const qualifier = value.slice(0, end);
	const bare = qualifier.length === 0 && !sourceNames.runtimeShadows.has(property);
	const onWorkbook = qualifier.length === 2 && qualifier[1].rawText === '.' && SHEETS_RECEIVERS.has(tokenText(qualifier[0]));
	return bare || onWorkbook ? { text: value.map((tok) => tok.rawText).join(''), collection } : undefined;
}

/** What a held-class lookup gives for a Variant known to hold Empty, a number or a String: no object. */
export const VALUE_HELD = '(value)';

/**
 * An object where a parameter takes a value, or a value where it takes an
 * object (issue #223, measured in Excel 16.0):
 *
 *  - Nothing into a Long or String parameter: "Invalid use of object".
 *  - New Collection there: "Argument not optional", since its default
 *    member Item needs an index.
 *  - A literal, True, False, a date, or an expression of a known scalar
 *    type into a Collection or other known object parameter: "Type
 *    mismatch". So is a scalar variable passed by value (issue #410).
 *    Each is a compile error.
 *  - A Variant holding no object, passed by value to one: 424 when the
 *    call runs (issue #410).
 *  - Array(...) or Split(...) into a Long or String parameter: an array,
 *    which raises 13 when the call runs.
 */
const ARITHMETIC_OPERATORS: ReadonlySet<string> = new Set(['+', '-', '*', '/', '\\', '^', 'mod']);

/**
 * Whether the tokens are an arithmetic expression of number literals and
 * locals declared a scalar type, `b + 0` or `d * 2`, with no call, member
 * or parenthesis in it: its value is a scalar (issue #647).
 */
export function arithmeticOfScalars(toks: readonly VbaToken[], env: ReadonlyMap<string, string>): boolean {
	let operand = true;
	let operators = 0;
	for (const tok of toks) {
		if (operand) {
			const lower = tokenName(tok)?.toLowerCase();
			const declared = lower !== undefined && tok.kind === 'identifier' ? normalizeType(env.get(lower)) : undefined;
			if (tok.kind !== 'integerLiteral' && tok.kind !== 'floatLiteral' && !(declared && isKnownScalarType(declared))) {
				return false;
			}
		} else if (!ARITHMETIC_OPERATORS.has(tokenText(tok) || tok.rawText)) {
			return false;
		} else {
			operators++;
		}
		operand = !operand;
	}
	return !operand && operators > 0;
}

function objectValueArgumentProblem(
	expected: string,
	slot: readonly VbaToken[],
	actual: InferredArgumentType | undefined,
	memberCtx: MemberCompletionContext,
	sourceNames: SourceNameScope | undefined,
	isDeclared: (name: string) => boolean,
	heldClassOf?: (lower: string) => string | undefined,
	byValue = false,
	env: ReadonlyMap<string, string> = new Map(),
	objectQueries?: ArgumentObjectQueries,
): { rule: 'argumentObjectTypeMismatch' | 'argumentTypeMismatch'; what: string; reason: string; tokens: readonly VbaToken[] } | undefined {
	const toks = unwrapOuterParens(slot.filter((tok) => tok.kind !== 'comment' && tok.kind !== 'newline'));
	if (toks.length === 0) {
		return undefined;
	}
	const expectedType = normalizeType(expected);
	if (expectedType && isKnownScalarType(expectedType)) {
		const needsIndex = objectQueries?.needsIndex ?? ((type: string | undefined) => objectValueNeedsIndex(type, memberCtx));
		if (toks.length === 1 && tokenText(toks[0]) === 'nothing') {
			return { rule: 'argumentObjectTypeMismatch', what: 'Nothing', reason: 'This is a VBE compile error: Invalid use of object.', tokens: toks };
		}
		if (toks.length === 2 && tokenText(toks[0]) === 'new' && needsIndex(toks[1].rawText)) {
			return { rule: 'argumentObjectTypeMismatch', what: `New ${toks[1].rawText}, whose default member Item needs an index`, reason: 'This is a VBE compile error: Argument not optional.', tokens: toks };
		}
		// A Collection variable passed by value is read for its value, the
		// same way (issue #647, measured in Excel 16.0). ByRef it is
		// byref-argument-type-mismatch's.
		const passedName = toks.length === 1 ? tokenName(toks[0])?.toLowerCase() : undefined;
		const passedType = passedName ? env.get(passedName) : undefined;
		if (byValue && passedType && isDeclared(toks[0].rawText) && needsIndex(passedType)) {
			return { rule: 'argumentObjectTypeMismatch', what: `'${toks[0].rawText}', declared ${passedType}, whose default member Item needs an index`, reason: 'This is a VBE compile error: Argument not optional.', tokens: toks };
		}
		const callee = tokenText(toks[0]);
		if ((callee === 'array' || callee === 'split') && toks[1]?.rawText === '(' && matchParenFrom(toks, 1) === toks.length - 1
			&& !runtimeCallableSourceShadowed(toks[0].rawText, sourceNames)) {
			return { rule: 'argumentTypeMismatch', what: `${toks[0].rawText}(...), an array`, reason: "This will raise Run-time error '13': Type mismatch.", tokens: toks };
		}
		return undefined;
	}
	// A literal, True, False or a date, and an expression of a known scalar
	// type such as `v + 0` (issue #410, measured in Excel 16.0).
	const atom = toks[0].rawText === '-' ? toks.slice(1) : toks;
	const literal = atom.length === 1 && (['integerLiteral', 'floatLiteral', 'stringLiteral', 'dateLiteral'].includes(atom[0].kind)
		|| (toks.length === 1 && ['true', 'false'].includes(tokenText(atom[0]))));
	// `b + 0` and `d + 0` with b a Boolean and d a Date too (issue #647).
	const scalarExpression = toks.length > 1 && !literal
		&& ((actual !== undefined && isKnownScalarType(normalizeType(actual.type) ?? '')) || arithmeticOfScalars(toks, env));
	if ((literal || scalarExpression) && expectedType !== 'object' && isKnownObjectAssignmentType(expected, memberCtx, objectQueries?.resolveType)) {
		return { rule: 'argumentObjectTypeMismatch', what: actual?.label ?? toks.map((tok) => tok.rawText).join(' '), reason: 'An object parameter takes an object. This is a VBE compile error: Type mismatch.', tokens: toks };
	}
	// An object of another class, as a Set of it would be: TakeWs(Range("A1"))
	// and TakeWs(ThisWorkbook) into a Worksheet raise 13 when the call runs
	// (issue #223). ActiveSheet and Sheets(1), declared Object, run. A
	// declared variable is judged by what it holds, not its declared class:
	// one still Nothing passes, and one known to hold another class raises 13
	// ByVal or ByRef (issue #246, measured in Excel 16.0).
	const declaredName = toks.length === 1 && tokenName(toks[0]) !== undefined && isDeclared(toks[0].rawText);
	const held = declaredName ? heldClassOf?.(tokenName(toks[0])!.toLowerCase()) : undefined;
	// A scalar variable passed by value is a value, as a literal is; ByRef it
	// is byref-argument-type-mismatch's. A Variant holding Empty or a value
	// raises 424 (issue #410, measured in Excel 16.0).
	if (declaredName && byValue && expectedType !== 'object' && isKnownObjectAssignmentType(expected, memberCtx, objectQueries?.resolveType)) {
		const declared = normalizeType(env.get(tokenName(toks[0])!.toLowerCase()));
		if (declared && isKnownScalarType(declared)) {
			return { rule: 'argumentObjectTypeMismatch', what: `'${toks[0].rawText}', declared ${env.get(tokenName(toks[0])!.toLowerCase())}`, reason: 'An object parameter takes an object. This is a VBE compile error: Type mismatch.', tokens: toks };
		}
		if (held === VALUE_HELD) {
			return { rule: 'argumentTypeMismatch', what: `'${toks[0].rawText}', a Variant that holds no object here`, reason: "An object parameter takes an object. This will raise Run-time error '424': Object required.", tokens: toks };
		}
	}
	if (held && held !== VALUE_HELD && expectedType !== 'object' && isKnownObjectAssignmentType(expected, memberCtx, objectQueries?.resolveType)) {
		const holding = { type: held, label: `'${toks[0].rawText}', which holds a ${held} here`, span: { start: toks[0].start, end: toks[0].end } };
		const reason = objectAssignmentIncompatibilityReason(expected, holding, memberCtx, objectQueries?.resolveType, objectQueries?.shareInterfaces, objectQueries?.implementsType);
		if (reason) {
			return { rule: 'argumentTypeMismatch', what: holding.label, reason: `${reason} This will raise Run-time error '13': Type mismatch.`, tokens: toks };
		}
	}
	// `TakeC(ActiveSheet)` into a Collection, as the Object holding it (issue #685).
	if (!declaredName && toks.length === 1 && tokenText(toks[0]) === 'activesheet' && !isDeclared(toks[0].rawText)
		&& resolveHostGlobal('ActiveSheet', memberCtx.model) !== undefined && expectedType !== 'object' && isKnownObjectAssignmentType(expected, memberCtx, objectQueries?.resolveType)) {
		const sheet = { type: 'Worksheet or Chart', label: `'${toks[0].rawText}', a Worksheet or a Chart`, span: { start: toks[0].start, end: toks[0].end } };
		const reason = objectAssignmentIncompatibilityReason(expected, sheet, memberCtx, objectQueries?.resolveType, objectQueries?.shareInterfaces, objectQueries?.implementsType);
		if (reason) {
			return { rule: 'argumentTypeMismatch', what: sheet.label, reason: `${reason} This will raise Run-time error '13': Type mismatch.`, tokens: toks };
		}
	}
	// `TakeW(Worksheets)` into a parameter As Worksheets (issue #404).
	const sheets = !declaredName && sourceNames ? sheetsFromCollectionProperty(toks, expected, sourceNames, memberCtx) : undefined;
	if (sheets) {
		return { rule: 'argumentTypeMismatch', what: `'${sheets.text}', which returns a Sheets object`, reason: `Excel's Worksheets and Charts properties return a Sheets object, never a ${sheets.collection} one. This will raise Run-time error '13': Type mismatch.`, tokens: toks };
	}
	if (actual && !declaredName && expectedType !== 'object' && isKnownObjectAssignmentType(expected, memberCtx, objectQueries?.resolveType)
		&& !isKnownScalarType(normalizeType(actual.type) ?? '')) {
		const reason = objectAssignmentIncompatibilityReason(expected, actual, memberCtx, objectQueries?.resolveType, objectQueries?.shareInterfaces, objectQueries?.implementsType);
		if (reason) {
			return { rule: 'argumentTypeMismatch', what: actual.label, reason: `${reason} This will raise Run-time error '13': Type mismatch.`, tokens: toks };
		}
	}
	return undefined;
}

export function callableSignatureForCall(
	call: CallArguments,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallableTypeSignature | undefined {
	if (call.lookupKey) {
		return moduleSignatures.get(call.lookupKey);
	}
	return callableSignatureFor(call.name, moduleSignatures, sourceNames);
}

export function byRefVariableTypeMismatch(
	param: CallableParamType,
	slot: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): { name: string; actual: string; span: Span } | undefined {
	if (!param.byRef || !param.type) {
		return undefined;
	}
	const expected = normalizeType(param.type);
	if (!byRefExact(expected)) {
		return undefined;
	}
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	let name: string | undefined;
	let actualRaw: string | undefined;
	let span: Span | undefined;
	if (toks.length === 1) {
		name = tokenName(toks[0]);
		if (!name) {
			return undefined;
		}
		const declaredType = resolveExpressionType?.(name);
		// A Const is passed as a temporary copy, so its type never has to
		// match (issue #111: `Take(K)` with K an Integer Const compiles).
		if (declaredType?.resolved && declaredType.kind === 'constant') {
			return undefined;
		}
		// A non-array into an array parameter is argument-shape-mismatch's.
		if (param.isArray && declaredType?.resolved && !declaredType.isArray) {
			return undefined;
		}
		actualRaw = declaredType?.resolved
			? declaredType.asType
			: env.get(name.toLowerCase());
		span = { start: sliceStart + toks[0].start, end: sliceStart + toks[0].end };
		// A VARIABLE declared Variant (or with no type) passed ByRef to a typed
		// parameter is the compile error itself (issue #111): the VBE refuses
		// `Take v` with `Dim v As Variant` for `x As Long`, `x As String` and
		// `x As Object` alike (measured 2026-09-26). Only a variable or
		// parameter: a parameterless Function's name here is a call result,
		// which passes as a copy.
		const variantVariable = !param.isArray
			&& declaredType?.resolved
			&& (declaredType.kind === 'localVariable' || declaredType.kind === 'moduleVariable' || declaredType.kind === 'parameter')
			&& (normalizeType(declaredType.asType) ?? 'variant') === 'variant';
		if (variantVariable) {
			return { name, actual: declaredType?.asType ?? 'Variant', span };
		}
	} else if (toks.length >= 4 && toks[1].rawText === '(' && toks[toks.length - 1].rawText === ')' && matchParenFrom(toks, 1) === toks.length - 1) {
		// An element of an array variable passes ByRef as the variable itself
		// does: `Take a(1)` with `Dim a(1) As Variant` for `n As Long` is
		// "ByRef argument type mismatch" (issue #216, measured in Excel 16.0).
		name = tokenName(toks[0]);
		const declaredType = name ? resolveExpressionType?.(name) : undefined;
		if (
			!name || !declaredType?.resolved || !declaredType.isArray
			|| !(declaredType.kind === 'localVariable' || declaredType.kind === 'moduleVariable' || declaredType.kind === 'parameter')
		) {
			return undefined;
		}
		span = { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end };
		const element = normalizeType(declaredType.asType) ?? 'variant';
		if (element === 'variant' && !param.isArray) {
			return { name: `${name}(...)`, actual: 'Variant', span };
		}
		actualRaw = declaredType.asType;
		name = `${name}(...)`;
	} else if (toks.length === 3 && toks[1].rawText === '.') {
		const qualifier = tokenName(toks[0]);
		const member = tokenName(toks[2]);
		if (!qualifier || !member) {
			return undefined;
		}
		const declaredType = resolveQualifiedExpressionType?.(qualifier, member);
		if (!declaredType?.resolved) {
			return undefined;
		}
		name = `${qualifier}.${member}`;
		actualRaw = declaredType.asType;
		span = { start: sliceStart + toks[0].start, end: sliceStart + toks[2].end };
	} else {
		return undefined;
	}
	const actual = normalizeType(actualRaw);
	if (!byRefExact(actual) || sameByRefType(actual, expected)) {
		return undefined;
	}
	// An Object passes ByRef to `c As Collection`, and a Collection to
	// `o As Object`: both compile, and the wrong object raises 13 at run time
	// (issue #343, measured in Excel 16.0).
	if ([actual, expected].includes('object') && [actual, expected].includes('collection')) {
		return undefined;
	}
	return {
		name,
		actual: actualRaw ?? name,
		span,
	};
}

/**
 * A type a ByRef argument must match exactly: a scalar, Object, or a
 * Collection. A Collection into `p As Long`, and a Long, String or Variant
 * into `p As Collection`, are "ByRef argument type mismatch" (issue #410,
 * measured in Excel 16.0).
 */
function byRefExact(type: string | undefined): boolean {
	return isKnownByRefExactType(type) || type === 'collection';
}

/**
 * Whether a ByRef argument's type is the parameter's. LongPtr is LongLong in
 * 64-bit Office, the platform the analyzer assumes: stdVBA passes a LongLong
 * array element to DispCallFunc's `paValues As LongPtr`, and it compiles.
 */
function sameByRefType(actual: string | undefined, expected: string | undefined): boolean {
	const widen = (type: string | undefined): string | undefined => (type === 'longptr' ? 'longlong' : type);
	return widen(actual) === widen(expected);
}

export function isKnownByRefExactType(type: string | undefined): boolean {
	if (!type || type === 'variant') {
		return false;
	}
	return type === 'object' || isKnownScalarType(type);
}

export function namedArgumentSlot(slot: VbaToken[]): { name: string; value: VbaToken[] } | undefined {
	if (!isNamedSlot(slot)) {
		return undefined;
	}
	return {
		name: stripHeaderBrackets(slot[0].rawText),
		value: slot.slice(2),
	};
}

export function callableSignatureFor(
	name: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallableTypeSignature | undefined {
	if (bareCallableSourceShadowed(name, sourceNames)) {
		return undefined;
	}
	const user = moduleSignatures.get(name.toLowerCase());
	if (user) {
		return user;
	}
	if (runtimeCallableSourceShadowed(name, sourceNames)) {
		return undefined;
	}
	// The runtime table leaves out the type names; called, `String(n, c)` is
	// String$'s Variant twin and takes its arguments (issue #332).
	const runtime = resolveRuntimeFunction(name) ?? (name.toLowerCase() === 'string' ? resolveRuntimeFunction('String$') : undefined);
	if (!runtime) {
		return undefined;
	}
	return runtimeTypeSignature(runtime);
}

export function runtimeTypeSignature(runtime: VbaRuntimeFunction): CallableTypeSignature {
	if (runtime.params) {
		return {
			name: runtime.name,
			params: runtime.params.map((p) => ({
				name: p.name,
				type: p.type,
				optional: p.optional ?? false,
				paramArray: p.paramArray ?? false,
				...(p.isArray ? { isArray: true } : {}),
				...(p.nullRaises ? { nullRaises: true } : {}),
			})),
			returnType: runtime.returns,
		};
	}
	return parseRuntimeDisplaySignature(runtime.name, runtime.signature, runtime.returns);
}

export function runtimeAritySignature(runtime: VbaRuntimeFunction): CallableTypeSignature | undefined {
	if (runtime.params || runtimeSignatureParameterText(runtime.signature) !== undefined) {
		return runtimeTypeSignature(runtime);
	}
	return undefined;
}

export function parseRuntimeDisplaySignature(
	name: string,
	signature: string,
	returnType?: string,
): CallableTypeSignature {
	const inner = runtimeSignatureParameterText(signature);
	if (inner === undefined) {
		return { name, params: [], returnType };
	}
	const params = splitSignatureTopLevel(inner)
		.map(parseRuntimeParamType)
		.filter((p): p is CallableParamType => p !== undefined);
	return { name, params, returnType };
}

/**
 * The text of a display signature's parameter list: from its first `(` to the
 * `)` that closes it.
 *
 * Not to the LAST `)`. A signature can go on past its parameter list with a
 * return type that has parentheses of its own - `Values() As Long()`, a
 * Function returning an array - and reading to the last one took `) As Long(`
 * for the parameters: one required parameter named `As`. Every call to such a
 * member with its empty argument list then reported "expected 1 argument",
 * `Bridge.EmptyBytes()` in vbaSQLBridge among them, in code that compiles.
 * A `)` inside a quoted default value closes nothing.
 */
export function runtimeSignatureParameterText(signature: string): string | undefined {
	const open = signature.indexOf('(');
	if (open < 0) {
		return undefined;
	}
	let depth = 0;
	let quoted = false;
	for (let i = open; i < signature.length; i++) {
		const ch = signature[i];
		if (ch === '"') {
			quoted = !quoted;
		} else if (quoted) {
			continue;
		} else if (ch === '(') {
			depth++;
		} else if (ch === ')') {
			depth--;
			if (depth === 0) {
				return signature.slice(open + 1, i);
			}
		}
	}
	return undefined;
}

export function parseRuntimeParamType(raw: string): CallableParamType | undefined {
	let text = raw.trim();
	if (!text) {
		return undefined;
	}
	const optional = text.startsWith('[') && text.endsWith(']');
	text = text.replace(/^\[/, '').replace(/\]$/, '').trim();
	const paramArray = /^ParamArray\b/i.test(text);
	text = text.replace(/^ParamArray\b\s*/i, '');
	text = text.replace(/^(?:ByVal|ByRef)\b\s*/i, '');
	text = text.replace(/\s*=\s*.*$/, '').trim();
	const as = /\bAs\s+([\p{L}_][\p{L}\p{M}\p{N}_]*(?:\(\))?)/iu.exec(text);
	const first = /[\p{L}_][\p{L}\p{M}\p{N}_]*/u.exec(text)?.[0];
	if (!first) {
		return undefined;
	}
	return {
		name: first,
		type: as?.[1],
		optional,
		paramArray,
	};
}

/**
 * A parameter list split at its top-level commas. Quoted text is opaque, the
 * way runtimeSignatureParameterText treats it: a default of `")"` read as a
 * bracket used to leave the depth unbalanced, and every comma after it was
 * taken for the inside of a parameter - `F([s As String = ")"], [n As Long])`
 * came out as one parameter, and a call passing both was reported as passing
 * too many.
 */
export function splitSignatureTopLevel(text: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let start = 0;
	let quoted = false;
	for (let i = 0; i < text.length; i++) {
		const c = text[i];
		if (c === '"') {
			quoted = !quoted;
		} else if (quoted) {
			continue;
		} else if (c === '(' || c === '[') {
			depth++;
		} else if (c === ')' || c === ']') {
			depth--;
		} else if (c === ',' && depth === 0) {
			out.push(text.slice(start, i));
			start = i + 1;
		}
	}
	out.push(text.slice(start));
	return out;
}

export function inferArgumentType(
	slot: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): InferredArgumentType | undefined {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return inferExpressionType(
		toks,
		sliceStart,
		env,
		moduleSignatures,
		sourceNames,
		source,
		memberCtx,
		resolveExpressionType,
		resolveQualifiedExpressionType,
	);
}

/**
 * Runtime functions that return Null for a Null first argument, where the
 * $ spellings raise 94 at the call (issue #184, measured in Excel 16.0).
 */
const NULL_PROPAGATING_RUNTIME_FUNCTIONS: ReadonlySet<string> = new Set([
	'len', 'lenb', 'left', 'right', 'mid', 'ucase', 'lcase', 'trim', 'ltrim', 'rtrim',
]);

export function inferExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): InferredArgumentType | undefined {
	const first = toks[0];
	if (!first) {
		return undefined;
	}
	const unwrapped = unwrapOuterParens(toks);
	if (unwrapped !== toks) {
		return inferExpressionType(
			unwrapped,
			sliceStart,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
	}
	const signedNumericLiteral = inferSignedNumericLiteral(toks, sliceStart);
	if (signedNumericLiteral) {
		return signedNumericLiteral;
	}
	const concatenation = inferStringConcatenationExpressionType(
		toks,
		sliceStart,
		env,
		moduleSignatures,
		sourceNames,
		source,
		memberCtx,
		resolveExpressionType,
		resolveQualifiedExpressionType,
	);
	if (concatenation) {
		return concatenation;
	}
	const arithmetic = inferArithmeticExpressionType(
		toks,
		sliceStart,
		env,
		moduleSignatures,
		sourceNames,
		source,
		memberCtx,
		resolveExpressionType,
		resolveQualifiedExpressionType,
	);
	if (arithmetic) {
		return arithmetic;
	}
	return inferAtomicExpressionType(
		toks,
		sliceStart,
		env,
		moduleSignatures,
		sourceNames,
		source,
		memberCtx,
		resolveExpressionType,
		resolveQualifiedExpressionType,
	);
}

export function inferSignedNumericLiteral(
	toks: VbaToken[],
	sliceStart: number,
): InferredArgumentType | undefined {
	if (toks.length !== 2 || toks[0].kind !== 'operator') {
		return undefined;
	}
	const sign = toks[0].rawText;
	if (sign !== '+' && sign !== '-') {
		return undefined;
	}
	const literal = toks[1];
	if (literal.kind === 'floatLiteral') {
		// `-2147483649#` overflows a Long as its unsigned twin does (issue #223).
		const magnitude = Number(literal.rawText.replace(/[!#@]$/, ''));
		if (!Number.isFinite(magnitude)) {
			return undefined;
		}
		return {
			type: 'Double',
			label: `numeric literal ${sign}${literal.rawText}`,
			span: { start: sliceStart + toks[0].start, end: sliceStart + literal.end },
			numericText: `${sign}${literal.rawText}`,
			floatValue: sign === '-' ? -magnitude : magnitude,
		};
	}
	if (literal.kind !== 'integerLiteral') {
		return undefined;
	}
	const value = parseDecimalIntegerLiteral(literal.rawText);
	if (value === undefined) {
		return undefined;
	}
	const signed = sign === '-' ? -value : value;
	const text = `${sign}${literal.rawText}`;
	return {
		type: 'Double',
		label: `numeric literal ${text}`,
		span: { start: sliceStart + toks[0].start, end: sliceStart + literal.end },
		numericValue: signed,
		numericText: text,
	};
}

export function inferAtomicExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): InferredArgumentType | undefined {
	const first = toks[0];
	if (!first) {
		return undefined;
	}
	const span = { start: sliceStart + first.start, end: sliceStart + first.end };
	if (toks.length === 1) {
		switch (first.kind) {
			case 'stringLiteral': {
				const value = stringLiteralValue(first.rawText);
				return { type: 'String', label: `String literal ${first.rawText}`, span, stringValue: value };
			}
			case 'integerLiteral':
			case 'floatLiteral': {
				const numericValue =
					first.kind === 'integerLiteral'
						? parseDecimalIntegerLiteral(first.rawText)
						: undefined;
				const floatValue = first.kind === 'floatLiteral'
					? Number(first.rawText.replace(/[!#@]$/, ''))
					: undefined;
				return {
					type: 'Double',
					label: `numeric literal ${first.rawText}`,
					span,
					numericValue,
					numericText: first.rawText,
					...(floatValue !== undefined && Number.isFinite(floatValue) ? { floatValue } : {}),
				};
			}
			case 'dateLiteral':
				return { type: 'Date', label: 'Date literal', span };
			case 'keyword': {
				const word = first.rawText.toLowerCase();
				if (word === 'true' || word === 'false') {
					return { type: 'Boolean', label: 'Boolean literal', span };
				}
				if (word === 'nothing') {
					return { type: 'Nothing', label: 'Nothing', span };
				}
				if (word === 'null') {
					return { type: 'Null', label: 'Null', span };
				}
				break;
			}
			default:
				break;
		}
	}
	const name = tokenName(first);
	if (name && toks.length === 1) {
		const declaredType = resolveExpressionType?.(name);
		// A Const that holds a string literal converts as the literal does (issue #255).
		if (declaredType?.resolved && declaredType.stringValue !== undefined) {
			return { type: 'String', label: `constant '${name}' (${JSON.stringify(declaredType.stringValue)})`, span, stringValue: declaredType.stringValue };
		}
		const type = declaredType?.resolved
			? declaredType.asType
			: env.get(name.toLowerCase());
		if (type) {
			return { type, label: `${name} As ${type}`, span };
		}
		const sig = parameterlessValueSignature(name, moduleSignatures, sourceNames);
		if (sig?.returnType) {
			return { type: sig.returnType, label: `${name} As ${sig.returnType}`, span };
		}
		const externalObject = inferBareExternalObjectExpressionType(
			name,
			span,
			sourceNames,
			memberCtx,
		);
		if (externalObject) {
			return externalObject;
		}
		const external = inferBareExternalConstantExpressionType(name, span, sourceNames, memberCtx?.model);
		if (external) {
			return external;
		}
		return undefined;
	}
	if (tokenText(first) === 'new' && toks.length === 2) {
		const typeName = tokenName(toks[1]);
		if (typeName) {
			return {
				type: typeName,
				label: `New ${typeName}`,
				span: { start: sliceStart + toks[1].start, end: sliceStart + toks[1].end },
			};
		}
	}
	if (name) {
		const callName = parenthesizedCallNameAt(toks, 0);
		const errorVariant = callName?.parenIndex === 1
			? inferIntrinsicCverrErrorVariant(
				toks,
				sliceStart,
				moduleSignatures,
				sourceNames,
			)
			: undefined;
		if (errorVariant) {
			return errorVariant;
		}
		// `Len(Null)` and `UCase(Null)` return Null, whatever type they
		// otherwise return (issue #184, measured in Excel 16.0).
		if (
			callName
			&& NULL_PROPAGATING_RUNTIME_FUNCTIONS.has(callName.name.toLowerCase())
			&& !moduleSignatures.has(callName.name.toLowerCase())
			&& !bareCallableSourceShadowed(callName.name, sourceNames)
			&& !runtimeCallableSourceShadowed(callName.name, sourceNames)
			&& matchParenFrom(toks, callName.parenIndex) === toks.length - 1
		) {
			const first = splitArgSlots(toks.slice(callName.parenIndex + 1, -1), sliceStart).slots[0] ?? [];
			if (first.length === 1 && tokenText(first[0]) === 'null') {
				return {
					type: 'Null',
					label: `${callName.name}(Null), which is Null`,
					span: { start: span.start, end: sliceStart + toks[toks.length - 1].end },
				};
			}
		}
		if (callName) {
			const sig = callableSignatureFor(callName.name, moduleSignatures, sourceNames);
			if (sig?.returnType && matchParenFrom(toks, callName.parenIndex) === toks.length - 1) {
				return {
					type: sig.returnType,
					label: `${callName.name}(...) As ${sig.returnType}`,
					span: { start: span.start, end: sliceStart + toks[callName.nameEndIndex].end },
				};
			}
			const hostGlobal = inferBareHostGlobalCallType(toks, callName, sliceStart, moduleSignatures, sourceNames, memberCtx);
			if (hostGlobal) {
				return hostGlobal;
			}
		}
	}
	if (name && toks[1]?.rawText === '.') {
		const member = tokenName(toks[2]);
		const errorVariant = inferIntrinsicCverrErrorVariant(
			toks,
			sliceStart,
			moduleSignatures,
			sourceNames,
		);
		if (errorVariant) {
			return errorVariant;
		}
		if (member && toks.length === 3) {
			const lookupKey = qualifiedProcedureKey(name, member);
			const sig = parameterlessValueSignature(lookupKey, moduleSignatures);
			if (sig?.returnType) {
				return {
					type: sig.returnType,
					label: `${name}.${member} As ${sig.returnType}`,
					span: { start: sliceStart + toks[2].start, end: sliceStart + toks[2].end },
				};
			}
			const declaredType = resolveQualifiedExpressionType?.(name, member);
			if (declaredType?.resolved) {
				return declaredType.asType
					? {
						type: declaredType.asType,
						label: `${name}.${member} As ${declaredType.asType}`,
						span: { start: sliceStart + toks[2].start, end: sliceStart + toks[2].end },
					}
					: undefined;
			}
			const external = inferQualifiedExternalConstantExpressionType(
				name,
				member,
				{ start: sliceStart + toks[2].start, end: sliceStart + toks[2].end },
				memberCtx?.model,
			);
			if (external) {
				return external;
			}
		}
		if (member && toks[3]?.rawText === '(' && matchParenFrom(toks, 3) === toks.length - 1) {
			const lookupKey = qualifiedProcedureKey(name, member);
			const sig = moduleSignatures.get(lookupKey);
			if (sig?.returnType) {
				return {
					type: sig.returnType,
					label: `${name}.${member}(...) As ${sig.returnType}`,
					span: { start: sliceStart + toks[2].start, end: sliceStart + toks[2].end },
				};
			}
		}
	}
	const memberType = source && memberCtx
		? inferMemberExpressionType(source, toks, sliceStart, memberCtx)
		: undefined;
	if (memberType) {
		return memberType;
	}
	return undefined;
}

/**
 * `Range("A1")`, `Cells(1, 1)`, `Names(1)`: a member of the host's hidden
 * Global interface called bare, which the member chains never see because
 * nothing precedes it. Its arguments index what it returns the way any
 * member call's do (issue #202).
 */
function inferBareHostGlobalCallType(
	toks: readonly VbaToken[],
	callName: { name: string; parenIndex: number },
	sliceStart: number,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
	memberCtx: MemberCompletionContext | undefined,
): InferredArgumentType | undefined {
	const lower = callName.name.toLowerCase();
	if (
		!memberCtx
		|| callName.parenIndex !== 1
		|| matchParenFrom(toks, 1) !== toks.length - 1
		|| moduleSignatures.has(lower)
		|| bareCallableSourceShadowed(callName.name, sourceNames)
		|| runtimeCallableSourceShadowed(callName.name, sourceNames)
	) {
		return undefined;
	}
	const member = resolveHostGlobalMember(callName.name, memberCtx.model);
	if (!member?.returns) {
		return undefined;
	}
	const owner = memberCtx.model?.globalType ?? '';
	const type = memberExpressionReturnType({ ...member, owner }, toks.slice(2, -1), memberCtx);
	return {
		type,
		label: `${callName.name}(...) As ${type}`,
		span: { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end },
	};
}

export function inferIntrinsicCverrErrorVariant(
	toks: readonly VbaToken[],
	sliceStart: number,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): InferredArgumentType | undefined {
	const firstName = tokenName(toks[0]);
	if (!firstName) {
		return undefined;
	}
	let parenIndex = -1;
	let displayName = '';
	if (firstName.toLowerCase() === 'cverr' && toks[1]?.rawText === '(') {
		if (
			moduleSignatures.has(firstName.toLowerCase()) ||
			bareCallableSourceShadowed(firstName, sourceNames) ||
			runtimeCallableSourceShadowed(firstName, sourceNames)
		) {
			return undefined;
		}
		parenIndex = 1;
		displayName = firstName;
	} else if (
		firstName.toLowerCase() === 'vba' &&
		toks[1]?.rawText === '.' &&
		tokenName(toks[2])?.toLowerCase() === 'cverr' &&
		toks[3]?.rawText === '('
	) {
		parenIndex = 3;
		displayName = `${firstName}.${toks[2].rawText}`;
	}
	if (parenIndex < 0) {
		return undefined;
	}
	const close = matchParenFrom(toks, parenIndex);
	if (close !== toks.length - 1) {
		return undefined;
	}
	const inner = toks.slice(parenIndex + 1, close);
	if (inner.length === 0) {
		return undefined;
	}
	const split = splitArgSlots(inner, sliceStart);
	if (split.slots.length !== 1 || split.slots[0].length === 0) {
		return undefined;
	}
	return {
		type: 'Error',
		label: `${displayName}(...) Error Variant`,
		span: spanForTokens(toks, sliceStart),
	};
}

export function inferMemberExpressionType(
	source: string,
	toks: VbaToken[],
	sliceStart: number,
	memberCtx: MemberCompletionContext,
): InferredArgumentType | undefined {
	if (hasTopLevelOperator(toks)) {
		return undefined;
	}
	const resolved = finalMemberTokenInExpression(toks);
	if (!resolved) {
		return undefined;
	}
	const member = resolveExactMemberCompletion(
		source,
		resolved.name,
		sliceStart + resolved.token.end,
		memberCtx,
	);
	if (!member?.returns) {
		return undefined;
	}
	if (!resolved.called && member.kind === 'method' && !memberAcceptsZeroArguments(member)) {
		return undefined;
	}
	const returnType = memberExpressionReturnType(member, resolved.argumentTokens, memberCtx);
	const labelStart = toks[0]?.start ?? resolved.token.start;
	const labelEnd = resolved.called ? toks[toks.length - 1].end : resolved.token.end;
	const labelText = source.slice(sliceStart + labelStart, sliceStart + labelEnd).trim();
	return {
		type: returnType,
		label: `${labelText} As ${returnType}`,
		span: { start: sliceStart + resolved.token.start, end: sliceStart + resolved.token.end },
	};
}

export function memberExpressionReturnType(
	member: MemberCompletion,
	argumentTokens: readonly VbaToken[] | undefined,
	memberCtx: MemberCompletionContext,
): string {
	// Calling a member with arguments indexes into it. When the member returns a
	// host collection (one whose Item resolves to an element type), the call
	// yields that element - e.g. ws.ChartObjects(1) is a ChartObject, pt.PivotFields(1)
	// is a PivotField - not the collection itself. This holds regardless of how the
	// accessor is modelled (method- or property-kind, signed or not), so it covers
	// the whole `Collection([Index])` family. defaultHostItemReturnType returns
	// undefined for any non-collection return, so a concrete-typed call keeps its
	// declared type: Application.Intersect(a, b) stays Range, Workbooks.Add(t) stays
	// Workbook, ws.Range("A1") stays Range. Item/_Default/Add are excluded because
	// they already return the resolved element/result (see isExplicitElementAccessor).
	// A member that takes an argument of its own is what it returns:
	// Shapes.Range(Array("A")) is a ShapeRange, not a Shape (issue #197).
	if (
		member.returns &&
		argumentTokens &&
		argumentTokens.length > 0 &&
		!isExplicitElementAccessor(member.name) &&
		!memberTakesOwnArguments(member.signature)
	) {
		// VBA's Collection holds Variants: `acc.children(i)` may be anything.
		if (normalizeType(member.returns) === 'collection') {
			return 'Variant';
		}
		return defaultHostItemReturnType(member.returns, memberCtx) ?? member.returns;
	}
	return member.returns ?? 'Variant';
}

export function defaultHostItemReturnType(
	typeName: string,
	memberCtx: MemberCompletionContext,
): string | undefined {
	const members = getHostMembers(typeName, memberCtx.model);
	const item = members.find((member) => member.name.toLowerCase() === 'item');
	if (item?.returns) {
		// The library declares most Item accessors `As Object` and the model
		// repairs the type from the reference prose, which is right for
		// completion and chaining. It is not a compile-time binding: the VBE
		// compiles `Worksheets(1).NoSuchMember` and `Workbooks(1).NoSuchMember`
		// (measured in Excel 16.0, issue #114), so the item's members are late
		// bound. A one-part union carries the type without closing it. The
		// hand-written collections carry the repaired type on Item, so the
		// library's word is read off `_Default` too.
		const defaultMember = members.find((member) => member.name === '_Default');
		const declaredObject = [item, defaultMember].some((member) =>
			member?.declaredType === 'Object' || /\bAs Object\s*$/i.test(member?.signature ?? ''));
		return declaredObject ? `union:${item.returns}` : item.returns;
	}
	// A mixed-element collection - e.g. Sheets, whose Item is a Worksheet OR a
	// Chart - carries `returnsAnyOf` instead of a single `returns`. Its indexed
	// element is a late-bound Object in VBA, so resolve to Object: a generic
	// object is assignable to any specific object target, which avoids a false
	// assignment-object-type-mismatch on `Set ws = ThisWorkbook.Sheets("x")`,
	// while single-typed collections (Worksheets -> Worksheet) stay strict.
	if (item?.returnsAnyOf?.length) {
		return 'Object';
	}
	return undefined;
}

export function finalMemberTokenInExpression(
	toks: readonly VbaToken[],
): { name: string; token: VbaToken; called: boolean; argumentTokens?: readonly VbaToken[] } | undefined {
	const last = toks[toks.length - 1];
	if (!last) {
		return undefined;
	}
	if (tokenName(last) && toks[toks.length - 2]?.rawText === '.') {
		return { name: tokenName(last)!, token: last, called: false };
	}
	if (last.rawText !== ')') {
		return undefined;
	}
	const open = matchingOpenParenIndex(toks, toks.length - 1);
	if (open < 2) {
		return undefined;
	}
	const member = toks[open - 1];
	if (!tokenName(member) || toks[open - 2]?.rawText !== '.') {
		return undefined;
	}
	return {
		name: tokenName(member)!,
		token: member,
		called: true,
		argumentTokens: toks.slice(open + 1, -1),
	};
}

export function matchingOpenParenIndex(toks: readonly VbaToken[], close: number): number {
	let depth = 0;
	for (let i = close; i >= 0; i--) {
		const raw = toks[i].rawText;
		if (raw === ')') {
			depth++;
		} else if (raw === '(') {
			depth--;
			if (depth === 0) {
				return i;
			}
		}
	}
	return -1;
}

export function hasTopLevelOperator(toks: readonly VbaToken[]): boolean {
	let depth = 0;
	for (const tok of toks) {
		const raw = tok.rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (depth === 0 && tok.kind === 'operator') {
			return true;
		}
	}
	return false;
}

export function memberAcceptsZeroArguments(member: MemberCompletion): boolean {
	if (!member.signature) {
		return false;
	}
	return callableAcceptsZeroArguments(parseRuntimeDisplaySignature(member.name, member.signature));
}

export function parameterlessValueSignature(
	name: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallableTypeSignature | undefined {
	const sig = callableSignatureFor(name, moduleSignatures, sourceNames);
	return sig?.returnType && callableAcceptsZeroArguments(sig) ? sig : undefined;
}

export function unwrapOuterParens(toks: VbaToken[]): VbaToken[] {
	if (toks.length < 2 || toks[0].rawText !== '(') {
		return toks;
	}
	const close = matchParenFrom(toks, 0);
	return close === toks.length - 1 ? toks.slice(1, -1) : toks;
}

export function inferArithmeticExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): InferredArgumentType | undefined {
	const parts = splitTopLevelArithmeticOperands(toks);
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const inferred = inferExpressionType(
			part,
			sliceStart,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
		const normalized = normalizeType(inferred?.type);
		if (!normalized || !isNumericType(normalized)) {
			return undefined;
		}
	}
	return {
		type: 'Double',
		label: 'numeric expression',
		span: spanForTokens(toks, sliceStart),
	};
}

export function nonnumericStringArithmeticOperand(
	expectedRaw: string,
	slot: VbaToken[],
	sliceStart: number,
): InferredArgumentType | undefined {
	const expected = normalizeType(expectedRaw);
	if (!expected || !isNumericType(expected)) {
		return undefined;
	}
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return findNonnumericStringInArithmeticExpression(toks, sliceStart);
}

export function findNonnumericStringInArithmeticExpression(
	toks: VbaToken[],
	sliceStart: number,
): InferredArgumentType | undefined {
	const unwrapped = unwrapOuterParens(toks);
	if (unwrapped !== toks) {
		return findNonnumericStringInArithmeticExpression(unwrapped, sliceStart);
	}
	const parts = splitTopLevelArithmeticOperands(toks);
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const nested = findNonnumericStringInArithmeticExpression(part, sliceStart);
		if (nested) {
			return nested;
		}
		const operand = unwrapOuterParens(part);
		if (operand.length === 1 && operand[0].kind === 'stringLiteral') {
			const value = stringLiteralValue(operand[0].rawText);
			if (isProvablyNonNumericString(value)) {
				return {
					type: 'String',
					label: `nonnumeric string literal ${operand[0].rawText}`,
					span: { start: sliceStart + operand[0].start, end: sliceStart + operand[0].end },
					stringValue: value,
				};
			}
		}
	}
	return undefined;
}

export function inferStringConcatenationExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
	resolveExpressionType?: SourceDeclaredTypeResolver,
	resolveQualifiedExpressionType?: SourceQualifiedDeclaredTypeResolver,
): InferredArgumentType | undefined {
	const parts = splitTopLevelOperands(toks, '&');
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const inferred = inferExpressionType(
			part,
			sliceStart,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
			resolveExpressionType,
			resolveQualifiedExpressionType,
		);
		const normalized = normalizeType(inferred?.type);
		if (!normalized || !isStringConcatenationOperandType(normalized)) {
			return undefined;
		}
	}
	return {
		type: 'String',
		label: 'string concatenation expression',
		span: spanForTokens(toks, sliceStart),
	};
}

export function splitTopLevelArithmeticOperands(toks: VbaToken[]): VbaToken[][] {
	const parts = splitTopLevelOperands(toks, '+', '-', '*', '/', '\\', '^');
	if (parts.length < 2) {
		return [];
	}
	return parts;
}

export function splitTopLevelOperands(toks: VbaToken[], ...operators: string[]): VbaToken[][] {
	const allowed = new Set(operators);
	const parts: VbaToken[][] = [];
	let start = 0;
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
			continue;
		}
		if (raw === ')' || raw === ']') {
			depth--;
			continue;
		}
		if (depth !== 0) {
			continue;
		}
		if (toks[i].kind !== 'operator' || !allowed.has(toks[i].rawText)) {
			if (toks[i].kind === 'operator') {
				return [];
			}
			continue;
		}
		// A +/- at the start of an operand is a unary sign (e.g. `2 * -3`,
		// `x + -1`); fold it into the following operand rather than treating it
		// as a separator. Any other operator at the operand start is malformed.
		if (i === start) {
			if (toks[i].rawText === '+' || toks[i].rawText === '-') {
				continue;
			}
			return [];
		}
		if (i === toks.length - 1) {
			return [];
		}
		parts.push(toks.slice(start, i));
		start = i + 1;
	}
	if (parts.length === 0) {
		return [];
	}
	parts.push(toks.slice(start));
	return parts;
}

export function isStringConcatenationOperandType(type: string): boolean {
	return (
		type === 'string' ||
		type === 'boolean' ||
		type === 'date' ||
		isNumericType(type)
	);
}

export function spanForTokens(toks: readonly VbaToken[], sliceStart: number): Span {
	const first = toks[0];
	const last = toks[toks.length - 1];
	return { start: sliceStart + first.start, end: sliceStart + last.end };
}

export function incompatibilityReason(
	expectedRaw: string,
	actual: InferredArgumentType,
): string | undefined {
	const expected = normalizeType(expectedRaw);
	const actualType = normalizeType(actual.type);
	if (!expected || !actualType || expected === 'variant' || actualType === 'variant') {
		return undefined;
	}
	if (actualType === 'error' && isKnownScalarType(expected)) {
		return "An Error Variant cannot be coerced to this scalar type. This will raise Run-time error '13': Type mismatch.";
	}
	if (actualType === 'null' && isKnownScalarType(expected)) {
		return "Null cannot be coerced to this scalar type. This will raise Run-time error '94': Invalid use of Null.";
	}
	if (expected === 'object') {
		return actualType === 'nothing' || actualType === 'object' || !isKnownScalarType(actualType)
			? undefined
			: 'An object parameter requires an object value.';
	}
	if (isNumericType(expected)) {
		const overflow = numericLiteralOverflowReason(expected, actual);
		if (overflow) {
			return overflow;
		}
		if (isNumericType(actualType) || actualType === 'boolean') {
			return undefined;
		}
		if (actualType === 'string' && actual.stringValue !== undefined) {
			// The string's number where every locale reads it alike: "&H10000"
			// is 65536 and overflows an Integer (issue #188).
			const verdict = numericStringVerdict(actual.stringValue);
			if (verdict.kind === 'invalid') {
				return actual.heldBy !== undefined
					? `'${actual.heldBy}' holds ${JSON.stringify(actual.stringValue)} here, which converts to no number. This will raise Run-time error '13': Type mismatch.`
					: "This string literal cannot be converted to a numeric value. This will raise Run-time error '13': Type mismatch.";
			}
			const bounds = verdict.value === undefined ? undefined : numericLiteralBounds(expected);
			if (bounds && (verdict.value! < bounds.min || verdict.value! > bounds.max)) {
				return `The string ${JSON.stringify(actual.stringValue)} converts to ${verdict.value}, outside the ${bounds.label} range ${bounds.min} to ${bounds.max}. This will raise Run-time error '6': Overflow.`;
			}
		}
		return undefined;
	}
	if (expected === 'boolean') {
		if (actualType === 'boolean' || isNumericType(actualType)) {
			return undefined;
		}
		// A String whose value is not known may be "True" or "5", which
		// convert (measured in Excel 16.0).
		if (actualType === 'string' && actual.stringValue !== undefined && isInvalidBooleanString(actual.stringValue)) {
			return "This string literal cannot be converted to Boolean. This will raise Run-time error '13': Type mismatch.";
		}
		return undefined;
	}
	if (expected === 'date' && actualType === 'string' && actual.stringValue !== undefined) {
		return isInvalidDateString(actual.stringValue)
			? "This string literal cannot be converted to a Date. This will raise Run-time error '13': Type mismatch."
			: undefined;
	}
	if (expected === 'string') {
		return undefined; // VBA can stringify scalar values; do not warn.
	}
	return undefined;
}

export function numericLiteralOverflowReason(
	expected: string,
	actual: InferredArgumentType,
): string | undefined {
	if (actual.numericValue === undefined) {
		return floatLiteralOverflowReason(expected, actual);
	}
	const bounds = numericLiteralBounds(expected);
	if (!bounds) {
		return undefined;
	}
	if (actual.numericValue >= bounds.min && actual.numericValue <= bounds.max) {
		return undefined;
	}
	// A resolved named constant must not be described as a "numeric literal"; name
	// the constant and show its value instead.
	if (actual.numericConstantName !== undefined) {
		return `The value of constant '${actual.numericConstantName}' (${actual.numericValue}) is outside the ${bounds.label} range ${bounds.min} to ${bounds.max}. This will raise Run-time error '6': Overflow.`;
	}
	const literal = actual.numericText ?? String(actual.numericValue);
	return `${heldOrLiteral(actual, literal)} is outside the ${bounds.label} range ${bounds.min} to ${bounds.max}. This will raise Run-time error '6': Overflow.`;
}

/**
 * A float literal passed to a Byte, Integer or Long: VBA rounds it half to
 * even and raises 6 when the result is out of range. `EchoL(3000000000#)`
 * overflows, `EchoI(32767.4)` runs (issue #203, measured in Excel 16.0).
 * Currency keeps four decimal places and is left out, as for whole numbers.
 */
function floatLiteralOverflowReason(expected: string, actual: InferredArgumentType): string | undefined {
	if (actual.floatValue !== undefined && (expected === 'longlong' || expected === 'longptr')) {
		// Past LongLong's range a LongPtr overflows too, whatever its width
		// (issue #232): `Take(1E+19)` with ByVal v As LongLong raises 6.
		const rounded = bankersRound(actual.floatValue);
		if (rounded < 2 ** 63 && rounded >= -(2 ** 63)) {
			return undefined;
		}
		const label = expected === 'longptr' ? 'LongPtr' : 'LongLong';
		return `${heldOrLiteral(actual, String(actual.numericText ?? actual.floatValue))} is outside the ${label} range${expected === 'longptr' ? ', at most' : ''} -9223372036854775808 to 9223372036854775807. This will raise Run-time error '6': Overflow.`;
	}
	if (actual.floatValue === undefined || (expected !== 'byte' && expected !== 'integer' && expected !== 'long')) {
		return undefined;
	}
	const bounds = numericLiteralBounds(expected)!;
	const rounded = bankersRound(actual.floatValue);
	if (rounded >= bounds.min && rounded <= bounds.max) {
		return undefined;
	}
	const shown = rounded === actual.floatValue ? '' : `, which VBA rounds to ${rounded},`;
	return `${heldOrLiteral(actual, String(actual.numericText ?? actual.floatValue))}${shown} is outside the ${bounds.label} range ${bounds.min} to ${bounds.max}. This will raise Run-time error '6': Overflow.`;
}

/** "The numeric literal 5", or "The value 5 that 'a' holds here" for a local's known value. */
function heldOrLiteral(actual: InferredArgumentType, text: string): string {
	return actual.heldBy !== undefined ? `The value ${text} that '${actual.heldBy}' holds here` : `The numeric literal ${text}`;
}

export function numericLiteralBounds(
	expected: string,
): { min: number; max: number; label: string } | undefined {
	switch (expected) {
		case 'byte':
			return { min: 0, max: 255, label: 'Byte' };
		case 'integer':
			return { min: -32768, max: 32767, label: 'Integer' };
		case 'long':
			// VBE oracle: a decimal integer literal outside ±2^31 compiles (typed
			// as Double) then narrows to Long, raising Run-time error '6': Overflow.
			// Only bare decimal literals within JS safe-integer range reach here
			// (hex/octal/suffixed/float literals leave numericValue undefined), so
			// every value tested is exactly representable - no boundary false
			// positives. LongLong/LongPtr are intentionally omitted: any safe-integer
			// literal already fits ±2^63, and LongPtr width is platform-dependent.
			return { min: -2147483648, max: 2147483647, label: 'Long' };
		case 'currency':
			// VBE oracle (currency_*_literal_runtime): a bare whole-number decimal
			// literal outside Currency's range compiles (typed as Double) then
			// narrows to Currency, raising Run-time error '6': Overflow. The
			// fractional limits -922337203685477.5808 / +922337203685477.5807 both
			// round inward to the same whole-number magnitude, so the integer
			// boundary is symmetric; 922337203685477 is accepted and 922337203685478
			// overflows on both signs. Every reachable literal is a safe integer and
			// thus an exact IEEE-754 double, so the range check cannot disagree with
			// VBE - no boundary false positives. (Fractional/@-suffixed Currency
			// literals are floatLiteral tokens with no numericValue, so they never
			// reach this entry; their overflow is intentionally out of scope.)
			return { min: -922337203685477, max: 922337203685477, label: 'Currency' };
		default:
			return undefined;
	}
}

export function normalizeType(type: string | undefined): string | undefined {
	if (!type) {
		return undefined;
	}
	return type
		.replace(/\s*\(\s*\)\s*$/, '')
		.replace(/^vb/i, '')
		.trim()
		.toLowerCase();
}

const NUMERIC_TYPES: ReadonlySet<string> = new Set([
	'byte', 'integer', 'long', 'longlong', 'longptr', 'single', 'double', 'currency', 'decimal',
]);

export function isNumericType(type: string): boolean {
	return NUMERIC_TYPES.has(type);
}

export function isKnownScalarType(type: string): boolean {
	return type === 'string' || type === 'boolean' || type === 'date' || isNumericType(type);
}

export function isKnownObjectAssignmentType(
	type: string | undefined,
	memberCtx: MemberCompletionContext,
	resolveType?: ReturnType<typeof createObjectAssignmentTypeResolver>,
): boolean {
	return (resolveType ? resolveType(type) : resolveKnownObjectAssignmentType(type, memberCtx)) !== undefined;
}

export type KnownObjectAssignmentType =
	// `generic` short-circuits compatibility in both directions, which is what
	// the untyped `Object` needs and what keeps `Collection` from inventing
	// mismatch errors while still requiring `Set`.
	| { kind: 'generic'; display: string; key: 'object' | 'collection' }
	| { kind: 'host'; display: string; key: string }
	| { kind: 'project'; display: string; key: string; implements: readonly string[] };

/**
 * Resolve object types against metadata stable for one analysis pass. Project
 * indexing stays lazy behind the normal host/library resolution priority.
 * A fresh resolver observes metadata changes on the next pass.
 */
export function createObjectAssignmentTypeResolver(
	memberCtx: MemberCompletionContext,
): (type: string | undefined) => KnownObjectAssignmentType | undefined {
	let queries: Map<string | undefined, KnownObjectAssignmentType | undefined> | undefined;
	let projectTypes: Map<string, VbaProjectClassMembers | undefined> | undefined;
	const projectType = (lower: string): VbaProjectClassMembers | undefined => {
		if (!projectTypes) {
			projectTypes = new Map();
			for (const candidate of memberCtx.projectClassMembers ?? []) {
				if (candidate.kind === 'userType' || candidate.kind === 'enum' || candidate.kind === 'standardModule') {
					continue;
				}
				const name = candidate.name.toLowerCase();
				// Every eligible duplicate makes the name ambiguous, even the same object twice.
				projectTypes.set(name, projectTypes.has(name) ? undefined : candidate);
			}
		}
		return projectTypes.get(lower);
	};
	return (type) => {
		queries ??= new Map();
		// Raw keys preserve the display spelling used for generic and host types.
		if (!queries.has(type)) {
			queries.set(type, resolveKnownObjectAssignmentType(type, memberCtx, projectType));
		}
		return queries.get(type);
	};
}

export function resolveKnownObjectAssignmentType(
	type: string | undefined,
	memberCtx: MemberCompletionContext,
	projectTypeLookup?: (lower: string) => VbaProjectClassMembers | undefined,
): KnownObjectAssignmentType | undefined {
	if (!type) {
		return undefined;
	}
	const normalized = normalizeType(type);
	if (!normalized || normalized === 'variant') {
		return undefined;
	}
	if (normalized === 'object') {
		return { kind: 'generic', display: type, key: 'object' };
	}
	if (isKnownScalarType(normalized)) {
		return undefined;
	}
	// VBA's own creatable class. It belongs to no host model and to no project,
	// so neither lookup below reaches it, and `Dim c As Collection : c = ...`
	// was reported clean while refusing to compile.
	if (normalized === 'collection') {
		return { kind: 'generic', display: type, key: 'collection' };
	}
	const host = resolveHostAlias(type, memberCtx.model);
	if (host) {
		return { kind: 'host', display: type, key: host.toLowerCase() };
	}
	const library = libraryObjectType(type);
	if (library) {
		return { kind: 'host', display: type, key: library.toLowerCase() };
	}
	const simple = simpleTypeNameForAssignment(type);
	if (!simple) {
		return undefined;
	}
	const lower = simple.toLowerCase();
	const matches = projectTypeLookup ? undefined : (memberCtx.projectClassMembers ?? []).filter(
		(projectType) =>
			// userType and enum are VALUE types - `Dim c As Corner` is a Long,
			// not an object - so neither can make an assignment require Set.
			projectType.kind !== 'userType' &&
			projectType.kind !== 'enum' &&
			projectType.kind !== 'standardModule' &&
			projectType.name.toLowerCase() === lower,
	);
	const match = projectTypeLookup ? projectTypeLookup(lower) : matches?.length === 1 ? matches[0] : undefined;
	if (!match) {
		return undefined;
	}
	return {
		kind: 'project',
		display: match.name,
		key: lower,
		implements: match.implements ?? [],
	};
}

interface ObjectDefaultTypeQueries {
	resolveType: ReturnType<typeof createObjectAssignmentTypeResolver>;
	projectTypeNamed: (key: string) => VbaProjectClassMembers | undefined;
}

/** Lazy default-member facts for one public query, preserving first-surface lookup. */
export function createObjectDefaultQueries(memberCtx: MemberCompletionContext) {
	const resolveType = createObjectAssignmentTypeResolver(memberCtx);
	const projectLookup = (kind?: 'class') => {
		const projectTypes = new Map<string, VbaProjectClassMembers>();
		let projectIndex = 0;
		return (key: string): VbaProjectClassMembers | undefined => {
			if (projectTypes.has(key)) { return projectTypes.get(key); }
			const types = memberCtx.projectClassMembers ?? [];
			while (projectIndex < types.length) {
				const type = types[projectIndex++];
				if (kind && type.kind !== kind) { continue; }
				const lower = type.name.toLowerCase();
				if (!projectTypes.has(lower)) { projectTypes.set(lower, type); }
				if (lower === key) { return projectTypes.get(key); }
			}
			return undefined;
		};
	};
	const projectTypeNamed = projectLookup(), projectClassNamed = projectLookup('class');
	const types = { resolveType, projectTypeNamed };
	const verdicts = new Map<string | undefined, ReturnType<typeof objectLetAssignmentVerdict>>();
	const verdictFor = (type: string | undefined): ReturnType<typeof objectLetAssignmentVerdict> => {
		let verdict = verdicts.get(type);
		if (verdict === undefined) { verdict = objectLetAssignmentVerdict(type, memberCtx, types); verdicts.set(type, verdict); }
		return verdict;
	};
	const indexes = new Map<string | undefined, boolean>();
	const needsIndex = (type: string | undefined): boolean => {
		let answer = indexes.get(type);
		if (answer === undefined) { answer = objectValueNeedsIndex(type, memberCtx, verdictFor); indexes.set(type, answer); }
		return answer;
	};
	return { resolveType, projectTypeNamed, projectClassNamed, verdictFor, needsIndex };
}

/** A getter can supply an object whose default property receives a Let. */
export function getterMayReturnObject(type: string | undefined, memberCtx: MemberCompletionContext): boolean {
	const normalized = normalizeType(type);
	return !normalized || normalized === 'variant' || normalized === 'object'
		|| isKnownObjectAssignmentType(type, memberCtx);
}

/**
 * What a bare `name = value` does to a variable of a known object type
 * (issue #107, each case measured in Excel 16.0). The VBE compiles it as a
 * Let through the type's default member, so it is never "Set required" at
 * compile time:
 *
 *  - `lets`: the type has a parameterless default member (Range's `_Default`
 *    is Value; a project class marks one with VB_UserMemId = 0), or is the
 *    generic Object, whose default member is looked up when it runs. `r = 5`
 *    writes A1. Nothing to report.
 *  - `argument`: the default member takes an argument, so the VBE refuses the
 *    statement: `c = 5` on a Collection is "Argument not optional".
 *  - `noDefault`: the type is fully known and has no default member, so the
 *    statement compiles and raises error 438 when it runs (`ws = 9`).
 *  - `unknown`: the model cannot say. Nothing is reported.
 */
export function objectLetAssignmentVerdict(
	expectedRaw: string | undefined,
	memberCtx: MemberCompletionContext,
	queries?: ObjectDefaultTypeQueries,
): 'lets' | 'argument' | 'noDefault' | 'unknown' {
	const expected = queries ? queries.resolveType(expectedRaw) : resolveKnownObjectAssignmentType(expectedRaw, memberCtx);
	if (!expected) {
		return 'unknown';
	}
	if (expected.kind === 'generic') {
		return expected.key === 'collection' ? 'argument' : 'lets';
	}
	if (expected.kind === 'project') {
		const projectType = queries ? queries.projectTypeNamed(expected.key) : (memberCtx.projectClassMembers ?? []).find(
			(candidate) => candidate.name.toLowerCase() === expected.key,
		);
		if (!projectType || projectType.exhaustive !== true) {
			return 'unknown';
		}
		const defaultMember = projectType.members.find((member) => member.defaultMember);
		if (!defaultMember) {
			return 'noDefault';
		}
		return defaultMember.signature && /\([^)]/.test(defaultMember.signature) ? 'argument' : 'lets';
	}
	const resolved = resolveHostAlias(expectedRaw ?? '', memberCtx.model) ?? libraryObjectType(expectedRaw) ?? expectedRaw ?? '';
	const libraryDefault = libraryDefaultVerdict(resolved);
	if (libraryDefault) {
		return libraryDefault;
	}
	const defaultMember = hostDefaultMember(resolved, memberCtx);
	if (defaultMember) {
		if (defaultMember.kind === 'method' || /\([^)]/.test(defaultMember.signature ?? '')) {
			return 'argument';
		}
		// The model keeps no parameters for a default property. One typed as
		// an element of the collection is its Item and takes the index:
		// Hyperlinks, Areas, Borders, Windows, Workbooks. One typed as a value
		// (Range, Style, Application) gives it. One typed Object is not
		// judged: Worksheets read as a value raises 13, Sheets 450 (issue #221,
		// measured in Excel 16.0).
		const declared = normalizeType(defaultMember.declaredType);
		if (declared === 'object') {
			return 'unknown';
		}
		return defaultMember.returns ? 'argument' : 'lets';
	}
	return hostTypeHasNoDefault(resolved, memberCtx) ? 'noDefault' : 'unknown';
}

/**
 * The verdict from a Word, PowerPoint or Access type's default member as its
 * type library gives it (DISPID 0, issue #438): a Range's Text takes a Let
 * and gives a value; a collection's Item needs an index; a Paragraph's Range
 * gives what a Range gives. Undefined where the table has no entry.
 */
function libraryDefaultVerdict(qualified: string, depth = 0): 'lets' | 'argument' | undefined {
	const found = HOST_DEFAULT_MEMBERS[qualified];
	if (!found) {
		return undefined;
	}
	if (found.kind === 'method' || found.required > 0) {
		return 'argument';
	}
	return depth < 4 && HOST_DEFAULT_MEMBERS[found.returns] ? libraryDefaultVerdict(found.returns, depth + 1) : 'lets';
}

let libraryTypesByLower: Map<string, string> | undefined;

/**
 * A DAO type named with its library, `DAO.Recordset`, as the default-member
 * table keys it. DAO has no host model, so this is how a variable of a DAO
 * type is known to hold an object whose default member the table gives
 * (issue #464, measured in Access 16.0).
 */
function libraryObjectType(type: string | undefined): string | undefined {
	if (!type || !/^dao\./i.test(type.trim())) {
		return undefined;
	}
	libraryTypesByLower ??= new Map(Object.keys(HOST_DEFAULT_MEMBERS).filter((key) => /^DAO\./.test(key)).map((key) => [key.toLowerCase(), key]));
	return libraryTypesByLower.get(type.trim().toLowerCase());
}

/**
 * The run-time error DAO raises where an object of this type is read whole
 * as a value, `v = rs`: its default member, or the one that holds, needs an
 * index DAO checks for itself. Measured in Access 16.0 with the database
 * held (issue #464); a type not measured is undefined.
 */
const DAO_WHOLE_VALUE_ERRORS: Readonly<Record<string, string>> = {
	'DAO.Database': `'3001': Invalid argument`,
	'DAO.Fields': `'3001': Invalid argument`,
	'DAO.Properties': `'3001': Invalid argument`,
	'DAO.QueryDefs': `'3001': Invalid argument`,
	'DAO.Recordset': `'3001': Invalid argument`,
	'DAO.Recordset2': `'3001': Invalid argument`,
	'DAO.TableDef': `'450': Wrong number of arguments or invalid property assignment`,
	'DAO.TableDefs': `'3001': Invalid argument`,
	'DAO.Workspace': `'3001': Invalid argument`,
};

export function daoWholeValueError(type: string | undefined): string | undefined {
	const library = libraryObjectType(type);
	return library ? DAO_WHOLE_VALUE_ERRORS[library] : undefined;
}

/**
 * A Word, PowerPoint or Access type whose default member is a property that
 * holds an object, as a Paragraph's Range does: its name and type. Read whole
 * into a Variant it gives that object's value, but the VBE refuses a Let to it
 * ("Invalid use of property"), an operator on it and a Let of it into a typed
 * value ("Type mismatch") while compiling (issue #462, measured in Word 16.0).
 */
export function objectHoldingDefault(type: string | undefined, memberCtx: MemberCompletionContext): { name: string; returns: string } | undefined {
	const resolved = resolveHostAlias(type ?? '', memberCtx.model) ?? libraryObjectType(type) ?? type ?? '';
	const found = HOST_DEFAULT_MEMBERS[resolved];
	return found && found.kind === 'property' && found.required === 0 && HOST_DEFAULT_MEMBERS[found.returns]
		? { name: found.name, returns: found.returns }
		: undefined;
}

/**
 * A Word, PowerPoint or Access type whose default member is a property no
 * Let reaches, as a Document's Name: `x = 5` does not compile, "Invalid use
 * of property" (issue #438, measured in Word 16.0). One holding an object is
 * {@link objectHoldingDefault}'s.
 */
export function readOnlyHostDefault(type: string | undefined, memberCtx: MemberCompletionContext): string | undefined {
	const resolved = resolveHostAlias(type ?? '', memberCtx.model) ?? libraryObjectType(type) ?? type ?? '';
	const found = HOST_DEFAULT_MEMBERS[resolved];
	return found && found.kind === 'property' && found.required === 0 && !found.writable && !HOST_DEFAULT_MEMBERS[found.returns] ? found.name : undefined;
}

/**
 * The default members a Word, PowerPoint or Access type reaches through,
 * `Range.Text` for a Paragraph, when none of them takes an argument: then
 * `x(1)` does not compile, "Wrong number of arguments or invalid property
 * assignment" (issue #438, measured in Word 16.0).
 */
export function argumentlessHostDefault(type: string | undefined, memberCtx: MemberCompletionContext): string | undefined {
	let key = resolveHostAlias(type ?? '', memberCtx.model) ?? libraryObjectType(type) ?? type ?? '';
	const names: string[] = [];
	for (let depth = 0; depth < 4; depth++) {
		const found = HOST_DEFAULT_MEMBERS[key];
		if (!found || found.kind !== 'property' || found.params > 0) {
			return undefined;
		}
		names.push(found.name);
		if (!HOST_DEFAULT_MEMBERS[found.returns]) {
			// A Variant or an object it gives may still take an index.
			const returns = normalizeType(found.returns);
			return returns !== undefined && returns !== 'variant' && isKnownScalarType(returns) ? names.join('.') : undefined;
		}
		key = found.returns;
	}
	return undefined;
}

/** A host type's default member (DISPID 0, `_Default` in the model), if any. */
function hostDefaultMember(qualified: string, memberCtx: MemberCompletionContext): HostMember | undefined {
	return getHostMembers(qualified, memberCtx.model).find((member) => member.name === '_Default');
}

/**
 * Whether a host type provably has no default member: its member list is
 * complete (hidden members included), and either the type library resolves
 * members while compiling or the type is Excel's. Excel's open types raise
 * 438 read as a value just as its closed ones do: Workbook, Font, Interior,
 * Validation, Window, PageSetup, Border, Shape, Hyperlink (issue #221,
 * measured in Excel 16.0). The other hosts keep the closed-type test.
 */
function hostTypeHasNoDefault(resolved: string, memberCtx: MemberCompletionContext): boolean {
	if (getHostType(resolved, memberCtx.model)?.exhaustive !== true) {
		return false;
	}
	return hostTypeResolvesWhenCompiling(resolved) || /^excel\./i.test(resolved);
}

/**
 * The run-time error reading an object of this type as a value raises where
 * its default member needs an index, or undefined where that is not known:
 * 450 for a Collection, a host default property typed as an element
 * (Hyperlinks), or a host default method with a required parameter (Shapes)
 * (issue #221, measured in Excel 16.0). Names, whose default method takes
 * only optional parameters, raises 449 and is not judged.
 */
export function objectValueNeedsIndex(type: string | undefined, memberCtx: MemberCompletionContext, verdictFor?: (type: string | undefined) => ReturnType<typeof objectLetAssignmentVerdict>): boolean {
	if (normalizeType(type) === 'collection') {
		return true;
	}
	if ((verdictFor ? verdictFor(type) : objectLetAssignmentVerdict(type, memberCtx)) !== 'argument') {
		return false;
	}
	const resolved = resolveHostAlias(type ?? '', memberCtx.model) ?? libraryObjectType(type);
	// Word's Paragraphs and Tables, PowerPoint's Slides: Item(Index) raises
	// 450 read as a value (issue #462, measured in Word and PowerPoint 16.0).
	const library = resolved ? HOST_DEFAULT_MEMBERS[resolved] : undefined;
	if (library) {
		return library.required > 0;
	}
	const defaultMember = resolved ? hostDefaultMember(resolved, memberCtx) : undefined;
	if (!defaultMember) {
		return false;
	}
	if (defaultMember.kind === 'method') {
		return /\((?!\s*\[)[^)]/.test(defaultMember.signature ?? '');
	}
	return true;
}

/** A local whose value the procedure's text fixes: its default, or one literal. */
export interface KnownLocalValue {
	/** 'empty' is a Variant nothing ever assigns, whose value 0 is Empty's as a number. */
	kind: 'number' | 'string' | 'empty';
	value: number | string;
	/** 'default' when nothing ever assigns it, 'literal' when every assignment is the same literal. */
	origin: 'default' | 'literal';
	/**
	 * A `Mid(x, ...) = ` statement rewrites characters of the value without
	 * changing its length, so the length is still known and the characters are
	 * not.
	 */
	contentMutated?: boolean;
}

const MODULE_MEMBER_NAMES = new WeakMap<ReturnType<typeof buildModuleSymbols>, ReadonlySet<string>>();

/** The lowercased names the module declares at its top level, which shadow the VBA library's. */
function moduleMemberNames(symbols: ReturnType<typeof buildModuleSymbols>): ReadonlySet<string> {
	let names = MODULE_MEMBER_NAMES.get(symbols);
	if (!names) {
		names = new Set((symbols.root.children ?? []).map((child) => child.name.toLowerCase()));
		MODULE_MEMBER_NAMES.set(symbols, names);
	}
	return names;
}

/**
 * The locals of a procedure whose value is plain from the text (issues #118
 * and #119): a variable nothing ever assigns holds its default - 0 for a
 * number, "" for a String - and one whose every assignment is the same
 * literal holds that literal. Anything that could change it another way -
 * passing it whole to a call (ByRef), a For counter, `Input #`/`Get #`/
 * `Line Input #`, `Mid(x, ...) =`, ReDim, `Set` - drops it from the map, as
 * does any assignment whose value is not a plain literal. Variant and object
 * locals are left out: Empty and Nothing are not the values these rules ask
 * about.
 */
export function knownLocalLiteralValues(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
): Map<string, KnownLocalValue> {
	// A Variant (or untyped) local has no kind until a literal gives it one;
	// literals of two kinds, or none, leave it unknown (issue #121: `v = 5`
	// then `v.Foo`).
	const candidates = new Map<string, { kind: 'number' | 'string' | undefined; literals: Set<string>; mutated: boolean; contentMutated: boolean }>();
	for (const [lower, kind] of literalValueLocals(proc, symbols)) {
		candidates.set(lower, { kind, literals: new Set(), mutated: false, contentMutated: false });
	}
	if (candidates.size === 0) {
		return moduleVariableDefaults(source, proc, symbols);
	}
	const bytes = byteVariables(proc, symbols);
	const wholes = wholeNumberVariables(proc, symbols);
	const mutate = (lower: string | undefined): void => {
		const entry = lower ? candidates.get(lower) : undefined;
		if (entry) {
			entry.mutated = true;
		}
	};
	const visit = (body: readonly BodyNode[]): void => {
		for (const node of body) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (node.kind === 'ForBlock') {
				mutate(node.controlVariable?.toLowerCase());
			}
			if ('body' in node && Array.isArray(node.body)) {
				// A header passes a name ByRef as a value does:
				// `If Take(1, d) = 0 Then`, `Loop While Take(d)`. An ElseIf
				// line is a statement of the If's flat body.
				for (const span of [blockHeaderLineSpan(source, node.span), blockFooterLineSpan(source, node.span)]) {
					mutateWholeArguments(statementTokens(source, span), 0, true);
				}
				visit(node.body as BodyNode[]);
				continue;
			}
			if (!isLeafStatement(node)) {
				continue;
			}
			for (const span of statementAndBranchSpans(node)) {
				const toks = statementTokens(source, span);
				const first = firstExecutableTokenIndex(toks);
				const head = tokenText(toks[first]);
				const bare = bareAssignmentTarget(source, span);
				// `d = d + 0` leaves d as it was (issue #350).
				if (bare && identityAssignment(bare.name, bare.valueTokens)) {
					continue;
				}
				if (bare) {
					const entry = candidates.get(bare.name.toLowerCase());
					if (entry) {
						const value = toks.slice(first + 2).filter((tok) => tok.kind !== 'comment');
						const kind = entry.kind ?? (unwrapOuterParens(value)[0]?.kind === 'stringLiteral' ? 'string' : 'number');
						const literal = plainLiteralText(value, kind, entry.kind !== undefined, bytes.has(bare.name.toLowerCase()), wholes.has(bare.name.toLowerCase()));
						if (literal === undefined || (entry.kind !== undefined && entry.kind !== kind)) {
							entry.mutated = true;
						} else {
							entry.kind = kind;
							entry.literals.add(literal);
						}
					}
					// The value may pass a name ByRef: `w = Take(d)` (issue #238).
					mutateWholeArguments(toks, first + 2, true);
					continue;
				}
				// A ReDim writes its arrays alone: `ReDim a(n)` reads n (issue #350).
				if (head === 'redim') {
					const start = tokenText(toks[first + 1]) === 'preserve' ? first + 2 : first + 1;
					for (const group of splitTopLevelTokenGroups(toks, start, ',', toks.length)) {
						mutate(tokenName(group.find((tok) => tok.kind !== 'comment'))?.toLowerCase());
					}
					continue;
				}
				if (head === 'set' || head === 'input' || head === 'get' || head === 'line' || head === 'erase') {
					for (const tok of toks) {
						mutate(tokenName(tok)?.toLowerCase());
					}
					continue;
				}
				// `Mid$` lexes as Mid and a `$` of its own (issue #327).
				const midOpen = toks[first + 1]?.rawText === '$' ? first + 2 : first + 1;
				if ((head === 'mid' || head === 'mid$') && toks[midOpen]?.rawText === '(') {
					// `Mid(x, start, len) = value` rewrites characters of x and
					// keeps its length; anything else named in it is read.
					const target = candidates.get(tokenName(toks[midOpen + 1])?.toLowerCase() ?? '');
					if (target) {
						target.contentMutated = true;
					}
					continue;
				}
				if (head === 'lset' || head === 'rset') {
					for (const tok of toks) {
						mutate(tokenName(tok)?.toLowerCase());
					}
					continue;
				}
				// A Case line reads values: `Case w` passes w to nothing, and
				// only a call in it can (issue #268).
				mutateWholeArguments(toks, head === 'case' ? first + 1 : 0, head === 'case');
			}
		}
	};
	// A whole name passed to any call may be ByRef: `Take d`, `Take(d)`,
	// `Call Take(d)`, `x = Take(d)`. Only a name standing alone in an
	// argument slot counts; `Take(d + 1)` copies.
	const mutateWholeArguments = (toks: readonly VbaToken[], from: number, inValue: boolean): void => {
		// In a value a call takes parentheses: `x = 10 Mod d` passes nothing.
		for (let i = from; i < toks.length; i++) {
			const name = tokenName(toks[i])?.toLowerCase();
			if (!name || !candidates.has(name)) {
				continue;
			}
			const prev = toks[i - 1];
			const next = toks[i + 1];
			// `SetN n:=n` passes n too, by name (issue #449).
			const opensSlot = prev?.rawText === '(' || prev?.rawText === ',' || prev?.rawText === ':=' || (!inValue && (prev === undefined || prev.kind === 'identifier' || prev.kind === 'keyword'));
			const closesSlot = next === undefined || next.rawText === ')' || next.rawText === ',' || next.rawText === ':' || next.kind === 'comment';
			if (opensSlot && closesSlot && !(prev?.kind === 'operator' && prev.rawText !== ':=') && !(next?.kind === 'operator') && !(inValue && libraryFunctionArgument(toks, i))) {
				mutate(name);
			}
		}
	};
	// A VBA library function assigns none of its arguments: `n = CLng(s)`.
	const libraryFunctionArgument = (toks: readonly VbaToken[], at: number): boolean => {
		let depth = 0;
		for (let j = at - 1; j > 0; j--) {
			if (toks[j].rawText === ')') {
				depth++;
			} else if (toks[j].rawText === '(' && depth-- === 0) {
				// `Left$(` lexes as Left and a `$` (issue #334).
				const at = toks[j - 1]?.rawText === '$' ? j - 2 : j - 1;
				const callee = tokenName(toks[at])?.toLowerCase();
				const qualified = toks[at - 1]?.rawText === '.';
				if (!callee || (qualified && tokenText(toks[at - 2]) !== 'vba') || (!qualified && moduleMemberNames(symbols).has(callee))) {
					return false;
				}
				// The library knows String only as String$.
				return (resolveRuntimeFunction(callee) ?? resolveRuntimeFunction(`${callee}$`))?.kind === 'function';
			}
		}
		return false;
	};
	visit(proc.body);
	const out = new Map<string, KnownLocalValue>();
	for (const [lower, entry] of candidates) {
		if (entry.mutated) {
			continue;
		}
		if (entry.kind === undefined) {
			// A Variant nothing assigns is Empty, which divides as 0: `5 / v`
			// raises 11 and `v / v` 6 (issue #219, measured in Excel 16.0).
			if (!entry.contentMutated) {
				out.set(lower, { kind: 'empty', value: 0, origin: 'default' });
			}
			continue;
		}
		const contentMutated = entry.contentMutated ? { contentMutated: true } : {};
		if (entry.literals.size === 0) {
			out.set(lower, { kind: entry.kind, value: entry.kind === 'number' ? 0 : '', origin: 'default', ...contentMutated });
		} else if (entry.literals.size === 1) {
			const [text] = entry.literals;
			out.set(lower, {
				kind: entry.kind,
				value: entry.kind === 'number' ? Number(text) : text,
				origin: 'literal',
				...contentMutated,
			});
		}
	}
	for (const [lower, value] of moduleVariableDefaults(source, proc, symbols)) {
		if (!out.has(lower)) {
			out.set(lower, value);
		}
	}
	return out;
}

/**
 * The initial value of each module variable nothing writes, as a procedure
 * sees it (issue #241): 0 for a number, "" for a String, Empty for a
 * Variant. A local or parameter of the same name hides it.
 */
function moduleVariableDefaults(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, KnownLocalValue> {
	const out = new Map<string, KnownLocalValue>();
	for (const [lower, variable] of untouchedModuleVariablesIn(source, symbols, proc)) {
		if (variable.isArray) {
			continue;
		}
		const type = normalizeType(variable.asType);
		if (type === undefined || type === 'variant') {
			out.set(lower, { kind: 'empty', value: 0, origin: 'default' });
		} else if (isNumericType(type)) {
			out.set(lower, { kind: 'number', value: 0, origin: 'default' });
		} else if (type === 'string') {
			out.set(lower, { kind: 'string', value: '', origin: 'default' });
		}
	}
	return out;
}

/**
 * The locals whose literal value the rules follow, with their kind: 'number'
 * or 'string' from the declared type, undefined for a Variant, which takes
 * its kind from the literal.
 */
function literalValueLocals(
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, 'number' | 'string' | undefined> {
	const out = new Map<string, 'number' | 'string' | undefined>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		if (child.kind !== 'localVariable' || child.isArray || child.visibility === 'Static') {
			continue;
		}
		const type = normalizeType(child.asType);
		const kind = type === undefined || type === 'variant' ? undefined : isNumericType(type) || type === 'boolean' || type === 'date' ? 'number' : type === 'string' ? 'string' : 'other';
		if (kind === 'other' || child.fixedLength !== undefined) {
			continue;
		}
		out.set(child.name.toLowerCase(), kind);
	}
	return out;
}

/** The Static locals {@link literalValueLocals} leaves out, with the same kinds. */
function staticValueLocals(proc: ProcedureNode, symbols: ReturnType<typeof buildModuleSymbols>): Map<string, 'number' | 'string' | undefined> {
	const out = new Map<string, 'number' | 'string' | undefined>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		if (child.kind !== 'localVariable' || child.isArray || child.visibility !== 'Static' || child.fixedLength !== undefined) {
			continue;
		}
		const type = normalizeType(child.asType);
		const kind = type === undefined || type === 'variant' ? undefined : isNumericType(type) || type === 'boolean' || type === 'date' ? 'number' : type === 'string' ? 'string' : 'other';
		if (kind !== 'other') {
			out.set(child.name.toLowerCase(), kind);
		}
	}
	return out;
}

/**
 * The values one statement sees: the procedure-wide values, with each local
 * a straight-line assignment reaches read from that assignment instead. A
 * name is worked out when it is first asked for, and the whole map only when
 * something walks it: every statement holds every local's reaching value,
 * so building each map in full cost statements times locals (issue #322).
 */
class StatementValues implements ReadonlyMap<string, KnownLocalValue> {
	private readonly known = new Map<string, KnownLocalValue | undefined>();
	private full: Map<string, KnownLocalValue> | undefined;

	constructor(
		private readonly whole: ReadonlyMap<string, KnownLocalValue>,
		private readonly assignments: ReachingAssignments,
		/** The value an assignment gives a name, undefined for none, 'whole' for the procedure-wide one. */
		private readonly derive: (lower: string, value: readonly VbaToken[]) => KnownLocalValue | undefined | 'whole',
		/** Whether a name no assignment reaches keeps its procedure-wide value. */
		private readonly keepsWhole: (lower: string) => boolean = () => true,
	) {}

	get(lower: string): KnownLocalValue | undefined {
		if (this.full) {
			return this.full.get(lower);
		}
		if (this.known.has(lower)) {
			return this.known.get(lower);
		}
		// A Variant's starting Empty is for the guards (issue #691): these
		// values keep reading what the procedure as a whole says of it.
		const raw = this.assignments.get(lower);
		const value = raw === VARIANT_EMPTY ? undefined : raw;
		const derived = value ? this.derive(lower, value) : this.keepsWhole(lower) ? 'whole' : undefined;
		const out = derived === 'whole' ? this.whole.get(lower) : derived;
		this.known.set(lower, out);
		return out;
	}

	has(lower: string): boolean {
		return this.get(lower) !== undefined;
	}

	get size(): number {
		return this.all().size;
	}

	forEach(visit: (value: KnownLocalValue, key: string, map: ReadonlyMap<string, KnownLocalValue>) => void, thisArg?: unknown): void {
		this.all().forEach((value, key) => visit.call(thisArg, value, key, this));
	}

	entries(): MapIterator<[string, KnownLocalValue]> {
		return this.all().entries();
	}

	keys(): MapIterator<string> {
		return this.all().keys();
	}

	values(): MapIterator<KnownLocalValue> {
		return this.all().values();
	}

	[Symbol.iterator](): MapIterator<[string, KnownLocalValue]> {
		return this.all()[Symbol.iterator]();
	}

	private all(): Map<string, KnownLocalValue> {
		if (!this.full) {
			const full = new Map(this.whole);
			for (const lower of this.whole.keys()) {
				if ((!this.assignments.has(lower) || this.assignments.get(lower) === VARIANT_EMPTY) && !this.keepsWhole(lower)) {
					full.delete(lower);
				}
			}
			for (const [lower, value] of this.assignments) {
				if (value === VARIANT_EMPTY) {
					continue;
				}
				const derived = this.derive(lower, value);
				if (derived === 'whole') {
					continue;
				}
				if (derived === undefined) {
					full.delete(lower);
				} else {
					full.set(lower, derived);
				}
			}
			this.full = full;
		}
		return this.full;
	}
}

/**
 * A view of a statement's values through `pick`, worked out per name on
 * first use: the rules that keep only a statement's strings or literals
 * look up the names the statement uses, and walking every local at every
 * statement cost statements times locals (issue #322).
 */
export function pickedValues<T>(values: ReadonlyMap<string, KnownLocalValue>, pick: (lower: string, value: KnownLocalValue) => T | undefined): ReadonlyMap<string, T> {
	return new PickedValues(values, pick);
}

class PickedValues<T> implements ReadonlyMap<string, T> {
	private readonly known = new Map<string, T | undefined>();
	private full: Map<string, T> | undefined;

	constructor(private readonly base: ReadonlyMap<string, KnownLocalValue>, private readonly pick: (lower: string, value: KnownLocalValue) => T | undefined) {}

	get(lower: string): T | undefined {
		if (this.known.has(lower)) {
			return this.known.get(lower);
		}
		const value = this.base.get(lower);
		const out = value === undefined ? undefined : this.pick(lower, value);
		this.known.set(lower, out);
		return out;
	}

	has(lower: string): boolean {
		return this.get(lower) !== undefined;
	}

	get size(): number {
		return this.all().size;
	}

	forEach(visit: (value: T, key: string, map: ReadonlyMap<string, T>) => void, thisArg?: unknown): void {
		this.all().forEach((value, key) => visit.call(thisArg, value, key, this));
	}

	entries(): MapIterator<[string, T]> {
		return this.all().entries();
	}

	keys(): MapIterator<string> {
		return this.all().keys();
	}

	values(): MapIterator<T> {
		return this.all().values();
	}

	[Symbol.iterator](): MapIterator<[string, T]> {
		return this.all()[Symbol.iterator]();
	}

	private all(): Map<string, T> {
		if (!this.full) {
			const full = new Map<string, T>();
			for (const [lower, value] of this.base) {
				const picked = this.pick(lower, value);
				if (picked !== undefined) {
					full.set(lower, picked);
				}
			}
			this.full = full;
		}
		return this.full;
	}
}

type LocalValuesAt = (stmt: BodyNode | undefined) => ReadonlyMap<string, KnownLocalValue>;

// Rules share one immutable view of a procedure's values. Symbols own the
// cache so a changed module/project context cannot reuse the old analysis.
const LOCAL_VALUES_AT = new WeakMap<ReturnType<typeof buildModuleSymbols>, WeakMap<ProcedureNode, {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	valuesAt: LocalValuesAt;
}>>();

/**
 * {@link knownLocalLiteralValues} at each statement (issue #180). Where the
 * last assignment to reach a statement in a straight line is a literal, the
 * statement sees that literal, though other assignments in the procedure
 * disagree with it: `d = 0: x = 10 / d: d = 2` divides by 0. Elsewhere it
 * sees the procedure-wide value.
 */
export function knownLocalLiteralValuesAt(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
): LocalValuesAt {
	const cache = perProcedureCache(LOCAL_VALUES_AT, symbols);
	const cached = cache.get(proc);
	if (cached && cached.source === source && cached.activity === activity) {
		cached.source = source;
		return cached.valuesAt;
	}
	const valuesAt = buildKnownLocalLiteralValuesAt(source, proc, symbols, activity);
	cache.set(proc, { source, activity, valuesAt });
	return valuesAt;
}

function buildKnownLocalLiteralValuesAt(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
): LocalValuesAt {
	const whole = knownLocalLiteralValues(source, proc, symbols, activity);
	const startLocals = literalValueLocals(proc, symbols);
	// A Static local holds what the last call left as the procedure starts,
	// but what the straight line assigns it after that (issue #685): it has
	// no start of its own and no procedure-wide value.
	const locals = new Map([...startLocals, ...staticValueLocals(proc, symbols)]);
	const moduleVariables = followedModuleVariables(proc, symbols);
	const bytes = byteVariables(proc, symbols);
	const wholes = wholeNumberVariables(proc, symbols);
	let writes: ReadonlyMap<string, readonly BodyNode[]> | undefined;
	// The same start as unreachableStatementsIn, so the two share one walk.
	const reaching = locals.size === 0 && moduleVariables.size === 0
		? new Map()
		: straightLineAssignments(source, proc.body, activity, walkStartWithEffects(source, symbols, proc, startLocals, activity));
	// Statements in a run share one reaching map, so they share one result.
	const results = new Map<ReachingAssignments, ReadonlyMap<string, KnownLocalValue>>();
	// A block's opening line, given as a statement of its own, sees what
	// reaches the block: `x = "abc"` then `While x` (issue #424).
	let blocksByStart: Map<number, BodyNode> | undefined;
	const blockAt = (stmt: BodyNode): BodyNode | undefined => {
		blocksByStart ??= new Map([...reaching.keys()].filter((node: BodyNode) => !isLeafStatement(node)).map((node: BodyNode) => [node.span.start, node]));
		return blocksByStart.get(stmt.span.start);
	};
	// One assignment reaches many statements, its tokens shared by all of
	// them: what it gives a name is worked out once (issue #322).
	const derived = new WeakMap<readonly VbaToken[], Map<string, KnownLocalValue | undefined | 'whole'>>();
	const derive = (lower: string, value: readonly VbaToken[]): KnownLocalValue | undefined | 'whole' => {
		let byName = derived.get(value);
		if (!byName) {
			byName = new Map();
			derived.set(value, byName);
		}
		if (byName.has(lower)) {
			return byName.get(lower);
		}
		let out: KnownLocalValue | undefined | 'whole';
		if (!locals.has(lower)) {
			out = 'whole';
		} else {
			const kind = locals.get(lower) ?? (unwrapOuterParens([...value])[0]?.kind === 'stringLiteral' ? 'string' : 'number');
			const literal = plainLiteralText([...value], kind, locals.get(lower) !== undefined, bytes.has(lower), wholes.has(lower));
			const origin = value === DEFAULT_NUMBER || value === DEFAULT_STRING ? 'default' : 'literal';
			out = literal === undefined ? undefined : { kind, value: kind === 'number' ? Number(literal) : literal, origin };
		}
		byName.set(lower, out);
		return out;
	};
	return (stmt) => {
		// No statement: a block header, which sees the procedure-wide values.
		const assignments = stmt ? reaching.get(stmt) ?? (stmt.kind === 'Statement' && isLeafStatement(stmt) ? reaching.get(blockAt(stmt)!) : undefined) : undefined;
		if (!assignments) {
			return whole;
		}
		let result = results.get(assignments);
		if (!result) {
			// A typed local starts the walk with its default, so one it no
			// longer holds may have kept that default on some path: with
			// `If b > 5000 Then a = 4` undecided, a is 0 or 4 after it, not
			// the 4 it holds wherever it is assigned (issue #565).
			result = new StatementValues(whole, assignments, derive, (lower) => locals.get(lower) === undefined || whole.get(lower)?.origin !== 'literal');
			results.set(assignments, result);
		}
		if (moduleVariables.size === 0) {
			return result;
		}
		// A module variable written in the straight line holds the value
		// while nothing between could run other code (issue #348).
		let withModule: Map<string, KnownLocalValue> | undefined;
		for (const [lower, value] of assignments) {
			if (!moduleVariables.has(lower) || locals.has(lower)) {
				continue;
			}
			const kind = moduleVariables.get(lower) ?? (unwrapOuterParens(value)[0]?.kind === 'stringLiteral' ? 'string' : 'number');
			const literal = plainLiteralText([...value], kind, moduleVariables.get(lower) !== undefined, bytes.has(lower), wholes.has(lower));
			// The reaching write is the last one that runs before the
			// statement: a later one in a block would have ended the value.
			const dead = unreachableStatementsIn(source, proc, symbols, activity);
			const write = [...(writes ??= moduleVariableWrites(source, proc, activity)).get(lower) ?? []]
				.reverse().find((node) => node.span.end <= stmt!.span.start && !dead.has(node));
			// A call to a procedure of this module that leaves the variable
			// alone, and runs no code but the module's own, keeps its value
			// (issue #618).
			const leavesAlone = (name: string): boolean => calleeLeavesAlone(source, symbols, name, lower);
			if (literal === undefined || !write
				|| codeMayRun(statementTokens(source, { start: write.span.end, end: stmt!.span.end }), proc, symbols, leavesAlone)) {
				continue;
			}
			withModule ??= new Map(result);
			withModule.set(lower, { kind, value: kind === 'number' ? Number(literal) : literal, origin: 'literal' });
		}
		return withModule ?? result;
	};
}

/**
 * The module variables a procedure may follow through its straight line,
 * with their kind as {@link literalValueLocals} gives a local's. A local or
 * parameter of the same name hides one.
 */
const MODULE_BYTE_VARIABLES = new WeakMap<object, ReadonlySet<string>>();

/** The locals and module variables a procedure sees that are declared As Byte, by lowercased name. */
function byteVariables(proc: ProcedureNode, symbols: ReturnType<typeof buildModuleSymbols>): { has(lower: string): boolean } {
	let module = MODULE_BYTE_VARIABLES.get(symbols);
	if (!module) {
		module = new Set((symbols.root.children ?? [])
			.filter((sym) => sym.kind === 'moduleVariable' && normalizeType(sym.asType) === 'byte')
			.map((sym) => sym.name.toLowerCase()));
		MODULE_BYTE_VARIABLES.set(symbols, module);
	}
	const locals = new Map((procedureSymbolFor(symbols, proc)?.children ?? []).map((child) => [child.name.toLowerCase(), child]));
	const moduleBytes = module;
	return {
		has: (lower) => {
			const local = locals.get(lower);
			return local ? local.kind === 'localVariable' && normalizeType(local.asType) === 'byte' : moduleBytes.has(lower);
		},
	};
}

const MODULE_WHOLE_VARIABLES = new WeakMap<object, ReadonlySet<string>>();

/** The whole-number types, which store a fraction rounded half to even: `a As Long = 4.4` holds 4. */
const WHOLE_NUMBER_TYPES: ReadonlySet<string> = new Set(['byte', 'integer', 'long', 'longlong', 'longptr']);

/** The locals and module variables a procedure sees that are declared a whole-number type (issue #685). */
function wholeNumberVariables(proc: ProcedureNode, symbols: ReturnType<typeof buildModuleSymbols>): { has(lower: string): boolean } {
	let module = MODULE_WHOLE_VARIABLES.get(symbols);
	if (!module) {
		module = new Set((symbols.root.children ?? [])
			.filter((sym) => sym.kind === 'moduleVariable' && WHOLE_NUMBER_TYPES.has(normalizeType(sym.asType) ?? ''))
			.map((sym) => sym.name.toLowerCase()));
		MODULE_WHOLE_VARIABLES.set(symbols, module);
	}
	const locals = new Map((procedureSymbolFor(symbols, proc)?.children ?? []).map((child) => [child.name.toLowerCase(), child]));
	const moduleWholes = module;
	return {
		has: (lower) => {
			const local = locals.get(lower);
			return local ? local.kind === 'localVariable' && WHOLE_NUMBER_TYPES.has(normalizeType(local.asType) ?? '') : moduleWholes.has(lower);
		},
	};
}

function followedModuleVariables(
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
): ReadonlyMap<string, 'number' | 'string' | undefined> {
	const hidden = new Set([proc.name, ...proc.params.map((param) => param.name), ...(procedureSymbolFor(symbols, proc)?.children ?? []).map((child) => child.name)]
		.map((name) => name.toLowerCase()));
	const out = new Map<string, 'number' | 'string' | undefined>();
	for (const child of symbols.root.children ?? []) {
		if (child.kind !== 'moduleVariable' || child.isArray || child.isAutoInstantiated || child.fixedLength !== undefined || hidden.has(child.name.toLowerCase())) {
			continue;
		}
		const type = normalizeType(child.asType);
		const kind = type === undefined || type === 'variant' ? undefined : isNumericType(type) ? 'number' : type === 'string' ? 'string' : 'other';
		if (kind !== 'other') {
			out.set(child.name.toLowerCase(), kind);
		}
	}
	return out;
}

/** The statements of a procedure that assign a name, `x = value`, in source order, by lowercased name. */
function moduleVariableWrites(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
): Map<string, BodyNode[]> {
	const out = new Map<string, BodyNode[]>();
	const visit = (body: readonly BodyNode[]): void => {
		for (const node of body) {
			if (activity?.isInactive(node.span)) {
				continue;
			}
			if (node.kind === 'IfBlock') {
				for (const branch of (node as IfBlockNode).branches) {
					visit(branch.body);
				}
				continue;
			}
			if ('body' in node && Array.isArray(node.body)) {
				visit(node.body as BodyNode[]);
				continue;
			}
			const bare = isLeafStatement(node) ? bareAssignmentTarget(source, node.span) : undefined;
			if (bare) {
				const lower = bare.name.toLowerCase();
				out.set(lower, [...(out.get(lower) ?? []), node]);
			}
		}
	};
	visit(proc.body);
	return out;
}

/** VBA functions that wait for the user or call by name, so other code may run meanwhile. */
const CODE_RUNNING_FUNCTIONS: ReadonlySet<string> = new Set(['callbyname', 'doevents', 'inputbox', 'msgbox']);

/**
 * Whether the tokens may run code other than the procedure's own, which
 * could write a module variable: a call to a procedure, a member of an
 * object (a class's property, or a host event), `New` of a class,
 * RaiseEvent, or a ByRef parameter, which may be the module variable itself.
 * The procedure's locals, its ByVal parameters, the module's variables and
 * Consts, and the VBA library's functions and constants run nothing.
 */
function codeMayRun(toks: readonly VbaToken[], proc: ProcedureNode, symbols: ReturnType<typeof buildModuleSymbols>, leavesAlone?: (name: string) => boolean): boolean {
	const safe = new Set<string>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		safe.add(child.name.toLowerCase());
	}
	for (const param of proc.params) {
		if (!param.byVal) {
			safe.delete(param.name.toLowerCase());
		}
	}
	return codeMayRunWith(toks, proc.name.toLowerCase(), safe, symbols, leavesAlone);
}

/**
 * Whether these tokens may run code that changes what the rules follow:
 * `safe` names are this procedure's own, and a call `leavesAlone` clears is
 * one the rules have read through.
 */
function codeMayRunWith(
	toks: readonly VbaToken[],
	ownName: string,
	ownSafe: ReadonlySet<string>,
	symbols: ReturnType<typeof buildModuleSymbols>,
	leavesAlone?: (name: string) => boolean,
): boolean {
	const safe = new Set(ownSafe);
	for (const child of symbols.root.children ?? []) {
		if (child.kind === 'moduleVariable' || child.kind === 'constant') {
			safe.add(child.name.toLowerCase());
		}
	}
	for (let i = 0; i < toks.length; i++) {
		const tok = toks[i];
		if (tok.kind === 'comment') {
			continue;
		}
		const prev = toks[i - 1];
		const name = tokenName(tok)?.toLowerCase();
		if (prev?.rawText === '.' || prev?.rawText === '!') {
			// `Debug.Print` and `Err.Number` run nothing of the project's.
			const object = tokenText(toks[i - 2]);
			if ((object === 'debug' || object === 'err') && name !== 'raise') {
				continue;
			}
			return true;
		}
		if (!name) {
			continue;
		}
		if (name === 'raiseevent') {
			return true;
		}
		if (tokenText(prev) === 'new') {
			// `Dim c As New Collection` makes nothing until c is used.
			if (name !== 'collection' && tokenText(toks[i - 2]) !== 'as') {
				return true;
			}
			continue;
		}
		if (tok.kind === 'keyword' && !resolveRuntimeFunction(name)) {
			continue;
		}
		if (tokenText(prev) === 'as' || safe.has(name) || resolveRuntimeConstant(name)) {
			continue;
		}
		// The procedure's own name is its result; with an argument list it is a call.
		if (name === ownName && toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const fn = resolveRuntimeFunction(name) ?? resolveRuntimeFunction(`${name}$`);
		if (fn && !CODE_RUNNING_FUNCTIONS.has(name) && !moduleMemberNames(symbols).has(name)) {
			continue;
		}
		if (leavesAlone?.(name)) {
			continue;
		}
		return true;
	}
	return false;
}

/**
 * Whether a statement may run code that changes the module variable: a call
 * to a procedure of the module that writes it, or any code outside the
 * module's own (issue #618).
 */
export function statementMayChangeModuleVariable(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	span: Span,
	variable: string,
): boolean {
	return codeMayRun(statementTokens(source, span), proc, symbols, (name) => calleeLeavesAlone(source, symbols, name, variable));
}

const LEAVES_ALONE = new WeakMap<object, Map<string, boolean>>();

/**
 * Whether a call to the module's procedure `name` leaves the module variable
 * alone: neither it nor any procedure of the module it calls writes the
 * variable, and none runs code outside the module's own (issue #618,
 * measured in Excel 16.0). Kept per module and asked by name.
 */
function calleeLeavesAlone(source: string, symbols: ReturnType<typeof buildModuleSymbols>, name: string, variable: string): boolean {
	let cache = LEAVES_ALONE.get(symbols);
	if (!cache) {
		cache = new Map();
		LEAVES_ALONE.set(symbols, cache);
	}
	const visiting = new Set<string>();
	const check = (callee: string): boolean => {
		const key = `${callee}|${variable}`;
		const known = cache!.get(key);
		if (known !== undefined) {
			return known;
		}
		if (visiting.has(callee)) {
			return true; // a cycle adds nothing the other procedures do not
		}
		const procedure = (symbols.root.children ?? []).filter((child) => isProcedureKind(child.kind) && child.name.toLowerCase() === callee);
		if (procedure.length === 0) {
			return false;
		}
		visiting.add(callee);
		let alone = true;
		for (const symbol of procedure) {
			const text = source.slice(symbol.fullSpan.start, symbol.fullSpan.end);
			const own = new Set((symbol.children ?? []).map((child) => child.name.toLowerCase()));
			// The callee's own lines, its header left out.
			const body = statementTokens(source, symbol.fullSpan).filter((tok) => tok.start >= symbol.nameSpan.end - symbol.fullSpan.start);
			if (writtenNamesIn(text).has(variable) || codeMayRunWith(body, callee, own, symbols, check)) {
				alone = false;
				break;
			}
		}
		visiting.delete(callee);
		cache!.set(key, alone);
		return alone;
	};
	return check(name);
}

/**
 * A typed local holds its declared default until something assigns it, the
 * statement that does included: `x = 1 / x` divides by 0 (issue #259).
 */
function declaredDefaults(locals: ReadonlyMap<string, 'number' | 'string' | undefined>): Map<string, readonly VbaToken[]> {
	const defaults = new Map<string, readonly VbaToken[]>();
	for (const [lower, kind] of locals) {
		if (kind !== undefined) {
			defaults.set(lower, kind === 'number' ? DEFAULT_NUMBER : DEFAULT_STRING);
		}
	}
	return defaults;
}


/**
 * Each Variant local, `As Variant` or with no type and no DefType for its
 * letter, holds Empty until assigned: `If Not IsEmpty(v) Then` never runs
 * its arm (issue #691).
 */
function variantStarts(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode): Map<string, readonly VbaToken[]> {
	const out = new Map<string, readonly VbaToken[]>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		if (child.kind === 'localVariable' && !child.isArray && child.visibility !== 'Static'
			&& (normalizeType(child.asType) ?? defTypeOf(symbols, child.name)?.toLowerCase() ?? 'variant') === 'variant') {
			out.set(child.name.toLowerCase(), VARIANT_EMPTY);
		}
	}
	return out;
}

/**
 * The statements of a procedure that never run, because a guard whose
 * value the straight-line walk knows decides against them (issue #273).
 */
export function unreachableStatementsIn(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
): ReadonlySet<BodyNode> {
	// Seven rules share these facts within one symbol/analysis context.
	const cache = perProcedureCache(UNREACHABLE, symbols);
	const kept = cache.get(proc);
	if (kept && kept.source === source && kept.activity === activity) {
		kept.source = source;
		return kept.dead;
	}
	// Walked even with no value known: `GoTo Done` leaves whatever the locals hold.
	const dead = straightLineUnreachable(source, proc.body, activity, walkStartWithEffects(source, symbols, proc, literalValueLocals(proc, symbols), activity));
	cache.set(proc, { source, activity, dead });
	return dead;
}

/** The one-line If branches of a procedure that never run (issue #430). */
export function deadBranchSpansIn(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
): readonly Span[] {
	return straightLineDeadBranches(source, proc.body, activity, walkStartWithEffects(source, symbols, proc, literalValueLocals(proc, symbols), activity));
}

/**
 * The straight-line walk from what holds as the procedure starts: each
 * local's default, and `Dim c As New Collection` empty (issue #614). The
 * same walk the value rules share, so it costs nothing more.
 */
export function defaultedStraightLine(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
): ReadonlyMap<BodyNode, ReachingAssignments> {
	return straightLineAssignments(source, proc.body, activity, walkStart(symbols, proc, literalValueLocals(proc, symbols)));
}

const UNREACHABLE = new WeakMap<ReturnType<typeof buildModuleSymbols>, WeakMap<ProcedureNode, {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	dead: ReadonlySet<BodyNode>;
}>>();

/**
 * What holds as a procedure starts: its Consts' values, and each local's
 * declared default. Several rules ask for each procedure, so it is kept, and
 * the walk's cache finds it by identity.
 */
function walkStart(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	locals: ReadonlyMap<string, 'number' | 'string' | undefined>,
): ReachingAssignments {
	// The parsed procedure can be reused with a different active symbol set.
	const cache = perProcedureCache(WALK_STARTS, symbols);
	let start = cache.get(proc);
	if (!start) {
		start = new Map([...conditionConstants(symbols, proc), ...declaredDefaults(locals), ...variantStarts(symbols, proc), ...objectStarts(symbols, proc)]);
		cache.set(proc, start);
	}
	return start;
}

const WALK_STARTS = new WeakMap<ReturnType<typeof buildModuleSymbols>, WeakMap<ProcedureNode, ReachingAssignments>>();

/**
 * Whether a statement that may leave the procedure early still runs: a
 * path that leaves returns what it held then, which the end state does not
 * show. One the walk found never runs may stay.
 */
function mayLeaveEarly(
	source: string,
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	dead: ReadonlySet<BodyNode>,
	deadSpans: readonly Span[],
): boolean {
	const neverRuns = (span: Span): boolean => deadSpans.some((deadSpan) => span.start >= deadSpan.start && span.end <= deadSpan.end);
	return body.some((node) => {
		if (activity?.isInactive(node.span) || dead.has(node)) {
			return false;
		}
		if (isLeafStatement(node)) {
			return statementAndBranchSpans(node).some((span) => {
				const toks = statementTokens(source, span);
				const head = tokenText(toks[firstExecutableTokenIndex(toks)]);
				return !neverRuns(span) && (RESULT_LEAVING_HEADS.has(head) || (head === 'err' && tokenText(toks[firstExecutableTokenIndex(toks) + 2]) === 'raise'));
			});
		}
		return 'body' in node && Array.isArray(node.body) && mayLeaveEarly(source, node.body as BodyNode[], activity, dead, deadSpans);
	});
}

/**
 * What the module's calls leave in the names they pass ByRef (issue #449,
 * measured in Excel 16.0): `ZeroN n`, where ZeroN's body runs through to
 * `n = 0`, leaves the caller's n 0. A parameter ByVal, an array, one the
 * callee may leave early or assign on one path only, and an argument in
 * parentheses of its own, which passes a copy, leave nothing known.
 */
function callEffectsFor(source: string, symbols: ReturnType<typeof buildModuleSymbols>, activity: ConditionalActivityTracker | undefined): CallEffects {
	const kept = CALL_EFFECT_READERS.get(symbols);
	if (kept && kept.source === source && kept.activity === activity) {
		kept.source = source;
		return kept.effects;
	}
	let procedures: Map<string, ProcedureNode | null> | undefined;
	const procedureNamed = (lower: string): ProcedureNode | undefined => {
		if (!procedures) {
			procedures = new Map();
			for (const member of parseModule(source).members) {
				if (member.kind === 'Procedure' && !activity?.isInactive(member.span) && (member.procKind === 'Sub' || member.procKind === 'Function')) {
					const key = member.name.toLowerCase();
					procedures.set(key, procedures.has(key) ? null : member);
				}
			}
		}
		return procedures.get(lower) ?? undefined;
	};
	const left = new Map<string, readonly VbaToken[] | undefined>();
	const leftIn = (proc: ProcedureNode, index: number): readonly VbaToken[] | undefined => {
		const key = `${proc.name.toLowerCase()}|${index}`;
		if (!left.has(key)) {
			const param = proc.params[index];
			let value: readonly VbaToken[] | undefined;
			if (param && !param.byVal && !param.paramArray && !param.isArray && !proc.modifiers.some((word) => word.toLowerCase() === 'static')) {
				// From a start of its own, which no effects ride: a callee
				// that calls itself is not followed into.
				const initial = new Map(walkStart(symbols, proc, literalValueLocals(proc, symbols)));
				initial.delete(param.name.toLowerCase());
				const { exit, dead, deadSpans } = straightLineExit(source, proc.body, activity, initial);
				const held = exit && !mayLeaveEarly(source, proc.body, activity, dead, deadSpans) ? exit.get(param.name.toLowerCase()) : undefined;
				value = held && held !== OBJECT_NOTHING && held !== EMPTY_COLLECTION && held.every((tok) => tokenName(tok) === undefined || ['true', 'false'].includes(tokenText(tok))) ? held : undefined;
			}
			left.set(key, value);
		}
		return left.get(key);
	};
	const effects: CallEffects = (toks) => {
		const out = new Map<string, readonly VbaToken[]>();
		const apply = (proc: ProcedureNode | undefined, args: readonly (readonly VbaToken[])[]): void => {
			if (!proc) {
				return;
			}
			for (const [index, arg] of args.entries()) {
				const parts = arg.filter((tok) => tok.kind !== 'comment');
				if (parts.some((tok) => tok.rawText === ':=')) {
					return;
				}
				const name = parts.length === 1 ? tokenName(parts[0])?.toLowerCase() : undefined;
				const value = name ? leftIn(proc, index) : undefined;
				if (name && value) {
					out.set(name, value);
				}
			}
		};
		const first = firstExecutableTokenIndex(toks);
		const head = tokenText(toks[first]);
		if (head === 'call' && tokenName(toks[first + 1]) && toks[first + 2]?.rawText === '(') {
			const close = matchParenFrom(toks, first + 2);
			apply(procedureNamed(tokenName(toks[first + 1])!.toLowerCase()), close > first + 3 ? splitTopLevelTokenGroups(toks, first + 3, ',', close) : []);
		} else if (tokenName(toks[first]) && toks[first + 1]?.rawText !== '=' && toks[first + 1]?.rawText !== '.' && toks[first + 1]?.rawText !== '(') {
			apply(procedureNamed(tokenName(toks[first])!.toLowerCase()), toks.length > first + 1 ? splitTopLevelTokenGroups(toks, first + 1, ',', toks.length) : []);
		}
		for (let i = first + 1; i + 1 < toks.length; i++) {
			const name = tokenName(toks[i])?.toLowerCase();
			if (!name || toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.' || (head === 'call' && i === first + 1)) {
				continue;
			}
			const proc = procedureNamed(name);
			if (proc?.procKind === 'Function') {
				const close = matchParenFrom(toks, i + 1);
				apply(proc, close > i + 2 ? splitTopLevelTokenGroups(toks, i + 2, ',', close) : []);
			}
		}
		return out;
	};
	CALL_EFFECT_READERS.set(symbols, { source, activity, effects });
	return effects;
}

const CALL_EFFECT_READERS = new WeakMap<object, { source: string; activity: ConditionalActivityTracker | undefined; effects: CallEffects }>();

/** A procedure's start, with the effects of the module's calls riding its walks. */
function walkStartWithEffects(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
	locals: ReadonlyMap<string, 'number' | 'string' | undefined>,
	activity: ConditionalActivityTracker | undefined,
): ReachingAssignments {
	const start = walkStart(symbols, proc, locals);
	setCallEffects(start, callEffectsFor(source, symbols, activity));
	setDeclaredFacts(start, declaredFactsFor(source, symbols, proc));
	return start;
}

/** A bare upper bound's lower bound: 1 under Option Base 1, else 0. */
const OPTION_BASE_ONE = /^[ \t]*Option[ \t]+Base[ \t]+1\b/im;

/**
 * What the declarations say of each local, for the guards (issue #691):
 * its declared type ("long", "long()" for an array), and a fixed
 * one-dimension array's bounds.
 */
const DECLARED_FACT_READERS = new WeakMap<ReturnType<typeof buildModuleSymbols>, WeakMap<ProcedureNode, { source: string; facts: DeclaredFacts }>>();

function declaredFactsFor(source: string, symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode): DeclaredFacts {
	const cache = perProcedureCache(DECLARED_FACT_READERS, symbols);
	const kept = cache.get(proc);
	if (kept?.source === source) {
		kept.source = source;
		return kept.facts;
	}
	const types = new Map<string, string>();
	const bounds = new Map<string, readonly [number, number]>();
	let base: number | undefined;
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		if (child.kind !== 'localVariable' || child.visibility === 'Static' || child.fixedLength !== undefined) {
			continue;
		}
		const lower = child.name.toLowerCase();
		const type = normalizeType(child.asType) ?? defTypeOf(symbols, child.name)?.toLowerCase() ?? 'variant';
		types.set(lower, child.isArray ? `${type}()` : type);
		const fixed = child.isArray ? /^\s*(?:(-?\d+)\s+To\s+)?(-?\d+)\s*$/i.exec(child.arrayBounds ?? '') : null;
		if (fixed) {
			const lowerBound = fixed[1] !== undefined ? Number(fixed[1]) : (base ??= OPTION_BASE_ONE.test(source) ? 1 : 0);
			bounds.set(lower, [lowerBound, Number(fixed[2])]);
		}
	}
	// The module's Consts and Enum members, which a local or parameter of the same name hides.
	let constants: ReadonlyMap<string, number | undefined> | undefined;
	const params = new Set(proc.params.map((param) => param.name.toLowerCase()));
	const facts: DeclaredFacts = {
		type: (lower) => types.get(lower),
		bounds: (lower) => bounds.get(lower),
		constant: (lower) => (types.has(lower) || params.has(lower) ? undefined
			: (constants ??= collectModuleLiteralIntegerConstants(parseModule(source), undefined)).get(lower)),
	};
	cache.set(proc, { source, facts });
	return facts;
}

/** Statement heads after which a Function may end before its last line. */
const RESULT_LEAVING_HEADS: ReadonlySet<string> = new Set(['exit', 'goto', 'gosub', 'return', 'resume', 'on', 'stop', 'error', 'end']);

/**
 * What a Function of the module returns for one call's arguments, as the
 * tokens of the value it last assigns its name (issue #562): `Sign1(-1)` runs
 * `If n > 0 Then Sign1 = 1 Else Sign1 = 0` with n = -1, so 0. An omitted
 * Optional takes its default. Undefined where the result depends on more
 * than the walk follows: a ParamArray, a Static Function, a statement that
 * may leave early and still runs, or a value the walk does not know.
 */
export function functionResultFor(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	args: ReadonlyArray<readonly VbaToken[] | undefined>,
	objectResult = false,
): readonly VbaToken[] | undefined {
	if (proc.procKind !== 'Function' || proc.modifiers.some((word) => word.toLowerCase() === 'static') || args.length > proc.params.length) {
		return undefined;
	}
	// Calls share results only within the same source, symbols and active branch.
	const key = `${objectResult}|${args.map((arg) => arg?.map((tok) => tok.rawText).join(' ') ?? '').join(',')}`;
	const cache = perProcedureCache(CALL_RESULTS, symbols);
	let kept = cache.get(proc);
	if (!kept || kept.source !== source || kept.activity !== activity) {
		kept = { source, activity, calls: new Map() };
		cache.set(proc, kept);
	}
	if (kept.calls.has(key)) {
		kept.source = source;
		return kept.calls.get(key);
	}
	const result = runFunctionFor(source, proc, symbols, activity, args, objectResult);
	kept.calls.set(key, result);
	return result;
}

const CALL_RESULTS = new WeakMap<ReturnType<typeof buildModuleSymbols>, WeakMap<ProcedureNode, {
	source: string;
	activity: ConditionalActivityTracker | undefined;
	calls: Map<string, readonly VbaToken[] | undefined>;
}>>();

function runFunctionFor(
	source: string,
	proc: ProcedureNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	args: ReadonlyArray<readonly VbaToken[] | undefined>,
	objectResult: boolean,
): readonly VbaToken[] | undefined {
	// An object Function's result starts Nothing; the caller knows its type
	// is an object's.
	const type = normalizeType(proc.returnType ?? 'Variant');
	const resultDefault = objectResult ? OBJECT_NOTHING
		: type === 'string' ? DEFAULT_STRING : type !== undefined && (isNumericType(type) || type === 'boolean') ? DEFAULT_NUMBER : undefined;
	if (!resultDefault) {
		return undefined;
	}
	const initial = new Map(walkStart(symbols, proc, literalValueLocals(proc, symbols)));
	// A module variable nothing writes holds its default: `GetMod = m`.
	for (const [name, held] of moduleVariableDefaults(source, proc, symbols)) {
		if (!initial.has(name) && (held.kind === 'number' || held.kind === 'string')) {
			initial.set(name, held.kind === 'number' ? DEFAULT_NUMBER : DEFAULT_STRING);
		}
	}
	for (const [index, param] of proc.params.entries()) {
		const value = args[index] ?? (param.optional && param.defaultRaw !== undefined ? rawExpressionTokens(param.defaultRaw) : undefined);
		if (param.paramArray || !value) {
			return undefined;
		}
		initial.set(param.name.toLowerCase(), value);
	}
	const lower = proc.name.toLowerCase();
	initial.set(lower, resultDefault);
	const { exit, dead, deadSpans } = straightLineExit(source, proc.body, activity, initial);
	if (!exit) {
		return undefined;
	}
	const value = mayLeaveEarly(source, proc.body, activity, dead, deadSpans) ? undefined : exit.get(lower);
	if (!value || value === OBJECT_NOTHING || value === EMPTY_COLLECTION || !value.some((tok) => tokenName(tok) !== undefined)) {
		return value;
	}
	// `Twice = n * 2`: a name nothing reassigns holds what the call gave it,
	// so the value folds with it.
	const folded = evaluateIntegerConstantExpression(value.map((tok) => tok.rawText).join(' '), {
		get: (name) => {
			const held = exit.get(name.toLowerCase());
			return held && held === initial.get(name.toLowerCase()) && held.every((tok) => tokenName(tok) === undefined)
				? evaluateIntegerConstantExpression(held.map((tok) => tok.rawText).join(' '), { get: () => undefined })
				: undefined;
		},
	});
	return folded === undefined ? undefined : rawExpressionTokens(String(folded));
}

/**
 * What an object local holds as the procedure starts (issue #483): Nothing
 * for one never set, and an empty Collection for `Dim c As New Collection`.
 */
function objectStarts(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode): Map<string, readonly VbaToken[]> {
	const out = new Map<string, readonly VbaToken[]>();
	for (const child of procedureSymbolFor(symbols, proc)?.children ?? []) {
		const type = normalizeType(child.asType);
		if (child.kind !== 'localVariable' || child.visibility === 'Static' || child.isArray || type === undefined || type === 'variant' || type === 'string' || isKnownScalarType(type)) {
			continue;
		}
		if (!child.isAutoInstantiated) {
			out.set(child.name.toLowerCase(), OBJECT_NOTHING);
		} else if (type === 'collection' || type === 'vba.collection') {
			out.set(child.name.toLowerCase(), EMPTY_COLLECTION);
		}
	}
	return out;
}

/**
 * The Consts a procedure sees whose value is one literal, as the walk reads
 * a value: `Const DEBUGGING = False` holds 0, so `If DEBUGGING Then` never
 * runs its arm (issue #406). A local or parameter of the same name hides a
 * module's.
 */
function conditionConstants(symbols: ReturnType<typeof buildModuleSymbols>, proc: ProcedureNode): Map<string, readonly VbaToken[]> {
	const children = procedureSymbolFor(symbols, proc)?.children ?? [];
	const out = new Map<string, readonly VbaToken[]>();
	for (const symbol of [...(symbols.root.children ?? []), ...children]) {
		const toks = symbol.kind === 'constant' && symbol.defaultRaw !== undefined
			? rawExpressionTokens(symbol.defaultRaw)
			: [];
		const word = toks.length === 1 ? tokenText(toks[0]) : '';
		const value = word === 'true' ? CONSTANT_TRUE
			: word === 'false' ? DEFAULT_NUMBER
			: toks.length === 1 && (toks[0].kind === 'stringLiteral' || /^\d+$/.test(toks[0].rawText)) ? toks
			: undefined;
		if (value) {
			out.set(symbol.name.toLowerCase(), value);
		} else if (children.includes(symbol)) {
			out.delete(symbol.name.toLowerCase());
		}
	}
	return out;
}

const CONSTANT_TRUE: readonly VbaToken[] = rawExpressionTokens('-1');

/** What a local of a number type and a String hold before anything assigns them. */
const DEFAULT_NUMBER: readonly VbaToken[] = rawExpressionTokens('0');
const DEFAULT_STRING: readonly VbaToken[] = rawExpressionTokens('""');

/** The literal a plain `x = literal` assigns, as text, or undefined for any other value. */
function plainLiteralText(value: VbaToken[], kind: 'number' | 'string', typed = false, byte = false, whole = false): string | undefined {
	const text = plainLiteralTextAsWritten(value, kind, typed, byte);
	// A whole-number type keeps a fraction rounded half to even: `a As Long =
	// 4.4` holds 4, so `r(a)` on `Dim r(3)` raises 9 (issue #685).
	return whole && text !== undefined && kind === 'number' && !Number.isInteger(Number(text)) ? String(bankersRound(Number(text)) + 0) : text;
}

function plainLiteralTextAsWritten(value: VbaToken[], kind: 'number' | 'string', typed: boolean, byte: boolean): string | undefined {
	const toks = unwrapOuterParens(value);
	if (kind === 'string') {
		return toks.length === 1 && toks[0].kind === 'stringLiteral' ? stringLiteralValue(toks[0].rawText) : undefined;
	}
	// A typed number or Boolean stores True as -1 and False as 0 (issue
	// #491), and a Byte stores True as 255 (issue #556); a Variant keeps a
	// Boolean, which is no number here.
	const word = toks.length === 1 && toks[0].kind === 'keyword' ? tokenText(toks[0]) : '';
	if (typed && (word === 'true' || word === 'false')) {
		return word === 'false' ? '0' : byte ? '255' : '-1';
	}
	let sign = 1;
	let rest = toks;
	if (rest[0]?.rawText === '-' || rest[0]?.rawText === '+') {
		sign = rest[0].rawText === '-' ? -1 : 1;
		rest = rest.slice(1);
	}
	if (rest.length !== 1) {
		// Whole numbers and operators only: `x = 1 \ 3` stores 0 (issue #286).
		const folded = toks.length > 1 && toks.every((tok) => tok.kind === 'integerLiteral' || tok.kind === 'operator' || tokenText(tok) === 'mod')
			? evaluateIntegerConstantExpression(toks.map((tok) => tok.rawText).join(' '), { get: () => undefined })
			: undefined;
		return folded === undefined ? undefined : String(folded);
	}
	if (rest[0].kind === 'integerLiteral') {
		const parsed = parseVbaIntegerLiteral(rest[0].rawText);
		return parsed === undefined ? undefined : String(sign * parsed);
	}
	if (rest[0].kind === 'floatLiteral') {
		const parsed = Number(rest[0].rawText.replace(/[!#@]$/, ''));
		return Number.isFinite(parsed) ? String(sign * parsed) : undefined;
	}
	return undefined;
}

export function simpleTypeNameForAssignment(type: string): string | undefined {
	const trimmed = type.replace(/\s*\(\s*\)\s*$/, '').trim();
	return IDENT_RE.test(trimmed) ? trimmed : undefined;
}

/**
 * Host types the model returns where the type library returns another, so a
 * value of the model's type is also a value of the library's. Excel's library
 * types the Charts and Worksheets properties of Application and Workbook as
 * Sheets, and a Sheets object is what they return at run time. The model
 * returns its own Charts and Worksheets there, whose members completion offers
 * (issue #90).
 */
const HOST_VALUES_ALSO_OF_TYPE: ReadonlyMap<string, string> = new Map([
	['excel.charts', 'excel.sheets'],
	['excel.worksheets', 'excel.sheets'],
]);

/**
 * The reason a scalar value cannot be Set: a compile error, where every other
 * reason is an object of the wrong class, which compiles and raises 13 when
 * the Set runs (issue #202, measured in Excel 16.0).
 */
export const SCALAR_OBJECT_ASSIGNMENT_REASON = 'An object assignment requires an object value.';

export function objectAssignmentIncompatibilityReason(
	expectedRaw: string | undefined,
	actual: InferredArgumentType | undefined,
	memberCtx: MemberCompletionContext,
	resolveType: typeof resolveKnownObjectAssignmentType = resolveKnownObjectAssignmentType,
	shareInterfaces?: (expected: string, actual: string) => boolean,
	implementsType?: typeof implementsObjectType,
): string | undefined {
	const expected = resolveType(expectedRaw, memberCtx);
	if (!expected || !actual) {
		return undefined;
	}
	const actualType = normalizeType(actual.type);
	if (!actualType || actualType === 'variant' || actualType === 'nothing') {
		return undefined;
	}
	if (isKnownScalarType(actualType)) {
		return SCALAR_OBJECT_ASSIGNMENT_REASON;
	}
	// Object takes any object and could be any. VBA's Collection is a class
	// like any other here: `Set c = New Square` into a Collection, or
	// `Set q = New Collection` into a Square, raises 13 (issue #202).
	if (expected.kind === 'generic' && expected.key === 'object') {
		return undefined;
	}
	const actualObject = resolveType(actual.type, memberCtx);
	if (!actualObject) {
		// A Scripting object CreateObject made is no Collection, sheet or class
		// of the project: 13 (issue #685, measured in Excel 16.0).
		const scripting = /^scripting\.(\w+)$/.exec(actualType);
		if (scripting) {
			return expected.key !== actualType && expected.key !== scripting[1] ? `This object type is not compatible with ${expected.display}.` : undefined;
		}
		// ActiveSheet, a Worksheet or a Chart: anything else refuses it (issue #685).
		return actualType === 'worksheet or chart' && !/(^|\.)(worksheet|chart)$/.test(expected.key) ? `This object type is not compatible with ${expected.display}.` : undefined;
	}
	if (actualObject.kind === 'generic' && actualObject.key === 'object') {
		return undefined;
	}
	// DAO's types are known only by their default members, not by what each
	// one implements: a Recordset2 is a Recordset.
	if (expected.key.startsWith('dao.') || actualObject.key.startsWith('dao.')) {
		return undefined;
	}
	if (expected.key === actualObject.key) {
		return undefined;
	}
	if (expected.kind === 'generic' || actualObject.kind === 'generic') {
		// A project class that implements Collection can stand in for one.
		const project = actualObject.kind === 'project' ? actualObject : expected.kind === 'project' ? expected : undefined;
		return project && (implementsType
			? implementsType(project, { kind: 'generic', display: 'Collection', key: 'collection' })
			: project.implements.some((name) => name.toLowerCase() === 'collection'))
			? undefined
			: `This object type is not compatible with ${expected.display}.`;
	}
	if (actualObject.kind === 'host' && HOST_VALUES_ALSO_OF_TYPE.get(actualObject.key) === expected.key) {
		return undefined;
	}
	if (actualObject.kind === 'project' && (implementsType ?? implementsObjectType)(actualObject, expected)) {
		return undefined;
	}
	// A Set between two class types is checked when it runs, by QueryInterface,
	// so it compiles whenever the object could support the target (issue #109).
	// The class an interface is implemented by can hold the interface's value
	// (`Set c = o`, casting back), and two interfaces one class implements can
	// hold each other's (`Set b = o`). Only project interfaces are known here.
	if (expected.kind === 'project' && actualObject.kind === 'project'
		&& projectTypesCanShareInstance(expected, actualObject, memberCtx, shareInterfaces, implementsType)) {
		return undefined;
	}
	return `This object type is not compatible with ${expected.display}.`;
}

/**
 * Find common implementers against metadata stable for one analysis pass.
 * Retain only examined memberships and queried pairs, rather than all pairs
 * that a surface with many implemented interfaces could form.
 */
export function createProjectInterfaceSharingLookup(
	memberCtx: MemberCompletionContext,
): (expected: string, actual: string) => boolean {
	let memberships: Map<string, Set<number>> | undefined;
	let pairs: Map<string, Map<string, boolean>> | undefined;
	let nextIndex = 0;
	return (expected, actual) => {
		// Shared-implementer membership is symmetric; direct casts stay outside.
		const first = expected < actual ? expected : actual;
		const second = expected < actual ? actual : expected;
		pairs ??= new Map();
		let answers = pairs.get(first);
		if (answers?.has(second)) {
			return answers.get(second)!;
		}
		if (!answers) {
			answers = new Map();
			pairs.set(first, answers);
		}
		memberships ??= new Map();
		const left = memberships.get(first), right = memberships.get(second);
		if (left && right) {
			const smaller = left.size <= right.size ? left : right;
			const larger = smaller === left ? right : left;
			for (const index of smaller) {
				if (larger.has(index)) {
					answers.set(second, true);
					return true;
				}
			}
		}
		const surfaces = memberCtx.projectClassMembers ?? [];
		while (nextIndex < surfaces.length) {
			const index = nextIndex++;
			// Keep every kind: the original shared-interface scan did not filter kinds.
			for (const name of surfaces[index].implements ?? []) {
				const lower = name.toLowerCase();
				let owners = memberships.get(lower);
				if (!owners) {
					owners = new Set();
					memberships.set(lower, owners);
				}
				owners.add(index);
			}
			if (memberships.get(first)?.has(index) && memberships.get(second)?.has(index)) {
				answers.set(second, true);
				return true;
			}
		}
		answers.set(second, false);
		return false;
	};
}

/**
 * Whether one project class can carry a value declared as the other: the
 * expected class implements the actual type (a cast from an interface back to
 * the class), or some project class implements both (a cast between two
 * interfaces of one object).
 */
function projectTypesCanShareInstance(
	expected: Extract<KnownObjectAssignmentType, { kind: 'project' }>,
	actual: Extract<KnownObjectAssignmentType, { kind: 'project' }>,
	memberCtx: MemberCompletionContext,
	shareInterfaces?: (expected: string, actual: string) => boolean,
	implementsType?: typeof implementsObjectType,
): boolean {
	if ((implementsType ?? implementsObjectType)(expected, actual)) {
		return true;
	}
	if (shareInterfaces) {
		return shareInterfaces(expected.key, actual.key);
	}
	const wanted = new Set([expected.key, actual.key]);
	for (const projectType of memberCtx.projectClassMembers ?? []) {
		const implemented = new Set((projectType.implements ?? []).map((name) => name.toLowerCase()));
		if ([...wanted].every((name) => implemented.has(name))) {
			return true;
		}
	}
	return false;
}

export function implementsObjectType(
	actual: Extract<KnownObjectAssignmentType, { kind: 'project' }>,
	expected: KnownObjectAssignmentType,
): boolean {
	const expectedNames = objectTypeExpectedNames(expected);
	return actual.implements.some((implemented) => {
		const lower = implemented.toLowerCase();
		return expectedNames.has(lower) || expectedNames.has(`excel.${lower}`);
	});
}

/** Index direct interface memberships only when consulted, against one pass's metadata. */
export function createObjectTypeImplementationLookup(): typeof implementsObjectType {
	const namesByList = new WeakMap<readonly string[], ReadonlySet<string>>();
	return (actual, expected) => {
		let names = namesByList.get(actual.implements);
		if (!names) {
			const indexed = new Set<string>();
			// Like some, forEach skips sparse array holes.
			actual.implements.forEach((name) => indexed.add(name.toLowerCase()));
			namesByList.set(actual.implements, indexed);
			names = indexed;
		}
		for (const name of objectTypeExpectedNames(expected)) {
			if (names.has(name) || (name.startsWith('excel.') && names.has(name.slice(6)))) {
				return true;
			}
		}
		return false;
	};
}

function objectTypeExpectedNames(expected: KnownObjectAssignmentType): ReadonlySet<string> {
	const expectedNames = new Set([expected.key]);
	const simple = simpleTypeNameForAssignment(expected.display);
	if (simple) {
		expectedNames.add(simple.toLowerCase());
	}
	const expectedLastSegment = expected.key.split('.').pop();
	if (expectedLastSegment) {
		expectedNames.add(expectedLastSegment);
	}
	return expectedNames;
}

// One-way proof only: strings with digits are left unknown until VBA conversion
// semantics are modeled explicitly.
/** Whether no locale converts the string to a number (see stringConversion.ts). */
export function isProvablyNonNumericString(value: string): boolean {
	return isInvalidNumericString(value);
}

export function stringLiteralValue(raw: string): string {
	return raw
		.replace(/^"/, '')
		.replace(/"$/, '')
		.replace(/""/g, '"');
}

// Per-pass memo (read-only by the engine's derived-table convention): this is
// rebuilt per procedure inside sourceNameScopeFor, but is a pure function of the
// module symbols, so compute it once per parse and share it across members/rules.
const MODULE_NON_CALLABLE_SYMBOLS = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	Map<string, VbaSymbol>
>();

export function moduleNonCallableSymbols(
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, VbaSymbol> {
	const cached = MODULE_NON_CALLABLE_SYMBOLS.get(symbols);
	if (cached) {
		return cached;
	}
	const out = new Map<string, VbaSymbol>();
	const callableNames = new Set(
		(symbols.root.children ?? [])
			.filter((sym) => isProcedureKind(sym.kind) || sym.kind === 'declare')
			.map((sym) => sym.name.toLowerCase()),
	);
	for (const sym of symbols.root.children ?? []) {
		if (isNonCallableSymbol(sym) && !callableNames.has(sym.name.toLowerCase())) {
			out.set(sym.name.toLowerCase(), sym);
		}
		if (sym.kind === 'enum') {
			for (const child of sym.children ?? []) {
				if (!callableNames.has(child.name.toLowerCase())) {
					out.set(child.name.toLowerCase(), child);
				}
			}
		}
	}
	MODULE_NON_CALLABLE_SYMBOLS.set(symbols, out);
	return out;
}

const SAME_MODULE_TYPE_NAMES = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	ReadonlySet<string>
>();

/**
 * Lowercased names of `Type` (struct) declarations in this module, memoized per
 * parse. Shared by the operand rules (`non-scalar-binary-operand`,
 * `argument-shape-mismatch`) so each does not independently rescan `root.children`.
 */
export function sameModuleTypeNames(
	symbols: ReturnType<typeof buildModuleSymbols>,
): ReadonlySet<string> {
	const cached = SAME_MODULE_TYPE_NAMES.get(symbols);
	if (cached) {
		return cached;
	}
	const names = new Set<string>();
	for (const child of symbols.root.children ?? []) {
		if (child.kind === 'type') {
			names.add(child.name.toLowerCase());
		}
	}
	SAME_MODULE_TYPE_NAMES.set(symbols, names);
	return names;
}

export function isNonCallableSymbol(sym: VbaSymbol): boolean {
	return (
		sym.kind === 'parameter' ||
		sym.kind === 'localVariable' ||
		sym.kind === 'moduleVariable' ||
		sym.kind === 'constant' ||
		sym.kind === 'enum' ||
		sym.kind === 'enumMember' ||
		sym.kind === 'type'
	);
}