// Expression type inference and callable signature tables for diagnostics.
//
// Extracted verbatim from `analyzeModule.ts`: the per-module/project callable
// signature tables, procedure type environments and source-name scopes, the
// expression/atomic-expression type-inference engine, runtime signature
// parsing, expression-level call extraction (expressionCalls and the member
// call binders), argument-type validation, and the type-compatibility tables.
// Pure analysis: the only diagnostics emitted here flow through the PushFn
// passed by the argument-type rules.

import type { VbaToken } from '../lexer/tokenKinds';
import { matchParenFrom } from '../lexer/tokenHelpers';
import {
	parseDecimalIntegerLiteral,
	type IntegerConstantLookup,
} from '../constants/integerConstantExpression';
import {
	getHostMembers,
	resolveHostAlias,
	resolveHostConstant,
	resolveHostGlobal,
} from '../host/hostModel';
import {
	resolveRuntimeConstant,
	resolveRuntimeFunction,
	resolveRuntimeObject,
	type VbaRuntimeFunction,
} from '../runtime/vbaRuntime';
import { standaloneEmptyParenthesizedCallStatement } from '../call/callContext';
import type { ProcedureNode, Span } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
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
} from '../completion/memberAccess';
import { procedureSymbolFor, type PushFn } from './analysisContext';
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
	externalIntegerConstantValue,
	numericExternalConstantValue,
} from './constExpr';
import {
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
			out.set(symbol.name.toLowerCase(), callableTypeSignatureFromSymbol(symbol));
		}
	}
	MODULE_TYPE_SIGNATURES.set(symbols, out);
	return out;
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
		const signature = callableTypeSignatureFromSymbol(symbol);
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
	WeakMap<ProcedureNode, Map<string, string>>
>();
const DECLARATION_SHAPE_ENVIRONMENTS = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	WeakMap<ProcedureNode, Map<string, DeclaredValueShape>>
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

export function typeEnvironmentFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): Map<string, string> {
	const cache = perProcedureCache(TYPE_ENVIRONMENTS, symbols);
	const cached = cache.get(proc);
	if (cached) {
		return cached;
	}
	const out = new Map<string, string>();
	for (const sym of symbols.root.children ?? []) {
		if (sym.asType && !isProcedureKind(sym.kind)) {
			out.set(sym.name.toLowerCase(), sym.asType);
		}
	}
	const procSym = procedureSymbolFor(symbols, proc);
	const returnType = returnAssignmentTypeFor(proc);
	if (returnType) {
		out.set(proc.name.toLowerCase(), returnType);
	}
	for (const child of procSym?.children ?? []) {
		if (child.asType) {
			out.set(child.name.toLowerCase(), child.asType);
		}
	}
	cache.set(proc, out);
	return out;
}

export interface DeclaredValueShape {
	asType?: string;
	isArray: boolean;
	isFixedArray: boolean;
}

export function declarationShapeEnvironmentFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): Map<string, DeclaredValueShape> {
	const cache = perProcedureCache(DECLARATION_SHAPE_ENVIRONMENTS, symbols);
	const cached = cache.get(proc);
	if (cached) {
		return cached;
	}
	const out = new Map<string, DeclaredValueShape>();
	for (const sym of symbols.root.children ?? []) {
		if (isValueDeclarationSymbol(sym)) {
			out.set(sym.name.toLowerCase(), {
				asType: sym.asType,
				isArray: sym.isArray === true,
				isFixedArray: sym.arrayBounds !== undefined,
			});
		}
	}
	const procSym = procedureSymbolFor(symbols, proc);
	const returnType = returnAssignmentTypeFor(proc);
	if (returnType) {
		out.set(proc.name.toLowerCase(), {
			asType: returnType,
			isArray: returnAssignmentIsArray(proc),
			isFixedArray: false,
		});
	}
	for (const child of procSym?.children ?? []) {
		if (isValueDeclarationSymbol(child)) {
			out.set(child.name.toLowerCase(), {
				asType: child.asType,
				isArray: child.isArray === true,
				isFixedArray: child.arrayBounds !== undefined,
			});
		}
	}
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
	const typed = binding.definitions.find((definition) => definition.asType);
	return { resolved: true, asType: typed?.asType };
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
	const typed = valueDefinitions.find((definition) => definition.asType);
	return { resolved: true, asType: typed?.asType };
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
	const typed = matchingValues.find((definition) => definition.asType);
	return { resolved: true, asType: typed?.asType };
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
		(definition) => definition.asType || definition.isArray === true,
	);
	return {
		resolved: true,
		shape: {
			asType: shaped?.asType,
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
	callableShadows: ReadonlySet<string>;
	/**
	 * Any source-backed identifier visible in the current procedure. These block
	 * runtime fallback once source/project callable signatures have not resolved.
	 */
	runtimeShadows: ReadonlySet<string>;
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
	const callableShadows = new Set(moduleNonCallableSymbols(symbols).keys());
	const procSym = procedureSymbolFor(symbols, proc);
	const runtimeShadows = sourceIdentifierNames({
		currentModule: symbols,
		enclosingProcedure: procSym,
		projectVisibleSymbols,
	});
	for (const child of procSym?.children ?? []) {
		const lower = child.name.toLowerCase();
		if (isNonCallableSymbol(child)) {
			callableShadows.add(lower);
		}
	}
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

export function scopedIntegerConstantLookup(
	constants: ReadonlyMap<string, number | undefined>,
	symbols: ReturnType<typeof buildModuleSymbols>,
	procSym: VbaSymbol | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
): IntegerConstantLookup {
	return {
		get(name: string): number | undefined {
			const key = name.toLowerCase();
			if (key.includes('.')) {
				if (constants.has(key)) {
					return constants.get(key);
				}
				return externalIntegerConstantValue(key);
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
				return externalIntegerConstantValue(key);
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
): InferredArgumentType | undefined {
	if (runtimeCallableSourceShadowed(name, sourceNames)) {
		return undefined;
	}
	const candidates = [
		inferredExternalConstant(name, resolveRuntimeConstant(name)),
		inferredExternalConstant(name, resolveHostConstant(name)),
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
	const hostType = memberCtx
		? resolveHostGlobal(name, memberCtx.model)
		: resolveHostGlobal(name);
	if (hostType) {
		return { type: hostType, label: `${name} As ${hostType}`, span };
	}
	const runtimeObject = resolveRuntimeObject(name);
	if (runtimeObject) {
		return { type: runtimeObject.type, label: `${name} As ${runtimeObject.type}`, span };
	}
	return undefined;
}

export function inferQualifiedExternalConstantExpressionType(
	qualifier: string,
	name: string,
	span: Span,
): InferredArgumentType | undefined {
	const lower = qualifier.toLowerCase();
	if (lower === 'vba') {
		const inferred = inferredExternalConstant(`${qualifier}.${name}`, resolveRuntimeConstant(name));
		return inferred ? { ...inferred, span } : undefined;
	}
	if (lower === 'excel' || lower === 'office') {
		const inferred = inferredExternalConstant(`${qualifier}.${name}`, resolveHostConstant(name));
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
	return sourceNames?.runtimeShadows.has(name.toLowerCase()) === true;
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
	for (let i = 0; i < toks.length - 1; i++) {
		const callName = parenthesizedCallNameAt(toks, i);
		if (!callName) {
			continue;
		}
		const { name, parenIndex, nameEndIndex } = callName;
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
		});
	}
	return out;
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
		const signature = parseRuntimeDisplaySignature(member.name, member.signature);
		if (isPropertyResultIndexing(member, signature, inner)) {
			continue;
		}
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
		const byRefMismatch = byRefVariableTypeMismatch(
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
		const actual = inferArgumentType(
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
		if (!actual) {
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
	if (!isKnownByRefExactType(expected)) {
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
		actualRaw = declaredType?.resolved
			? declaredType.asType
			: env.get(name.toLowerCase());
		span = { start: sliceStart + toks[0].start, end: sliceStart + toks[0].end };
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
	if (!isKnownByRefExactType(actual) || actual === expected) {
		return undefined;
	}
	return {
		name,
		actual: actualRaw ?? name,
		span,
	};
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
	const runtime = resolveRuntimeFunction(name);
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

export function runtimeSignatureParameterText(signature: string): string | undefined {
	const open = signature.indexOf('(');
	const close = signature.lastIndexOf(')');
	if (open < 0 || close < open) {
		return undefined;
	}
	return signature.slice(open + 1, close);
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
	const as = /\bAs\s+([A-Za-z_][A-Za-z0-9_]*(?:\(\))?)/i.exec(text);
	const first = /[A-Za-z_][A-Za-z0-9_]*/.exec(text)?.[0];
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

export function splitSignatureTopLevel(text: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let start = 0;
	for (let i = 0; i < text.length; i++) {
		const c = text[i];
		if (c === '(' || c === '[') {
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
				return {
					type: 'Double',
					label: `numeric literal ${first.rawText}`,
					span,
					numericValue,
					numericText: first.rawText,
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
		const external = inferBareExternalConstantExpressionType(name, span, sourceNames);
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
		if (callName) {
			const sig = callableSignatureFor(callName.name, moduleSignatures, sourceNames);
			if (sig?.returnType && matchParenFrom(toks, callName.parenIndex) === toks.length - 1) {
				return {
					type: sig.returnType,
					label: `${callName.name}(...) As ${sig.returnType}`,
					span: { start: span.start, end: sliceStart + toks[callName.nameEndIndex].end },
				};
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
	if (
		member.kind !== 'method' &&
		member.returns &&
		argumentTokens &&
		argumentTokens.length > 0 &&
		(!member.signature || parseRuntimeDisplaySignature(member.name, member.signature).params.length === 0)
	) {
		return defaultHostItemReturnType(member.returns, memberCtx) ?? member.returns;
	}
	return member.returns ?? 'Variant';
}

export function defaultHostItemReturnType(
	typeName: string,
	memberCtx: MemberCompletionContext,
): string | undefined {
	const item = getHostMembers(typeName, memberCtx.model).find(
		(member) => member.name.toLowerCase() === 'item',
	);
	if (item?.returns) {
		return item.returns;
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
		if (actualType === 'string') {
			return actual.stringValue !== undefined && isProvablyNonNumericString(actual.stringValue)
				? "This string literal cannot be converted to a numeric value. This will raise Run-time error '13': Type mismatch."
				: undefined;
		}
		return undefined;
	}
	if (expected === 'boolean') {
		if (actualType === 'boolean' || isNumericType(actualType)) {
			return undefined;
		}
		if (actualType === 'string') {
			return actual.stringValue !== undefined && isBooleanString(actual.stringValue)
				? undefined
				: "This string literal cannot be converted to Boolean. This will raise Run-time error '13': Type mismatch.";
		}
		return undefined;
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
		return undefined;
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
	return `The numeric literal ${literal} is outside the ${bounds.label} range ${bounds.min} to ${bounds.max}. This will raise Run-time error '6': Overflow.`;
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

export function isNumericType(type: string): boolean {
	return new Set([
		'byte',
		'integer',
		'long',
		'longlong',
		'longptr',
		'single',
		'double',
		'currency',
		'decimal',
	]).has(type);
}

export function isKnownScalarType(type: string): boolean {
	return type === 'string' || type === 'boolean' || type === 'date' || isNumericType(type);
}

export function isKnownObjectAssignmentType(
	type: string | undefined,
	memberCtx: MemberCompletionContext,
): boolean {
	return resolveKnownObjectAssignmentType(type, memberCtx) !== undefined;
}

export type KnownObjectAssignmentType =
	| { kind: 'generic'; display: string; key: 'object' }
	| { kind: 'host'; display: string; key: string }
	| { kind: 'project'; display: string; key: string; implements: readonly string[] };

export function resolveKnownObjectAssignmentType(
	type: string | undefined,
	memberCtx: MemberCompletionContext,
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
	const host = resolveHostAlias(type, memberCtx.model);
	if (host) {
		return { kind: 'host', display: type, key: host.toLowerCase() };
	}
	const simple = simpleTypeNameForAssignment(type);
	if (!simple) {
		return undefined;
	}
	const lower = simple.toLowerCase();
	const matches = (memberCtx.projectClassMembers ?? []).filter(
		(projectType) =>
			projectType.kind !== 'userType' &&
			projectType.kind !== 'standardModule' &&
			projectType.name.toLowerCase() === lower,
	);
	if (matches.length !== 1) {
		return undefined;
	}
	return {
		kind: 'project',
		display: matches[0].name,
		key: lower,
		implements: matches[0].implements ?? [],
	};
}

export function simpleTypeNameForAssignment(type: string): string | undefined {
	const trimmed = type.replace(/\s*\(\s*\)\s*$/, '').trim();
	return /^[A-Za-z_][A-Za-z0-9_]*$/.test(trimmed) ? trimmed : undefined;
}

export function objectAssignmentIncompatibilityReason(
	expectedRaw: string | undefined,
	actual: InferredArgumentType | undefined,
	memberCtx: MemberCompletionContext,
): string | undefined {
	const expected = resolveKnownObjectAssignmentType(expectedRaw, memberCtx);
	if (!expected || !actual) {
		return undefined;
	}
	const actualType = normalizeType(actual.type);
	if (!actualType || actualType === 'variant' || actualType === 'nothing') {
		return undefined;
	}
	if (isKnownScalarType(actualType)) {
		return 'An object assignment requires an object value.';
	}
	if (expected.kind === 'generic') {
		return undefined;
	}
	const actualObject = resolveKnownObjectAssignmentType(actual.type, memberCtx);
	if (!actualObject) {
		return undefined;
	}
	if (actualObject.kind === 'generic') {
		return undefined;
	}
	if (expected.key === actualObject.key) {
		return undefined;
	}
	if (actualObject.kind === 'project' && implementsObjectType(actualObject, expected)) {
		return undefined;
	}
	return `This object type is not compatible with ${expected.display}.`;
}

export function implementsObjectType(
	actual: Extract<KnownObjectAssignmentType, { kind: 'project' }>,
	expected: KnownObjectAssignmentType,
): boolean {
	const expectedNames = new Set([expected.key]);
	const simple = simpleTypeNameForAssignment(expected.display);
	if (simple) {
		expectedNames.add(simple.toLowerCase());
	}
	const expectedLastSegment = expected.key.split('.').pop();
	if (expectedLastSegment) {
		expectedNames.add(expectedLastSegment);
	}
	return actual.implements.some((implemented) => {
		const lower = implemented.toLowerCase();
		return expectedNames.has(lower) || expectedNames.has(`excel.${lower}`);
	});
}

// One-way proof only: strings with digits are left unknown until VBA conversion
// semantics are modeled explicitly.
export function isProvablyNonNumericString(value: string): boolean {
	const trimmed = value.trim();
	return trimmed.length > 0 && !/[0-9]/.test(trimmed);
}

export function stringLiteralValue(raw: string): string {
	return raw
		.replace(/^"/, '')
		.replace(/"$/, '')
		.replace(/""/g, '"');
}

export function isBooleanString(value: string): boolean {
	return /^(true|false|0|-?1)$/i.test(value.trim());
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