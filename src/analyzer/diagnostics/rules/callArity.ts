import {sourceSetterAssignment} from '../setterAssignment';
import {procedureSymbolFor} from '../analysisContext';
// Rule family: call-argument arity (audit #0).
//
// Extracted verbatim from analyzeModule.ts: wrong-number-of-arguments
// validation for every callable form one statement can contain.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import { VBA_LIBRARY_CONTAINERS, VBA_LIBRARY_NAMES } from '../../runtime/vbaLibraryNames';
import { resolveRuntimeFunction, VBA_RUNTIME_FUNCTIONS } from '../../runtime/vbaRuntime';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import {
	qualifiedProcedureKey,
	type VbaProcedureSignature,
	type VbaSymbol,
} from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import {
	type CallableTypeSignature,
	type CallArguments,
	extractCall,
	extractQualifiedCall,
	isNamedSlot,
	validateArity,
} from '../callExtraction';
import {
	bareCallableSourceShadowed,
	callableTypeSignaturesFor,
	expressionCalls,
	memberExpressionCalls,
	memberStatementCalls,
	normalizeType,
	parseRuntimeDisplaySignature,
	resolveExactMemberCompletion,
	runtimeAritySignature,
	runtimeCallableSourceShadowed,
	sameModuleCallableSignatures,
	type SourceNameScope,
	sourceNameScopeFor,
	typeEnvironmentFor,
	uniqueProjectTypeSignatures,
} from '../typeInference';
import { matchParenFrom, statementAndBranchSpans, statementTokens, tokenName, type ProcedureStatementVisitor } from '../walker';
import { splitTopLevelTokenGroups } from '../../lexer/tokenHelpers';
import type { VbaToken } from '../../lexer/tokenKinds';
import { resolveHostGlobalMember } from '../../host/hostModel';
import { syntaxErrorLineSpan } from '../../parser/physicalLineSpans';

type ArityTables = {
	sameModuleSignatures: ReadonlyMap<string, readonly CallableTypeSignature[]>;
	projectSignatures: ReadonlyMap<string, CallableTypeSignature>;
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>;
};
const QUALIFIED_ARITY_TABLES = new WeakMap<object, ArityTables>();
const QUALIFIER_VALUE_BINDINGS = new WeakMap<object, WeakMap<object, ReadonlySet<string>>>();
const PROJECT_QUALIFIER_NAMES = new WeakMap<object, ReadonlySet<string>>();
function projectQualifierNames(symbols: readonly VbaSymbol[] | undefined): ReadonlySet<string> {
	if (!symbols) { return EMPTY_QUALIFIER_NAMES; }
	const cached = PROJECT_QUALIFIER_NAMES.get(symbols);
	if (cached) { return cached; }
	const names = new Set(symbols.map((symbol) => symbol.moduleName.toLowerCase()));
	PROJECT_QUALIFIER_NAMES.set(symbols, names);
	return names;
}
const EMPTY_QUALIFIER_NAMES = new Set<string>();
// Exact bare intrinsic arities measured in the 2026-10-19 Excel batch.
// Other spellings/counts retain the existing generic arity diagnostic.
const BARE_INTRINSIC_SYNTAX_ARITIES: ReadonlyMap<string, ReadonlySet<number>> = new Map([
	['clng', new Set([0, 2])], ['cint', new Set([0])], ['len', new Set([0])],
]);
const BARE_RUNTIME_MISSING_ARGUMENT = new Set(['left$']);
const NO_CONTEXT_MEMBERS: readonly unknown[] = [];
function qualifierValueBindings(context: MemberCompletionContext): ReadonlySet<string> {
	const classes = context.projectClassMembers ?? NO_CONTEXT_MEMBERS;
	const implicit = context.implicitMembers ?? NO_CONTEXT_MEMBERS;
	let byImplicit = QUALIFIER_VALUE_BINDINGS.get(classes);
	if (!byImplicit) { QUALIFIER_VALUE_BINDINGS.set(classes, byImplicit = new WeakMap()); }
	const cached = byImplicit.get(implicit);
	if (cached) { return cached; }
	const names = new Set<string>();
	for (const member of context.implicitMembers ?? []) { names.add(member.name.toLowerCase()); }
	for (const type of context.projectClassMembers ?? []) {
		const kind = type.kind;
		if (kind === 'document' || kind === 'userform' || kind === 'class' && type.predeclaredId !== false) { names.add(type.name.toLowerCase()); }
	}
	byImplicit.set(implicit, names);
	return names;
}

/**
 * Rule: a call to a known Sub/Function/Declare must supply an argument count the
 * procedure's parameter list accepts. Same-module procedures come directly from
 * this module's AST. Cross-module checks use the ProjectIndex signature map:
 * bare exported names are checked only when unique, and module-qualified calls
 * resolve through the named standard module only. Parenthesized object member
 * calls are checked only when the shared member-completion binder resolves a
 * known source or host/reference signature. Ambiguous or unresolved targets stay
 * silent to remain false-positive-free.
 *
 * The inspected forms are the parenless call statement (`Foo 1, 2`), the
 * explicit `Call Foo(1, 2)`, and parenthesized current-module calls inside
 * expressions (`x = Foo(1, 2)`) or member access (`Application.Calculate()`).
 *
 * Per-statement rule: rides the shared procedure-statement walk (audit #0).
 */
export function checkArgumentCount(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
	implicitMembers: MemberCompletionContext['implicitMembers'] = memberCtx.implicitMembers,
): ProcedureStatementVisitor {
	// Binding precedence needs designer names without expanding the metadata
	// used to validate ordinary host/control member invocations.
	const bindingContext = { projectClassMembers: memberCtx.projectClassMembers, implicitMembers };
	const baseSameModuleSignatures = sameModuleCallableSignatures(symbols);
	const baseProjectSignatures = uniqueProjectTypeSignatures(projectProcedures);
	const baseModuleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const qualifiedTables = () => {
		const cached = QUALIFIED_ARITY_TABLES.get(baseModuleSignatures);
		if (cached) { return cached; }
		const sameModuleSignatures = new Map(baseSameModuleSignatures);
		const projectSignatures = new Map(baseProjectSignatures);
		const moduleSignatures = new Map(baseModuleSignatures);
		if (symbols.moduleKind === 'standard') {
			for (const [name, signatures] of baseSameModuleSignatures) {
				const key = qualifiedProcedureKey(symbols.moduleName, name);
				moduleSignatures.set(key, signatures[0]);
				sameModuleSignatures.set(key, signatures);
			}
		}
		for (const runtime of VBA_RUNTIME_FUNCTIONS) {
			const signature = runtimeAritySignature(runtime);
			if (!signature || !VBA_LIBRARY_NAMES.has(runtime.name.toLowerCase())) { continue; }
			const keys = [qualifiedProcedureKey('VBA', runtime.name)];
			for (const [container, members] of VBA_LIBRARY_CONTAINERS) {
				if (members.has(runtime.name.toLowerCase())) {
					keys.push(qualifiedProcedureKey(container, runtime.name), `vba.${qualifiedProcedureKey(container, runtime.name)}`);
				}
			}
			for (const key of keys) {
				if (!moduleSignatures.has(key)) {
					moduleSignatures.set(key, signature);
					projectSignatures.set(key, signature);
				}
			}
		}
		const tables = { sameModuleSignatures, projectSignatures, moduleSignatures };
		QUALIFIED_ARITY_TABLES.set(baseModuleSignatures, tables);
		return tables;
	};
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const env = typeEnvironmentFor(symbols, member);
		const shadowedQualifier = (name: string) => sourceNames.runtimeShadows.has(name)
			|| qualifierValueBindings(bindingContext).has(name)
			|| name !== symbols.moduleName.toLowerCase() && projectQualifierNames(projectVisibleSymbols).has(name);
		// One offset index per token snapshot, shared with branch call checks.
		const tokenIndexes = new WeakMap<readonly VbaToken[], ReadonlyMap<number, number>>();
		const tokenIndex = (tokens: readonly VbaToken[], offset: number): number => {
			let indexes = tokenIndexes.get(tokens);
			if (!indexes) {
				indexes = new Map(tokens.map((token, index) => [token.start, index]));
				tokenIndexes.set(tokens, indexes);
			}
			return indexes.get(offset) ?? -1;
		};
		const provenCall = (call: CallArguments | undefined, tokens: readonly VbaToken[], start: number): CallArguments | undefined => {
			if (!call) { return undefined; }
			if (!call.qualifier) {
				// Designer/default instances shadow runtime aliases, without mutating
				// the shared source scope or scanning project bindings for other names.
				if (!resolveRuntimeFunction(call.name) || baseSameModuleSignatures.has(call.name.toLowerCase())
					|| sourceNames.runtimeShadows.has(call.name.toLowerCase())) { return call; }
				const token = tokens[tokenIndex(tokens, call.nameSpan.start - start)];
				if (token?.kind === 'keyword' && token.rawText.toLowerCase() === 'len') { return call; }
				return qualifierValueBindings(bindingContext).has(call.name.toLowerCase()) ? undefined : call;
			}
			const qualifier = call.qualifier.toLowerCase();
			const container = VBA_LIBRARY_CONTAINERS.has(qualifier);
			const index = tokenIndex(tokens, call.nameSpan.start - start);
			if (index < 2) { return undefined; }
			if (container && tokens[index - 3]?.rawText === '.' && tokenName(tokens[index - 4])?.toLowerCase() === 'vba'
				&& !['.', '!'].includes(tokens[index - 5]?.rawText) && !shadowedQualifier('vba')
				&& VBA_LIBRARY_CONTAINERS.get(qualifier)?.has(call.name.toLowerCase())) {
				return { ...call, qualifier: `VBA.${call.qualifier}`, lookupKey: `vba.${call.lookupKey}` };
			}
			// Declared receivers win over project module names, including library
			// container collisions. runtimeShadows also contains module identities,
			// so use the narrower value scope before accepting source signatures.
			if (sourceNames.callableShadows.has(qualifier) || qualifierValueBindings(bindingContext).has(qualifier)) { return undefined; }
			if (qualifier !== 'vba' && qualifier !== symbols.moduleName.toLowerCase() && !container) { return call; }
			if (qualifier === symbols.moduleName.toLowerCase() && !baseSameModuleSignatures.has(call.name.toLowerCase())) { return undefined; }
			if (['.', '!'].includes(tokens[index - 3]?.rawText)) { return undefined; }
			// A proven source module retains its own signature resolution.
			if (container && baseModuleSignatures.has(call.lookupKey!)) { return call; }
			return shadowedQualifier(qualifier) ? undefined : call;
		};
		return (stmt) => {
			const tokens = statementTokens(source, stmt.span);
			for (let i = 0; i + 1 < tokens.length; i++) {
				// Len is reserved syntax even when a designer exposes a Len member.
				// Brackets and member qualification select that member instead.
				if (tokens[i].kind === 'keyword' && tokens[i].rawText.toLowerCase() === 'len'
					&& tokens[i + 1].rawText.toLowerCase() === 'is' && !['.', '!'].includes(tokens[i - 1]?.rawText ?? '')) {
					push('argumentCount', "Len requires an expression. This is a VBE compile error: Syntax error.", syntaxErrorLineSpan(source, { start: stmt.span.start + tokens[i].start, end: stmt.span.start + tokens[i].end }));
				}
			}
			const needsQualifiedTables = tokens.some((token, index) => token.rawText === '.'
				&& (tokenName(tokens[index - 1])?.toLowerCase() === 'vba'
					|| tokenName(tokens[index - 1])?.toLowerCase() === symbols.moduleName.toLowerCase()
					|| VBA_LIBRARY_CONTAINERS.has(tokenName(tokens[index - 1])?.toLowerCase() ?? '')));
			const { sameModuleSignatures, projectSignatures, moduleSignatures } = needsQualifiedTables ? qualifiedTables()
				: { sameModuleSignatures: baseSameModuleSignatures, projectSignatures: baseProjectSignatures, moduleSignatures: baseModuleSignatures };
			for (const span of statementAndBranchSpans(stmt)) {
				checkUnmodelledArity(source, span, env, sourceNames, memberCtx, push);
			}
			const setterNames = new Set(statementAndBranchSpans(stmt).map(span => sourceSetterAssignment(source, span, symbols, procedureSymbolFor(symbols, member), projectVisibleSymbols, memberCtx)?.nameSpan.start).filter(start => start !== undefined));
			const projectQualifiedCallSpans = new Set<string>();
			const statementCall = extractCall(source, stmt.span);
			const qualifiedStatementCall = statementCall
				? undefined
				: extractQualifiedCall(source, stmt.span, moduleSignatures);
			const effectiveStatementCall = provenCall(statementCall ?? qualifiedStatementCall, tokens, stmt.span.start);
			if (effectiveStatementCall) {
				validateCallableArity(
					source,
					effectiveStatementCall,
					sameModuleSignatures,
					projectSignatures,
					sourceNames,
					push,
				);
				recordProjectQualifiedCallSpan(effectiveStatementCall, projectQualifiedCallSpans);
			}
			const expressionCallList = expressionCalls(source, stmt.span, moduleSignatures, sourceNames);
			for (const extracted of expressionCallList) {
				const call = provenCall(extracted, tokens, stmt.span.start);
				if (!call) { continue; }
				if (setterNames.has(call.nameSpan.start) || sameCallTarget(call, effectiveStatementCall)) {
					continue;
				}
				validateCallableArity(source, call, sameModuleSignatures, projectSignatures, sourceNames, push);
				recordProjectQualifiedCallSpan(call, projectQualifiedCallSpans);
			}
			for (const memberCall of memberExpressionCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				if (setterNames.has(memberCall.call.nameSpan.start) || projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call)) || takesPrintList(memberCall.signature)) {
					continue;
				}
				validateArity(source, memberCall.signature, memberCall.call, push);
			}
			for (const memberCall of memberStatementCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				if (setterNames.has(memberCall.call.nameSpan.start) || projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call)) || takesPrintList(memberCall.signature)) {
					continue;
				}
				validateArity(source, memberCall.signature, memberCall.call, push);
			}
			// A single-line If is one statement, so a CALL STATEMENT it carries -
			// `If ok Then Helper 1, 2, 3` - was never read as one and its arity went
			// unchecked (issue #46). Only the statement-call path repeats over the
			// branches: the expression scans above already cover the whole line, and
			// running them again would report the same call twice.
			for (const branch of statementAndBranchSpans(stmt).slice(1)) {
				const branchCall = provenCall(extractCall(source, branch)
					?? extractQualifiedCall(source, branch, moduleSignatures), statementTokens(source, branch), branch.start);
				if (branchCall && !projectQualifiedCallSpans.has(callTargetSpanKey(branchCall))) {
					validateCallableArity(
						source,
						branchCall,
						sameModuleSignatures,
						projectSignatures,
						sourceNames,
						push,
					);
					recordProjectQualifiedCallSpan(branchCall, projectQualifiedCallSpans);
				}
				for (const memberCall of memberStatementCalls(source, branch, memberCtx)) {
					if (setterNames.has(memberCall.call.nameSpan.start) || projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call)) || takesPrintList(memberCall.signature)) {
						continue;
					}
					validateArity(source, memberCall.signature, memberCall.call, push);
				}
			}
		};
	};
}

/** VBA's Collection methods, which no host model carries (issue #304). */
const COLLECTION_SIGNATURES: ReadonlyMap<string, string> = new Map([
	['add', 'Add(Item, [Key], [Before], [After])'],
	['item', 'Item(Index)'],
	['count', 'Count()'],
	['remove', 'Remove(Index)'],
]);

/** Excel's Global properties that take an index: Cells reaches Range.Item. */
const INDEXED_GLOBALS: ReadonlyMap<string, string> = new Map([
	['cells', 'Cells([RowIndex], [ColumnIndex])'],
	['range', 'Range(Cell1, [Cell2])'],
]);

const SCALAR_TYPES: ReadonlySet<string> = new Set(['long', 'integer', 'byte', 'double', 'single', 'currency', 'boolean', 'longlong']);

/**
 * The calls the signature tables above do not reach (issue #304, each
 * measured in Excel 16.0): a Collection's Add, Item, Count and Remove; a
 * bare Excel Global method such as Evaluate, Intersect or Union, and Cells
 * or Range given more than they take; and a host property that holds a
 * number, `Sheets.Count(1)`, given an argument.
 */
function checkUnmodelledArity(
	source: string,
	span: { start: number; end: number },
	env: ReadonlyMap<string, string>,
	sourceNames: SourceNameScope,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const toks = statementTokens(source, span);
	const validate = (signature: string, display: string, nameIndex: number, slots: VbaToken[][]): void => {
		const call: CallArguments = {
			name: display,
			nameSpan: { start: span.start + toks[nameIndex].start, end: span.start + toks[nameIndex].end },
			slots,
			sliceStart: span.start,
		};
		validateArity(source, parseRuntimeDisplaySignature(display, signature), call, push);
	};
	const argumentsAt = (open: number): VbaToken[][] | undefined => {
		const close = matchParenFrom(toks, open);
		if (close < 0) {
			return undefined;
		}
		return close === open + 1 ? [] : splitTopLevelTokenGroups(toks, open + 1, ',', close);
	};
	toks.forEach((tok, i) => {
		const name = tokenName(tok);
		if (!name) {
			return;
		}
		const lower = name.toLowerCase();
		const member = toks[i - 1]?.rawText === '.';
		// `c.Add 1`, `c.Item()`: a local As Collection's own methods.
		const receiver = member ? tokenName(toks[i - 2])?.toLowerCase() : undefined;
		if (member && receiver && toks[i - 3]?.rawText !== '.' && normalizeType(env.get(receiver)) === 'collection' && COLLECTION_SIGNATURES.has(lower)) {
			// A project class named Collection keeps its own members.
			if (resolveExactMemberCompletion(source, name, span.start + tok.end, memberCtx)?.definitions) {
				return;
			}
			const statement = i === 2 && toks[i + 1]?.rawText !== '(' && toks[i + 1]?.rawText !== '=';
			const slots = toks[i + 1]?.rawText === '(' ? argumentsAt(i + 1) : statement ? (i + 1 < toks.length ? splitTopLevelTokenGroups(toks, i + 1, ',', toks.length) : []) : undefined;
			if (slots) {
				validate(COLLECTION_SIGNATURES.get(lower)!, name, i, slots);
			}
			return;
		}
		if (toks[i + 1]?.rawText !== '(') {
			return;
		}
		// `Evaluate()`, `Union(r)`, `Cells(1, 1, 1)`: a bare Excel Global member.
		if (!member && !bareCallableSourceShadowed(name, sourceNames) && !runtimeCallableSourceShadowed(name, sourceNames) && !env.has(lower)) {
			const global = resolveHostGlobalMember(name, memberCtx.model);
			const signature = global?.kind === 'method' ? global.signature : global ? INDEXED_GLOBALS.get(lower) : undefined;
			const slots = signature ? argumentsAt(i + 1) : undefined;
			if (signature && slots) {
				validate(signature, name, i, slots);
			}
			return;
		}
		// `Sheets.Count(1)`: a host property that holds a number takes no argument.
		if (member) {
			const resolved = resolveExactMemberCompletion(source, name, span.start + tok.end, memberCtx);
			const type = normalizeType(resolved?.declaredType ?? resolved?.returns);
			const slots = argumentsAt(i + 1);
			if (resolved?.kind === 'property' && !resolved.signature && !resolved.definitions && resolved.letAccessor === undefined && type && SCALAR_TYPES.has(type) && slots && slots.length > 0) {
				validate(`${resolved.name}()`, name, i, slots);
			}
		}
	});
}

function recordProjectQualifiedCallSpan(call: CallArguments, out: Set<string>): void {
	if (call.lookupKey) {
		out.add(callTargetSpanKey(call));
	}
}

function callTargetSpanKey(call: CallArguments): string {
	return `${call.nameSpan.start}:${call.nameSpan.end}`;
}

function validateCallableArity(
	source: string,
	call: CallArguments,
	sameModuleSignatures: ReadonlyMap<string, readonly CallableTypeSignature[]>,
	projectSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
	push: PushFn,
): void {
	const lower = call.lookupKey ?? call.name.toLowerCase();
	const check = (signature: CallableTypeSignature, report: PushFn): void => {
		const missing = signature.params.find((param, index) => index >= call.slots.length && !param.optional && !param.paramArray);
		validateArity(source, signature, call, (key, message, span, data) => {
			if (key !== 'argumentCount' || !message.startsWith('Wrong number of arguments')) { report(key, message, span, data); return; }
			if (missing && (call.qualifier || signature.runtime && call.slots.length === 0 && BARE_RUNTIME_MISSING_ARGUMENT.has(signature.name.toLowerCase()))) {
				// Qualified runtime and source calls share the measured required-
				// parameter message and dot-prefixed selection (CLng and Foo).
				const name = call.qualifier ? `${call.qualifier}.${signature.name}` : signature.name;
				report(key, `Argument not optional: '${missing.name}' is required by '${name}'.`,
					call.qualifier ? { start: source.lastIndexOf('.', call.nameSpan.start), end: call.nameSpan.end } : call.nameSpan, data);
				return;
			}
			// Bare intrinsic grammar has separately measured syntax verdicts.
			if (signature.runtime && !call.qualifier && !call.slots.some(isNamedSlot)
				&& BARE_INTRINSIC_SYNTAX_ARITIES.get(signature.name.toLowerCase())?.has(call.slots.length)
				&& /^[ \t]*\(/.test(source.slice(call.nameSpan.end))) {
				report(key, `The intrinsic '${signature.name}' requires its expression argument. This is a VBE compile error: Syntax error.`, syntaxErrorLineSpan(source, call.nameSpan), data);
				return;
			}
			if (signature.runtime && call.qualifier && call.slots.length > signature.params.length && !signature.params.some((param) => param.paramArray)) {
				report(key, `Wrong number of arguments or invalid property assignment. ${message}`,
					{ start: source.lastIndexOf('.', call.nameSpan.start), end: call.nameSpan.end }, data);
				return;
			}
			report(key, message, span, data);
		});
	};
	if (!call.qualifier && bareCallableSourceShadowed(call.name, sourceNames)) {
		return;
	}
	const candidates = sameModuleSignatures.get(lower);
	if (candidates) {
		if (candidates.length === 1) {
			check(candidates[0], push);
			return;
		}
		// Several same-module signatures share the name. Since XLIDE cannot say
		// which one a build compiles, it can still say that NONE of them accepts
		// this call, which is wrong under every build. Declaring a procedure
		// once per arm of a `#If` chain is the shape that makes this common, and
		// it used to turn arity checking off for every call to that name
		// (github.com/WilliamSmithEdward/xlide_vscode/issues/58).
		const rejections = candidates.map((signature) => {
			const hits: Parameters<PushFn>[] = [];
			check(signature, (...args) => { hits.push(args); });
			return hits;
		});
		if (rejections.every((hits) => hits.length > 0)) {
			push(...rejections[0][0]);
		}
		return;
	}
	const projectSignature = projectSignatures.get(lower);
	if (projectSignature) {
		check(projectSignature, push);
		return;
	}
	if (!call.qualifier) {
		if (runtimeCallableSourceShadowed(call.name, sourceNames)) {
			return;
		}
		const runtime = resolveRuntimeFunction(call.name);
		const runtimeSignature = runtime ? runtimeAritySignature(runtime) : undefined;
		if (runtimeSignature) {
			check(runtimeSignature, push);
		}
	}
}

function sameCallTarget(a: CallArguments, b: CallArguments | undefined): boolean {
	return !!b && a.nameSpan.start === b.nameSpan.start && a.nameSpan.end === b.nameSpan.end;
}

/**
 * A host method named Print, a VB6 form's or picture box's (issue #358) or an
 * Access report's, takes what the Print statement takes: `Form1.Print "a"; x`.
 * Its listed signature has no parameters, so its arity is not judged.
 */
function takesPrintList(signature: { name: string }): boolean {
	return signature.name.toLowerCase() === 'print';
}
