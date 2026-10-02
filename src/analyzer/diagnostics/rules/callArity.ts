// Rule family: call-argument arity (audit #0).
//
// Extracted verbatim from analyzeModule.ts: wrong-number-of-arguments
// validation for every callable form one statement can contain.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import { resolveRuntimeFunction } from '../../runtime/vbaRuntime';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
	VbaSymbol,
} from '../../symbols/symbolModel';
import type { PushFn } from '../analysisContext';
import {
	type CallableTypeSignature,
	type CallArguments,
	extractCall,
	extractQualifiedCall,
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
): ProcedureStatementVisitor {
	const sameModuleSignatures = sameModuleCallableSignatures(symbols);
	const projectSignatures = uniqueProjectTypeSignatures(projectProcedures);
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	return (member) => {
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const env = typeEnvironmentFor(symbols, member);
		return (stmt) => {
			for (const span of statementAndBranchSpans(stmt)) {
				checkUnmodelledArity(source, span, env, sourceNames, memberCtx, push);
			}
			const projectQualifiedCallSpans = new Set<string>();
			const statementCall = extractCall(source, stmt.span);
			const qualifiedStatementCall = statementCall
				? undefined
				: extractQualifiedCall(source, stmt.span, moduleSignatures);
			const effectiveStatementCall = statementCall ?? qualifiedStatementCall;
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
			for (const call of expressionCallList) {
				if (sameCallTarget(call, effectiveStatementCall)) {
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
				if (projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call)) || takesPrintList(memberCall.signature)) {
					continue;
				}
				validateArity(source, memberCall.signature, memberCall.call, push);
			}
			for (const memberCall of memberStatementCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				if (projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call)) || takesPrintList(memberCall.signature)) {
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
				const branchCall = extractCall(source, branch)
					?? extractQualifiedCall(source, branch, moduleSignatures);
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
					if (projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call)) || takesPrintList(memberCall.signature)) {
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
	const toks = statementTokens(source, span).filter((tok) => tok.kind !== 'comment');
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
	if (!call.qualifier && bareCallableSourceShadowed(call.name, sourceNames)) {
		return;
	}
	const candidates = call.qualifier
		? undefined
		: sameModuleSignatures.get(call.name.toLowerCase());
	if (candidates) {
		if (candidates.length === 1) {
			validateArity(source, candidates[0], call, push);
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
			validateArity(source, signature, call, (...args) => { hits.push(args); });
			return hits;
		});
		if (rejections.every((hits) => hits.length > 0)) {
			push(...rejections[0][0]);
		}
		return;
	}
	const projectSignature = projectSignatures.get(lower);
	if (projectSignature) {
		validateArity(source, projectSignature, call, push);
		return;
	}
	if (!call.qualifier) {
		if (runtimeCallableSourceShadowed(call.name, sourceNames)) {
			return;
		}
		const runtime = resolveRuntimeFunction(call.name);
		const runtimeSignature = runtime ? runtimeAritySignature(runtime) : undefined;
		if (runtimeSignature) {
			validateArity(source, runtimeSignature, call, push);
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
