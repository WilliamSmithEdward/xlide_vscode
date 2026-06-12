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
	runtimeAritySignature,
	runtimeCallableSourceShadowed,
	sameModuleCallableSignatures,
	type SourceNameScope,
	sourceNameScopeFor,
	uniqueProjectTypeSignatures,
} from '../typeInference';
import type { ProcedureStatementVisitor } from '../walker';

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
		return (stmt) => {
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
				if (projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call))) {
					continue;
				}
				validateArity(source, memberCall.signature, memberCall.call, push);
			}
			for (const memberCall of memberStatementCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				if (projectQualifiedCallSpans.has(callTargetSpanKey(memberCall.call))) {
					continue;
				}
				validateArity(source, memberCall.signature, memberCall.call, push);
			}
		};
	};
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
		// Skip ambiguous same-module targets where the signature is not unique.
		if (candidates.length === 1) {
			validateArity(source, candidates[0], call, push);
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
