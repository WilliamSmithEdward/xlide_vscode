// Rule family: call-argument types (audit #0).
//
// Extracted verbatim from analyzeModule.ts: declared-signature argument-type
// validation for the same call surface the arity family walks.

import type { MemberCompletionContext } from '../../completion/memberAccess';
import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import type { ModuleNode } from '../../parser/nodes';
import { buildModuleSymbols } from '../../symbols/buildModuleSymbols';
import type {
	VbaProcedureSignature,
	VbaSymbol,
} from '../../symbols/symbolModel';
import {
	procedureSymbolFor,
	type PushFn,
} from '../analysisContext';
import {
	extractCall,
	extractQualifiedCall,
} from '../callExtraction';
import {
	callableTypeSignaturesFor,
	declaredValueTypeForQualifiedSourceBinding,
	declaredValueTypeForSourceBinding,
	expressionCalls,
	memberExpressionCalls,
	memberStatementCalls,
	sourceNameScopeFor,
	typeEnvironmentFor,
	validateArgumentTypes,
	validateArgumentTypesForSignature,
} from '../typeInference';
import {
	activeModuleMembers,
	forEachStatement,
} from '../walker';

/**
 * Rule: when both a callable parameter type and an argument type are known, flag
 * high-confidence mismatches. This first slice is deliberately conservative:
 * unknowns and Variant are accepted, and VBA's normal coercions are allowed
 * unless a literal is clearly incompatible (for example `"blah"` for Currency).
 */
export function checkArgumentTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const resolveExpressionType = (name: string) => declaredValueTypeForSourceBinding(
			symbols,
			procSym,
			projectVisibleSymbols,
			name,
		);
		const resolveQualifiedExpressionType = (qualifier: string, name: string) =>
			declaredValueTypeForQualifiedSourceBinding(
				symbols,
				projectVisibleSymbols,
				qualifier,
				name,
			);
		forEachStatement(member.body, (stmt) => {
			for (const call of expressionCalls(source, stmt.span, moduleSignatures, sourceNames)) {
				validateArgumentTypes(
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
			for (const memberCall of memberExpressionCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				validateArgumentTypesForSignature(
					memberCall.signature,
					memberCall.call,
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
			for (const memberCall of memberStatementCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				validateArgumentTypesForSignature(
					memberCall.signature,
					memberCall.call,
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
			const statementCall = extractCall(source, stmt.span);
			const qualifiedStatementCall = statementCall
				? undefined
				: extractQualifiedCall(source, stmt.span, moduleSignatures);
			const effectiveStatementCall = statementCall ?? qualifiedStatementCall;
			if (effectiveStatementCall) {
				validateArgumentTypes(
					effectiveStatementCall,
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
		}, activity);
	}
}
