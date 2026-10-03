// Rule family: call-argument types (audit #0).
//
// Extracted verbatim from analyzeModule.ts: declared-signature argument-type
// validation for the same call surface the arity family walks.

import type { ConditionalActivityTracker } from '../../conditional/conditionalCompilation';
import { heldObjectsAt } from '../heldObjects';
import type { MemberCompletionContext } from '../../completion/memberAccess';
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
	expressionCalls,
	memberExpressionCalls,
	memberStatementCalls,
	sourceBindingTypeResolvers,
	sourceNameScopeFor,
	typeEnvironmentFor,
	validateArgumentTypes,
	validateArgumentTypesForSignature,
} from '../typeInference';
import { statementAndBranchSpans, tokenText, type ProcedureStatementVisitor } from '../walker';
import { straightLineAssignments } from '../straightLineValues';

/**
 * Rule: when both a callable parameter type and an argument type are known, flag
 * high-confidence mismatches. This first slice is deliberately conservative:
 * unknowns and Variant are accepted, and VBA's normal coercions are allowed
 * unless a literal is clearly incompatible (for example `"blah"` for Currency).
 *
 * Per-statement rule: rides the shared procedure-statement walk (audit #0).
 */
export function checkArgumentTypes(
	source: string,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
	activity?: ConditionalActivityTracker,
): ProcedureStatementVisitor {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	return (member) => {
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		const procSym = procedureSymbolFor(symbols, member);
		const { resolveExpressionType, resolveQualifiedExpressionType } =
			sourceBindingTypeResolvers(symbols, procSym, projectVisibleSymbols);
		// What a local holds at the statement (issue #246).
		const heldAt = heldObjectsAt(source, member, symbols, activity);
		const variantLocals = new Set((procSym?.children ?? [])
			.filter((child) => child.kind === 'localVariable' && child.visibility !== 'Static' && !child.isArray && (!child.asType || child.asType.toLowerCase() === 'variant'))
			.map((child) => child.name.toLowerCase()));
		let reaching: ReturnType<typeof straightLineAssignments> | undefined;
		return (stmt) => {
			const heldClassOf = (lower: string): string | undefined => heldAt(stmt).classes.get(lower);
			// A Variant local a straight line has just given Null (issue #324).
			const heldNull = (lower: string): boolean => {
				if (!variantLocals.has(lower)) {
					return false;
				}
				const held = (reaching ??= straightLineAssignments(source, member.body, activity)).get(stmt)?.get(lower)?.filter((tok) => tok.kind !== 'comment');
				return held?.length === 1 && tokenText(held[0]) === 'null';
			};
			// `Call Two(Nothing, 1)` is found both as an expression call and as
			// the statement's call; report each argument once (issue #223).
			const reported = new Set<string>();
			const pushOnce: PushFn = (code, message, span, data) => {
				const key = `${span.start}:${span.end}:${message}`;
				if (!reported.has(key)) {
					reported.add(key);
					push(code, message, span, data);
				}
			};
			for (const call of expressionCalls(source, stmt.span, moduleSignatures, sourceNames)) {
				validateArgumentTypes(
					call,
					env,
					moduleSignatures,
					sourceNames,
					source,
					memberCtx,
					pushOnce,
					resolveExpressionType,
					resolveQualifiedExpressionType,
					heldClassOf,
					heldNull,
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
					pushOnce,
					resolveExpressionType,
					resolveQualifiedExpressionType,
					heldClassOf,
					heldNull,
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
					pushOnce,
					resolveExpressionType,
					resolveQualifiedExpressionType,
					heldClassOf,
					heldNull,
				);
			}
			// A single-line If's branch is a statement call too: `If x Then Sl Nothing` (issue #254).
			for (const span of statementAndBranchSpans(stmt)) {
				const statementCall = extractCall(source, span) ?? extractQualifiedCall(source, span, moduleSignatures);
				if (statementCall) {
					validateArgumentTypes(
						statementCall,
						env,
						moduleSignatures,
						sourceNames,
						source,
						memberCtx,
						pushOnce,
						resolveExpressionType,
						resolveQualifiedExpressionType,
						heldClassOf,
						heldNull,
					);
				}
			}
		};
	};
}
