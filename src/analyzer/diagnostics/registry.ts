// Diagnostics rule registry (audit #0).
//
// One ordered table of every active rule, each entry adapting the shared
// per-pass context to the rule function's signature. `runRules` in
// analyzeModule.ts is a loop over this array; the entry order is the
// engine's historical invocation order and is part of the public behavior
// (diagnostic output order), so append new rules thoughtfully and never
// reorder entries without updating the snapshot expectations in
// tests/vbaDiagnostics.test.ts and tests/diagnostics/.

import type { PushFn, RulePassContext } from './analysisContext';
import type { ProcedureStatementVisitor } from './walker';
import {
	checkInvalidLineContinuations,
	checkUnterminatedStrings,
} from './rules/lexical';
import {
	checkAmbiguousEnumMemberReferences,
	checkDuplicateDeclarations,
	checkDuplicateEnumMembers,
	checkDuplicateModuleMembers,
	checkDuplicateProcedures,
	checkDuplicateTypeFields,
} from './rules/duplicates';
import {
	checkDimInitializer,
	checkFixedLengthStringBounds,
	checkInvalidAsTypeNames,
	checkInvalidIdentifierStarts,
	checkModuleDeclarationsAfterProcedures,
	checkModuleDeclarationsInProcedureBodies,
	checkModuleLevelStatementsOutsideProcedures,
	checkNonConstantConstValues,
	checkNonConstantEnumMemberValues,
	checkNonConstantParameterDefaults,
	checkOptionPlacement,
	checkParameterDefaultValues,
	checkParameterOrder,
	checkProcedureHeader,
	checkPropertyAccessorSignatures,
	checkPropertySetterValueParameters,
	checkReservedDeclarationNames,
	checkTypeDeclarationCharacterAsClause,
	checkUnexpectedDeclarationTokens,
} from './rules/declarations';
import { checkArgumentCount } from './rules/callArity';
import { checkArgumentTypes } from './rules/argumentTypes';
import {
	checkRuntimeArgumentValues,
	checkRuntimeConversionValues,
} from './rules/runtimeValues';
import {
	checkAssignmentTypes,
	checkConstAssignment,
	checkMissingReturnAssignments,
	checkSetAssignments,
} from './rules/assignments';
import {
	checkObjectVariableNotSet,
	checkScalarMemberAccess,
} from './rules/objectState';
import { checkTypeOfIsCompatibility } from './rules/typeOfIs';
import {
	checkMemberNotFound,
	checkNonCallableCallStatement,
	checkOptionExplicit,
	checkUndeclaredVariables,
	checkUnknownCallStatement,
} from './rules/undeclared';
import {
	checkArrayBoundIntrinsicArguments,
	checkEraseTargets,
	checkInvalidRedimTargets,
	checkRedimImpossibleBounds,
	checkRedimPreserveDimensions,
	checkUnallocatedDynamicArrayAccess,
} from './rules/arrays';
import {
	checkDeclarePtrSafeForWin64,
	checkEventDeclarationModuleKind,
	checkEventHandlerModuleScope,
	checkFriendDeclarations,
	checkImplementsStatementPlacement,
	checkObjectModulePublicMembers,
	checkRaiseEventTargets,
	checkWithEventsDeclarations,
} from './rules/moduleKind';
import {
	checkCallParens,
	checkDivisionByZeroExpressions,
	checkExpressionCallParens,
	checkInvalidExpressionSyntax,
	checkUnbalancedParens,
} from './rules/expressions';
import {
	checkDuplicateLabels,
	checkElseBranchOrder,
	checkExitStatements,
	checkForEachLoopTypes,
	checkStatementContext,
	checkUndefinedLabels,
} from './rules/controlFlow';

/**
 * One registered rule: a stable name plus exactly one execution form.
 *
 * - `run` rules own their full traversal (module-level rules, rules with
 *   cross-member state, and rules whose internal walk order is part of their
 *   output order).
 * - `procedureStatements` rules are per-statement: the factory does the
 *   rule's per-pass setup and returns a visitor for the ONE shared
 *   procedure-statement walk (audit #0), instead of each rule walking the
 *   AST itself.
 *
 * Every rule reports through its own buffered `push`, and runRules flushes
 * the buffers in registry order, so both forms preserve the engine's
 * historical rule-major diagnostic order.
 */
export interface DiagnosticRuleEntry {
	name: string;
	run?(ctx: RulePassContext, push: PushFn): void;
	procedureStatements?(ctx: RulePassContext, push: PushFn): ProcedureStatementVisitor;
}

/**
 * Every active rule in invocation order. Rules are independent: each only
 * reads the shared context and reports through its own `push`, so an entry
 * can be understood (and profiled) in isolation.
 */
export const DIAGNOSTIC_RULE_REGISTRY: readonly DiagnosticRuleEntry[] = [
	{
		name: 'unterminatedStrings',
		run: (ctx, push) => checkUnterminatedStrings(ctx.source, push),
	},
	{
		name: 'invalidLineContinuations',
		run: (ctx, push) => checkInvalidLineContinuations(ctx.source, push),
	},
	{
		name: 'duplicateProcedures',
		run: (ctx, push) => checkDuplicateProcedures(ctx.symbols.root.children ?? [], push),
	},
	{
		name: 'duplicateDeclarations',
		run: (ctx, push) => checkDuplicateDeclarations(ctx.symbols.root.children ?? [], push),
	},
	{
		name: 'duplicateModuleMembers',
		run: (ctx, push) => checkDuplicateModuleMembers(ctx.symbols.root.children ?? [], push),
	},
	{
		name: 'duplicateEnumMembers',
		run: (ctx, push) => checkDuplicateEnumMembers(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'duplicateTypeFields',
		run: (ctx, push) => checkDuplicateTypeFields(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'ambiguousEnumMemberReferences',
		run: (ctx, push) => checkAmbiguousEnumMemberReferences(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.activity,
			ctx.moduleName,
			ctx.opts.knownProcedures,
			ctx.opts.projectProcedures,
			ctx.opts.projectClassMembers,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'constAssignment',
		procedureStatements: (ctx, push) => checkConstAssignment(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'optionExplicit',
		run: (ctx, push) => checkOptionExplicit(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'undeclaredVariables',
		run: (ctx, push) => checkUndeclaredVariables(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.activity,
			ctx.opts.knownIdentifiers,
			ctx.opts.projectProcedures,
			ctx.opts.projectClassMembers,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'optionPlacement',
		run: (ctx, push) => checkOptionPlacement(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'procedureHeader',
		run: (ctx, push) => checkProcedureHeader(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'invalidIdentifierStarts',
		run: (ctx, push) => checkInvalidIdentifierStarts(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'moduleDeclarationsInProcedureBodies',
		run: (ctx, push) => checkModuleDeclarationsInProcedureBodies(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'moduleDeclarationsAfterProcedures',
		run: (ctx, push) => checkModuleDeclarationsAfterProcedures(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'moduleLevelStatementsOutsideProcedures',
		run: (ctx, push) => checkModuleLevelStatementsOutsideProcedures(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'reservedDeclarationNames',
		run: (ctx, push) => checkReservedDeclarationNames(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'propertySetterValueParameters',
		run: (ctx, push) => checkPropertySetterValueParameters(ctx.source, ctx.mod, ctx.activity, ctx.opts, push),
	},
	{
		name: 'propertyAccessorSignatures',
		run: (ctx, push) => checkPropertyAccessorSignatures(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'parameterOrder',
		run: (ctx, push) => checkParameterOrder(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'parameterDefaultValues',
		run: (ctx, push) => checkParameterDefaultValues(ctx.source, ctx.mod, ctx.activity, ctx.memberCtx, push),
	},
	{
		name: 'parameterDefaultNotConstant',
		run: (ctx, push) => checkNonConstantParameterDefaults(ctx.source, ctx.mod, ctx.activity, ctx.memberCtx, push),
	},
	{
		name: 'constValueNotConstant',
		run: (ctx, push) => checkNonConstantConstValues(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'enumMemberNotConstant',
		run: (ctx, push) => checkNonConstantEnumMemberValues(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'unbalancedParens',
		run: (ctx, push) => checkUnbalancedParens(ctx.source, push),
	},
	{
		name: 'invalidExpressionSyntax',
		procedureStatements: (ctx, push) => checkInvalidExpressionSyntax(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'divisionByZeroExpressions',
		procedureStatements: (ctx, push) => checkDivisionByZeroExpressions(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectIntegerConstants,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
		),
	},
	{
		name: 'dimInitializer',
		run: (ctx, push) => checkDimInitializer(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'invalidRedimTargets',
		procedureStatements: (ctx, push) => checkInvalidRedimTargets(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
		),
	},
	{
		name: 'redimImpossibleBounds',
		procedureStatements: (ctx, push) => checkRedimImpossibleBounds(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'redimPreserveDimensions',
		run: (ctx, push) => checkRedimPreserveDimensions(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'unallocatedDynamicArrayAccess',
		run: (ctx, push) => checkUnallocatedDynamicArrayAccess(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'eraseTargets',
		procedureStatements: (ctx, push) => checkEraseTargets(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'typeDeclarationCharacterAsClause',
		run: (ctx, push) => checkTypeDeclarationCharacterAsClause(ctx.mod, ctx.activity, push),
	},
	{
		name: 'unexpectedDeclarationTokens',
		run: (ctx, push) => checkUnexpectedDeclarationTokens(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'fixedLengthStringBounds',
		run: (ctx, push) => checkFixedLengthStringBounds(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'objectModulePublicMembers',
		run: (ctx, push) => checkObjectModulePublicMembers(ctx.source, ctx.mod, ctx.moduleKind, ctx.activity, push),
	},
	{
		name: 'eventDeclarationModuleKind',
		run: (ctx, push) => checkEventDeclarationModuleKind(ctx.source, ctx.mod, ctx.moduleKind, ctx.activity, push),
	},
	{
		name: 'withEventsDeclarations',
		run: (ctx, push) => checkWithEventsDeclarations(ctx.source, ctx.mod, ctx.moduleKind, ctx.activity, push),
	},
	{
		name: 'friendDeclarations',
		run: (ctx, push) => checkFriendDeclarations(ctx.source, ctx.mod, ctx.moduleKind, ctx.activity, push),
	},
	{
		name: 'implementsStatementPlacement',
		run: (ctx, push) => checkImplementsStatementPlacement(ctx.source, ctx.mod, ctx.moduleKind, ctx.activity, push),
	},
	{
		name: 'raiseEventTargets',
		run: (ctx, push) => checkRaiseEventTargets(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'declarePtrSafeForWin64',
		run: (ctx, push) => checkDeclarePtrSafeForWin64(
			ctx.source,
			ctx.mod,
			ctx.opts.conditionalCompilation,
			ctx.activity,
			push,
		),
	},
	{
		name: 'eventHandlerModuleScope',
		run: (ctx, push) => checkEventHandlerModuleScope(
			ctx.source,
			ctx.mod,
			ctx.moduleName,
			ctx.moduleKind,
			ctx.opts.documentType,
			ctx.activity,
			push,
		),
	},
	{
		name: 'invalidAsTypeNames',
		run: (ctx, push) => checkInvalidAsTypeNames(ctx.source, ctx.mod, ctx.activity, ctx.opts, push),
	},
	{
		name: 'callParens',
		procedureStatements: (ctx, push) => checkCallParens(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectVisibleSymbols,
			ctx.memberCtx,
			push,
		),
	},
	{
		name: 'expressionCallParens',
		procedureStatements: (ctx, push) => checkExpressionCallParens(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'setAssignments',
		procedureStatements: (ctx, push) => checkSetAssignments(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			ctx.memberCtx,
			push,
		),
	},
	{
		name: 'exitStatements',
		procedureStatements: (ctx, push) => checkExitStatements(ctx.source, push),
	},
	{
		name: 'duplicateLabels',
		run: (ctx, push) => checkDuplicateLabels(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'undefinedLabels',
		run: (ctx, push) => checkUndefinedLabels(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'elseBranchOrder',
		run: (ctx, push) => checkElseBranchOrder(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'statementContext',
		run: (ctx, push) => checkStatementContext(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'forEachLoopTypes',
		run: (ctx, push) => checkForEachLoopTypes(ctx.mod, ctx.symbols, ctx.opts, ctx.activity, push),
	},
	{
		name: 'arrayBoundIntrinsicArguments',
		procedureStatements: (ctx, push) => checkArrayBoundIntrinsicArguments(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'scalarMemberAccess',
		procedureStatements: (ctx, push) => checkScalarMemberAccess(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'objectVariableNotSet',
		run: (ctx, push) => checkObjectVariableNotSet(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.memberCtx,
			ctx.activity,
			push,
		),
	},
	{
		name: 'memberNotFound',
		procedureStatements: (ctx, push) => checkMemberNotFound(ctx.source, ctx.memberCtx, push),
	},
	{
		name: 'nonCallableCallStatement',
		procedureStatements: (ctx, push) => checkNonCallableCallStatement(
			ctx.source,
			ctx.symbols,
			ctx.opts.knownProcedures,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'argumentCount',
		procedureStatements: (ctx, push) => checkArgumentCount(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectVisibleSymbols,
			ctx.memberCtx,
			push,
		),
	},
	{
		name: 'argumentTypes',
		procedureStatements: (ctx, push) => checkArgumentTypes(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectVisibleSymbols,
			ctx.memberCtx,
			push,
		),
	},
	{
		name: 'runtimeArgumentValues',
		procedureStatements: (ctx, push) => checkRuntimeArgumentValues(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectIntegerConstants,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
		),
	},
	{
		name: 'runtimeConversionValues',
		procedureStatements: (ctx, push) => checkRuntimeConversionValues(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			push,
		),
	},
	{
		name: 'assignmentTypes',
		run: (ctx, push) => checkAssignmentTypes(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			ctx.memberCtx,
			ctx.activity,
			push,
		),
	},
	{
		name: 'typeOfIsAlwaysFalse',
		run: (ctx, push) => checkTypeOfIsCompatibility(
			ctx.mod,
			ctx.symbols,
			ctx.memberCtx,
			ctx.activity,
			push,
		),
	},
	{
		name: 'missingReturnAssignments',
		run: (ctx, push) => checkMissingReturnAssignments(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.activity,
			push,
		),
	},
	{
		// Cross-module rule: only runs when the caller supplied the project's
		// visible procedure names (see AnalyzeModuleOptions.knownProcedures).
		name: 'unknownCallStatement',
		procedureStatements: (ctx, push) => {
			const knownProcedures = ctx.opts.knownProcedures;
			if (!knownProcedures) {
				return () => undefined;
			}
			return checkUnknownCallStatement(
				ctx.source,
				ctx.symbols,
				knownProcedures,
				ctx.opts.projectVisibleSymbols,
				push,
			);
		},
	},
];
