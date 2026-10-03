// Diagnostics rule registry (audit #0).
//
// One ordered table of every active rule, each entry adapting the shared
// per-pass context to the rule function's signature. `runRules` in
// analyzeModule.ts is a loop over this array; the entry order is the
// engine's historical invocation order and is part of the public behavior
// (diagnostic output order), so append new rules thoughtfully and never
// reorder entries without updating the snapshot expectations in
// tests/vbaDiagnostics.test.ts and tests/diagnostics/.

import { ownObjectMemberNames, type PushFn, type RulePassContext } from './analysisContext';
import type { ProcedureStatementVisitor } from './walker';
import type { ProcedureExpressionVisitor } from './exprWalk';
import {
	checkInvalidLineContinuations,
	checkUnterminatedStrings,
} from './rules/lexical';
import {
	checkAmbiguousEnumMemberReferences,
	checkDuplicateDeclarations,
	checkDuplicateEnumMembers,
	checkDuplicateModuleMembers,
	checkAmbiguousBareProcedureCalls,
	checkDuplicateProcedures,
	checkDuplicateTypeFields,
	checkEnumMemberNameClash,
	checkVariableProcedureNameClash,
} from './rules/duplicates';
import {
	checkDimInitializer,
	checkDuplicateOptions,
	checkOptionStatementForm,
	checkEmptyType,
	checkFixedLengthStringBounds,
	checkIdentifierTooLong,
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
	checkModuleName,
	checkReservedDeclarationNames,
	checkTooManyParameters,
	checkTypeDeclarationCharacterAsClause,
	checkUdtParameterConstraints,
	checkUnexpectedDeclarationTokens,
} from './rules/declarations';
import { checkArgumentCount } from './rules/callArity';
import { checkOmittedArgumentReads } from './rules/omittedArguments';
import { checkArgumentTypes } from './rules/argumentTypes';
import { checkArgumentShape } from './rules/argumentShape';
import {
	checkRuntimeArgumentValues,
	checkRuntimeConversionValues,
} from './rules/runtimeValues';
import {
	checkAssignmentTypes,
	checkConstAssignment,
	checkMidStatementLiteralTarget,
	checkMissingReturnAssignments,
	checkSetAssignments,
} from './rules/assignments';
import {
	checkObjectVariableNotSet,
	checkScalarMemberAccess,
} from './rules/objectState';
import {
	checkIsOperandsInConditions,
	checkIsOperatorOperands,
	checkTypeOfIsCompatibility,
	checkTypeOfMissingOperand,
} from './rules/typeOfIs';
import { checkBinaryOperandScalar } from './rules/binaryOperandScalar';
import { checkSuffixedLiteralOverflow } from './rules/numericLiterals';
import {
	checkMemberNotFound,
	checkNonCallableCallStatement,
	checkOptionExplicit,
	checkBuiltinsReadBare,
	checkUndeclaredVariables,
	checkUnknownCallStatement,
} from './rules/undeclared';
import {
	checkArrayBoundIntrinsicArguments,
	checkArrayDeclarationBounds,
	checkEraseTargets,
	checkInvalidRedimTargets,
	checkRedimImpossibleBounds,
	checkRedimTypeChange,
	checkRedimPreserveDimensions,
	checkUnallocatedDynamicArrayAccess,
	checkFixedArraySubscriptBounds,
} from './rules/arrays';
import { checkTypeFieldArrays } from './rules/typeFieldArrays';
import { checkTypeMembers } from './rules/typeMembers';
import { checkDeclareStatements, checkUnusableDeclareCalls } from './rules/declares';
import { checkLateBoundFriendMember } from './rules/lateBinding';
import { checkHandlerFlow } from './rules/handlerFlow';
import { checkFileStatements } from './rules/fileStatements';
import { checkEmptyFilePaths } from './rules/filePaths';
import { checkInvalidPropertyUse } from './rules/propertyUse';
import { checkOverflow } from './rules/overflow';
import { checkHostArguments, workbookSheetsToCheck } from './rules/hostArguments';
import { checkCollectionLoopCounters, checkCollectionState } from './rules/collectionState';
import { checkDictionaryState } from './rules/dictionaryState';
import { checkDocumentNames } from './rules/documentNames';
import { checkExcelSessionState } from './rules/excelSessionState';
import { checkErrorValues } from './rules/errorValues';
import { checkAccessData } from './rules/accessData';
import { checkLateBoundObjects } from './rules/lateBoundObjects';
import { checkByNameCalls } from './rules/byNameCalls';
import { checkModuleMemberForms } from './rules/moduleMembers';
import { checkParamArrayUse } from './rules/paramArrayUse';
import { callableTypeSignaturesFor } from './typeInference';
import { checkVariantValueMisuse } from './rules/variantValues';
import { checkRuntimeMemberNotFound } from './rules/lateBoundMembers';
import { checkFormContents } from './rules/formContents';
import { checkConditionValues } from './rules/conditionValues';
import { checkLockedArrays } from './rules/lockedArrays';
import { checkDeletedObjects } from './rules/deletedObjects';
import { checkLongLongNarrowing } from './rules/longLongNarrowing';
import { checkAddressOfUse } from './rules/addressOfUse';
import { checkObjectDefaultValues } from './rules/objectValues';
import { checkEventHandlerSignatures } from './rules/eventHandlerSignatures';
import { checkDeclarationForms } from './rules/declarationForms';
import { checkLineContinuationLimits } from './rules/lineContinuations';
import { checkImplementsMembers } from './rules/implementsMembers';
import { checkStatementForms } from './rules/statementForms';
import { checkVbaLibraryMembers } from './rules/vbaLibraryMembers';
import { checkStrayCharacters } from './rules/strayTokens';
import { checkDirectiveForms } from './rules/directiveForms';
import { checkMalformedLines } from './rules/malformedLines';
import { checkParentheses } from './rules/parentheses';
import { checkMissingLibraryReference, checkMissingScriptingReference } from './rules/missingReference';
import { getExcelObjectModel } from '../host/excelObjectModel';
import {
	checkDeclarePtrSafeForWin64,
	checkEventDeclarationModuleKind,
	checkEventHandlerModuleScope,
	checkFriendDeclarations,
	checkImplementsStatementPlacement,
	checkMeOutsideObjectModule,
	checkObjectModulePublicMembers,
	checkRaiseEventArguments,
	checkRaiseEventTargets,
	checkWithEventsDeclarations,
} from './rules/moduleKind';
import {
	checkCallParens,
	checkDivisionByZeroExpressions,
	checkExpressionCallParens,
	checkInvalidExpressionSyntax,
	checkStringArithmeticOperands,
	checkUnbalancedParens,
} from './rules/expressions';
import {
	checkDuplicateCaseElse,
	checkDuplicateLabels,
	checkLineNumberRange,
	checkElseBranchOrder,
	checkElseWithoutIf,
	checkExitStatements,
	checkForEachLoopTypes,
	checkMalformedStatements,
	checkStatementContext,
	checkReservedLabels,
	checkUndefinedLabels,
} from './rules/controlFlow';
import { checkDeclarationOrder } from './rules/declarationOrder';
import { checkLocalDeclarationOrder } from './rules/localDeclarationOrder';
import { checkRefusedDeclarations } from './rules/refusedDeclarations';
import { checkStatementTypes } from './rules/statementTypes';
import {
	checkUnreachableCode,
	checkUnusedDeclarations,
	checkUnusedPrivateProcedures,
} from './rules/deadCode';
import { checkDocComments } from './rules/docComments';

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
 * - `procedureExpressions` rules are per-expression: the factory returns a
 *   per-member visitor for the ONE shared expression-tree walk, so the operand
 *   rules cost a single traversal per body between them rather than one each.
 *
 * Every rule reports through its own buffered `push`, and runRules flushes
 * the buffers in registry order, so all forms preserve the engine's
 * historical rule-major diagnostic order.
 */
export interface DiagnosticRuleEntry {
	name: string;
	run?(ctx: RulePassContext, push: PushFn): void;
	procedureStatements?(ctx: RulePassContext, push: PushFn): ProcedureStatementVisitor;
	/**
	 * The statement visitor also takes each block's header line as a
	 * statement: a For's bounds, a Select Case subject, a Do, Loop or While
	 * condition, a With subject (issue #233). For rules that judge an
	 * expression wherever it stands, and read no statement form.
	 */
	blockHeaders?: boolean;
	procedureExpressions?(ctx: RulePassContext, push: PushFn): ProcedureExpressionVisitor;
}

/**
 * `push` for a rule that reads the raw text rather than the parse: the VBE
 * does not lex an inactive `#If` branch, so an unclosed string or a
 * continuation into a blank line there compiles (issue #234, measured in
 * Excel 16.0).
 */
function activeOnly(ctx: RulePassContext, push: PushFn): PushFn {
	const activity = ctx.activity;
	return activity ? (code, message, span, ...rest) => {
		if (!activity.isInactive(span)) {
			push(code, message, span, ...rest);
		}
	} : push;
}

/**
 * Every active rule in invocation order. Rules are independent: each only
 * reads the shared context and reports through its own `push`, so an entry
 * can be understood (and profiled) in isolation.
 */
export const DIAGNOSTIC_RULE_REGISTRY: readonly DiagnosticRuleEntry[] = [
	{
		name: 'unterminatedStrings',
		run: (ctx, push) => checkUnterminatedStrings(ctx.source, activeOnly(ctx, push)),
	},
	{
		name: 'invalidLineContinuations',
		run: (ctx, push) => checkInvalidLineContinuations(ctx.source, activeOnly(ctx, push)),
	},
	{
		name: 'duplicateProcedures',
		run: (ctx, push) => checkDuplicateProcedures(ctx.symbols.root.children ?? [], ctx.activity, push),
	},
	{
		name: 'duplicateDeclarations',
		run: (ctx, push) => checkDuplicateDeclarations(ctx.symbols.root.children ?? [], ctx.activity, push),
	},
	{
		name: 'variableProcedureNameClash',
		run: (ctx, push) => checkVariableProcedureNameClash(ctx.symbols.root.children ?? [], ctx.activity, push),
	},
	{
		name: 'enumMemberNameClash',
		run: (ctx, push) => checkEnumMemberNameClash(ctx.symbols.root.children ?? [], ctx.activity, push),
	},
	{
		name: 'duplicateModuleMembers',
		run: (ctx, push) => checkDuplicateModuleMembers(ctx.symbols.root.children ?? [], ctx.activity, push),
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
		name: 'emptyType',
		run: (ctx, push) => checkEmptyType(ctx.mod, ctx.activity, push),
	},
	{
		name: 'refusedDeclarations',
		run: (ctx, push) => checkRefusedDeclarations(ctx.source, ctx.mod, ctx.moduleKind, ctx.activity, push),
	},
	{
		name: 'redimTypeChange',
		procedureStatements: (ctx, push) => checkRedimTypeChange(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'tooManyParameters',
		run: (ctx, push) => checkTooManyParameters(ctx.mod, ctx.moduleKind, ctx.activity, push),
	},
	{
		name: 'identifierTooLong',
		run: (ctx, push) => checkIdentifierTooLong(ctx.mod, ctx.activity, push),
	},
	{
		name: 'udtParameterConstraints',
		run: (ctx, push) => checkUdtParameterConstraints(ctx.mod, ctx.activity, push),
	},
	{
		name: 'ambiguousBareProcedureCalls',
		procedureStatements: (ctx, push) => checkAmbiguousBareProcedureCalls(
			ctx.source,
			ctx.symbols,
			ctx.moduleName,
			ctx.opts.projectProcedures,
			ctx.opts.projectVisibleSymbols,
			push,
		),
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
			ctx.opts.hostModel,
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
		run: (ctx, push) => checkOptionExplicit(ctx.mod, ctx.activity, push),
	},
	{
		name: 'localDeclarationOrder',
		run: (ctx, push) => checkLocalDeclarationOrder(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			ctx.opts.hostModel,
			ctx.activity,
			push,
		),
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
			ctx.opts.implicitMembers,
			ctx.opts.moduleKind,
			ctx.opts.hostModel,
			ctx.opts.designerClass,
			ctx.opts.referencedHosts,
			push,
			ownObjectMemberNames(ctx.opts),
		),
	},
	{
		name: 'builtinsReadBare',
		run: (ctx, push) => checkBuiltinsReadBare(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.activity,
			ctx.opts.projectVisibleSymbols,
			ctx.opts.moduleKind,
			ctx.opts.hostModel,
			ctx.opts.designerClass,
			ctx.opts.implicitMembers,
			push,
			ownObjectMemberNames(ctx.opts),
		),
	},
	{
		name: 'conditionValues',
		run: (ctx, push) => checkConditionValues(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'lockedArrays',
		run: (ctx, push) => checkLockedArrays(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'deletedObjects',
		run: (ctx, push) => checkDeletedObjects(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'handlerFlow',
		run: (ctx, push) => checkHandlerFlow(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'fileStatements',
		run: (ctx, push) => checkFileStatements(ctx.source, ctx.mod, ctx.activity, push, ctx.opts.projectOpenedFileNumbers),
	},
	{
		name: 'emptyFilePaths',
		procedureStatements: (ctx, push) => checkEmptyFilePaths(ctx.source, ctx.symbols, ctx.activity, ctx.opts.projectVisibleSymbols, push),
	},
	{
		name: 'overflow',
		run: (ctx, push) => checkOverflow(
			ctx.source, ctx.mod, ctx.symbols, ctx.opts.projectVisibleSymbols, ctx.opts.hostModel, ctx.activity, push,
		),
	},
	{
		name: 'hostArguments',
		blockHeaders: true,
		procedureStatements: (ctx, push) => checkHostArguments(ctx.source, ctx.symbols, ctx.memberCtx, ctx.activity, push, workbookSheetsToCheck(ctx.opts)),
	},
	{
		name: 'collectionState',
		run: (ctx, push) => checkCollectionState(ctx.source, ctx.mod, ctx.activity, push, ctx.symbols, ctx.opts.projectIntegerConstants, ctx.opts.projectVisibleSymbols, ctx.opts.hostModel),
	},
	{
		name: 'byNameCalls',
		run: (ctx, push) => checkByNameCalls(ctx.source, ctx.mod, ctx.symbols, ctx.memberCtx, ctx.opts.projectRunnableProcedures, ctx.activity, push),
	},
	{
		name: 'dictionaryState',
		run: (ctx, push) => checkDictionaryState(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'documentNames',
		run: (ctx, push) => checkDocumentNames(
			ctx.source,
			ctx.mod,
			ctx.opts.hostModel?.hostName,
			new Set([
				...(ctx.symbols.root.children ?? []).filter((symbol) => ['sub', 'function', 'declare', 'propertyGet', 'propertyLet', 'propertySet'].includes(symbol.kind)).map((symbol) => symbol.name.toLowerCase()),
				...[...(ctx.opts.projectProcedures?.keys() ?? [])].map((name) => name.toLowerCase()),
			]),
			ctx.activity,
			push,
		),
	},
	{
		name: 'excelSessionState',
		run: (ctx, push) => checkExcelSessionState(
			ctx.source,
			ctx.mod,
			new Set([
				...(ctx.symbols.root.children ?? []).filter((symbol) => ['sub', 'function', 'declare', 'propertyGet', 'propertyLet', 'propertySet'].includes(symbol.kind)).map((symbol) => symbol.name.toLowerCase()),
				...[...(ctx.opts.projectProcedures?.keys() ?? [])].map((name) => name.toLowerCase()),
			]),
			ctx.memberCtx,
			ctx.activity,
			push,
		),
	},
	{
		name: 'errorValues',
		run: (ctx, push) => checkErrorValues(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'accessData',
		run: (ctx, push) => checkAccessData(ctx.source, ctx.mod, ctx.symbols, ctx.opts.hostModel?.hostName, ctx.activity, push),
	},
	{
		name: 'lateBoundObjectState',
		run: (ctx, push) => checkLateBoundObjects(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'variantValueMisuse',
		run: (ctx, push) => checkVariantValueMisuse(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push, ctx.opts.projectVisibleSymbols),
	},
	{
		name: 'objectDefaultValue',
		procedureStatements: (ctx, push) => checkObjectDefaultValues(ctx.source, ctx.symbols, ctx.memberCtx, push, ctx.activity),
	},
	{
		name: 'runtimeMemberNotFound',
		run: (ctx, push) => checkRuntimeMemberNotFound(ctx.source, ctx.mod, ctx.symbols, ctx.memberCtx, ctx.activity, push),
	},
	{
		name: 'formContents',
		run: (ctx, push) => checkFormContents(ctx.source, ctx.mod, ctx.opts.implicitMembers, ctx.opts.projectNameMentions, ctx.activity, push),
	},
	{
		name: 'declarationForms',
		run: (ctx, push) => checkDeclarationForms(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'lineContinuationLimits',
		run: (ctx, push) => checkLineContinuationLimits(ctx.source, ctx.mod, activeOnly(ctx, push)),
	},
	{
		name: 'eventHandlerSignatures',
		run: (ctx, push) => checkEventHandlerSignatures(ctx.mod, ctx.moduleKind, ctx.opts, ctx.memberCtx, ctx.activity, push),
	},
	{
		name: 'implementsMembers',
		run: (ctx, push) => checkImplementsMembers(
			ctx.source, ctx.mod, ctx.symbols, ctx.moduleKind, ctx.opts.projectClassMembers, ctx.activity, push,
		),
	},
	{
		name: 'statementForms',
		run: (ctx, push) => checkStatementForms(ctx.source, ctx.mod, ctx.symbols, ctx.opts.projectProcedures, ctx.activity, push, ctx.memberCtx),
	},
	{
		name: 'vbaLibraryMembers',
		run: (ctx, push) => checkVbaLibraryMembers(ctx.source, ctx.mod, ctx.symbols, ctx.opts.projectVisibleSymbols, ctx.activity, push, ctx.opts.hostModel?.hostName),
	},
	{
		name: 'statementTypes',
		run: (ctx, push) => checkStatementTypes(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			ctx.opts.projectTypes,
			ctx.activity,
			push,
		),
	},
	{
		name: 'strayCharacters',
		run: (ctx, push) => checkStrayCharacters(ctx.source, ctx.activity, push),
	},
	{
		name: 'malformedLines',
		run: (ctx, push) => checkMalformedLines(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'parentheses',
		run: (ctx, push) => checkParentheses(ctx.source, ctx.mod, ctx.symbols, ctx.memberCtx, ctx.activity, push),
	},
	{
		name: 'directiveForms',
		run: (ctx, push) => checkDirectiveForms(ctx.source, ctx.mod, ctx.opts.conditionalCompilation, push),
	},
	{
		name: 'optionPlacement',
		run: (ctx, push) => checkOptionPlacement(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'duplicateOption',
		run: (ctx, push) => checkDuplicateOptions(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'optionStatementForm',
		run: (ctx, push) => checkOptionStatementForm(ctx.source, ctx.mod, ctx.opts, ctx.activity, push),
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
		name: 'moduleName',
		run: (ctx, push) => checkModuleName(ctx.source, ctx.opts.moduleName, push, ctx.opts.hostModel?.hostName),
	},
	{
		name: 'propertySetterValueParameters',
		run: (ctx, push) => checkPropertySetterValueParameters(ctx.source, ctx.mod, ctx.activity, push),
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
		run: (ctx, push) => checkUnbalancedParens(ctx.source, push, ctx.activity),
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
		name: 'stringArithmeticOperands',
		procedureStatements: (ctx, push) => checkStringArithmeticOperands(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.activity,
			push,
		),
	},
	{
		name: 'divisionByZeroExpressions',
		blockHeaders: true,
		procedureStatements: (ctx, push) => checkDivisionByZeroExpressions(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectIntegerConstants,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
			ctx.opts.hostModel,
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
		procedureStatements: (ctx, push) => checkRedimImpossibleBounds(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectIntegerConstants,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
			ctx.opts.hostModel,
		),
	},
	{
		name: 'arrayDeclarationImpossibleBounds',
		run: (ctx, push) => checkArrayDeclarationBounds(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectIntegerConstants,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
			ctx.opts.hostModel,
		),
	},
	{
		name: 'redimPreserveDimensions',
		run: (ctx, push) => checkRedimPreserveDimensions(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'unallocatedDynamicArrayAccess',
		run: (ctx, push) => checkUnallocatedDynamicArrayAccess(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'collectionLoopCounters',
		run: (ctx, push) => checkCollectionLoopCounters(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'arraySubscriptOutOfBounds',
		run: (ctx, push) => checkFixedArraySubscriptBounds(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push, ctx.opts.projectIntegerConstants, ctx.opts.projectVisibleSymbols, ctx.opts.hostModel),
	},
	{
		name: 'declareStatements',
		run: (ctx, push) => checkDeclareStatements(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'unusableDeclareCalls',
		run: (ctx, push) => checkUnusableDeclareCalls(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'typeFieldArrays',
		run: (ctx, push) => checkTypeFieldArrays(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push, ctx.opts.projectIntegerConstants, ctx.opts.projectVisibleSymbols, ctx.opts.hostModel),
	},
	{
		name: 'typeMembers',
		run: (ctx, push) => checkTypeMembers(ctx.source, ctx.mod, ctx.symbols, ctx.memberCtx, ctx.activity, push),
	},
	{
		name: 'midStatementLiteralTarget',
		run: (ctx, push) =>
			checkMidStatementLiteralTarget(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
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
		name: 'meOutsideObjectModule',
		procedureStatements: (ctx, push) => checkMeOutsideObjectModule(ctx.moduleKind, ctx.source, push),
	},
	{
		name: 'withEventsDeclarations',
		run: (ctx, push) => checkWithEventsDeclarations(ctx.source, ctx.mod, ctx.moduleKind, ctx.activity, push, ctx.memberCtx.projectClassMembers),
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
		name: 'invalidPropertyUse',
		procedureStatements: (ctx, push) => checkInvalidPropertyUse(ctx.source, ctx.memberCtx, push),
	},
	{
		name: 'raiseEventArguments',
		run: (ctx, push) => checkRaiseEventArguments(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'longLongNarrowing',
		run: (ctx, push) => checkLongLongNarrowing(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.conditionalCompilation,
			ctx.opts.host,
			ctx.opts.projectProcedures,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
		),
	},
	{
		name: 'addressOfUse',
		run: (ctx, push) => checkAddressOfUse(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectClassMembers,
			ctx.opts.conditionalCompilation,
			ctx.activity,
			push,
		),
	},
	{
		name: 'declarePtrSafeForWin64',
		run: (ctx, push) => checkDeclarePtrSafeForWin64(
			ctx.source,
			ctx.mod,
			ctx.opts.conditionalCompilation,
			ctx.opts.host,
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
			ctx.activity,
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
		name: 'declarationOrder',
		run: (ctx, push) => checkDeclarationOrder(
			ctx.source,
			ctx.mod,
			ctx.opts.moduleName,
			ctx.opts.projectIntegerConstants,
			ctx.opts.projectClassMembers,
			ctx.activity,
			push,
		),
	},
	{
		name: 'lineNumberRange',
		run: (ctx, push) => checkLineNumberRange(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'undefinedLabels',
		run: (ctx, push) => checkUndefinedLabels(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'reservedLabels',
		run: (ctx, push) => checkReservedLabels(ctx.source, ctx.mod, ctx.activity, push),
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
		name: 'duplicateCaseElse',
		run: (ctx, push) => checkDuplicateCaseElse(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'malformedStatements',
		run: (ctx, push) => checkMalformedStatements(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'elseWithoutIf',
		run: (ctx, push) => checkElseWithoutIf(ctx.source, ctx.mod, ctx.activity, push),
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
			ctx.memberCtx,
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
		name: 'missingScriptingReference',
		run: (ctx, push) => checkMissingScriptingReference(
			ctx.source, ctx.opts.referencedLibraries,
			new Set([...(ctx.memberCtx.projectClassMembers ?? []).map((type) => type.name), ...(ctx.symbols.root.children ?? []).map((symbol) => symbol.name)].map((name) => name.toLowerCase())),
			push,
		),
	},
	{
		name: 'missingLibraryReference',
		// The resolved model, not opts.hostModel: a bare Excel project
		// passes undefined so the downstream default rides, and the rule
		// would then know of no library at all and stay silent.
		run: (ctx, push) => checkMissingLibraryReference(
			ctx.source, ctx.opts.hostModel ?? getExcelObjectModel(), push,
			new Set([ctx.opts.moduleName ?? '', ...(ctx.memberCtx.projectClassMembers ?? []).map((type) => type.name)].map((name) => name.toLowerCase())),
		),
	},
	{
		name: 'memberNotFound',
		procedureStatements: (ctx, push) => checkMemberNotFound(ctx.source, ctx.memberCtx, push),
	},
	{
		name: 'moduleMemberForms',
		procedureStatements: (ctx, push) => checkModuleMemberForms(
			ctx.source,
			ctx.symbols,
			ctx.memberCtx,
			ctx.opts.projectVisibleSymbols,
			/^[ \t]*Option[ \t]+Explicit\b/im.test(ctx.source),
			push,
		),
	},
	{
		name: 'invalidParamArrayUse',
		procedureStatements: (ctx, push) => checkParamArrayUse(ctx.source, callableTypeSignaturesFor(ctx.symbols, ctx.opts.projectProcedures), push),
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
		name: 'omittedArgumentReads',
		procedureStatements: (ctx, push) => checkOmittedArgumentReads(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.activity,
			ctx.opts.projectVisibleSymbols,
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
			ctx.activity,
		),
	},
	{
		name: 'runtimeArgumentValues',
		blockHeaders: true,
		procedureStatements: (ctx, push) => checkRuntimeArgumentValues(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectIntegerConstants,
			ctx.opts.projectVisibleSymbols,
			ctx.activity,
			push,
			ctx.opts.hostModel,
		),
	},
	{
		name: 'runtimeConversionValues',
		blockHeaders: true,
		procedureStatements: (ctx, push) => checkRuntimeConversionValues(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectVisibleSymbols,
			push,
			ctx.activity,
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
		procedureExpressions: (ctx, push) => checkTypeOfIsCompatibility(
			ctx.symbols,
			ctx.memberCtx,
			push,
		),
	},
	{
		name: 'typeofMissingOperand',
		run: (ctx, push) => checkTypeOfMissingOperand(ctx.source, ctx.activity, push),
	},
	{
		name: 'isOperatorNonObject',
		procedureExpressions: (ctx, push) => checkIsOperatorOperands(ctx.symbols, push),
	},
	{
		name: 'isOperandsInConditions',
		run: (ctx, push) => checkIsOperandsInConditions(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		name: 'nonScalarBinaryOperand',
		procedureExpressions: (ctx, push) => checkBinaryOperandScalar(ctx.symbols, push),
	},
	{
		name: 'argumentShapeMismatch',
		procedureStatements: (ctx, push) => checkArgumentShape(
			ctx.source,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.opts.projectVisibleSymbols,
			ctx.memberCtx,
			push,
			ctx.mod,
			ctx.activity,
		),
	},
	{
		name: 'suffixedLiteralOverflow',
		run: (ctx, push) => checkSuffixedLiteralOverflow(ctx.source, ctx.activity, push),
	},
	{
		name: 'missingReturnAssignments',
		run: (ctx, push) => checkMissingReturnAssignments(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.opts.projectProcedures,
			ctx.activity,
			ctx.opts.moduleName,
			ctx.opts.implementedInterfaces,
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
				ctx.opts.hostModel,
				ctx.opts.designerClass,
				push,
				ctx.opts.projectClassMembers,
				ownObjectMemberNames(ctx.opts),
			);
		},
	},
	{
		// Cross-module rule: needs the project's class-member surfaces to know
		// which member names are Friend-only (see AnalyzeModuleOptions).
		name: 'lateBoundFriendMember',
		procedureStatements: (ctx, push) => {
			const projectClassMembers = ctx.opts.projectClassMembers;
			if (!projectClassMembers) {
				return () => undefined;
			}
			return checkLateBoundFriendMember(
				ctx.source,
				ctx.symbols,
				ctx.opts.projectVisibleSymbols,
				projectClassMembers,
				ctx.opts.hostModel,
				push,
			);
		},
	},
	{
		name: 'unusedDeclarations',
		run: (ctx, push) => checkUnusedDeclarations(ctx.source, ctx.mod, ctx.symbols, ctx.activity, push),
	},
	{
		// A Private procedure is reachable from its own module only, so the
		// module's text decides; the project's string literals are consulted
		// for a name reached through Application.Run, OnTime and their kin.
		name: 'unusedPrivateProcedures',
		run: (ctx, push) => checkUnusedPrivateProcedures(
			ctx.source,
			ctx.mod,
			ctx.symbols,
			ctx.moduleKind,
			ctx.activity,
			ctx.opts.projectStringLiteralWords,
			push,
		),
	},
	{
		name: 'unreachableCode',
		run: (ctx, push) => checkUnreachableCode(ctx.source, ctx.mod, ctx.activity, push),
	},
	{
		name: 'docComments',
		run: (ctx, push) => checkDocComments(ctx.source, ctx.mod, ctx.activity, push),
	},
];
