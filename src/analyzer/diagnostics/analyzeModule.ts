// Active VBA diagnostics engine (MS-VBAL Phase 5).
//
// `analyzeModule` runs the high-confidence semantic rules from the rule
// catalogue over one module's source and returns offset-based diagnostics. It
// is pure (no `vscode`): the editor layer converts spans to ranges and severity
// names to the VS Code enum. Structural block-balance checking stays in
// `src/vbaStructuralDiagnostics.ts` (analyzeVbaStructure); this engine adds the semantic rules on
// top, so the two do not overlap or double-report.
//
// Design rule (see /memories): no "looks like" heuristics. Every rule here is
// deterministic - it flags a construct only when the language guarantees it is
// an error. The one cross-module rule, `unknown-call`, fires only on a call
// statement whose callee is a bare (non-member) identifier - the lone-identifier
// form, the parenless-argument form (`MsgBox "hi"`), or `Call name` - whose name
// resolves to no procedure anywhere in the project, no VBA runtime
// function/statement, no host global or Application member, and no in-scope
// declaration - the unambiguous VBE "Sub or Function not defined" error.
// The `undeclared-variable` rule follows the same conservative pattern: it runs
// only when the caller has supplied project-visible names, and it scans
// statement-level reads while deliberately skipping unresolved external-style
// calls and type-name positions to avoid false positives.

import { tokenizeCached } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import {
	absoluteSpan,
	physicalLineSpanAtOffset,
	statementTokens,
} from './walker';
import type {
	ModuleNode,
	Span,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import { createConditionalActivityTracker } from '../conditional/conditionalCompilation';
import type { ModuleSymbolKind } from '../symbols/symbolModel';
import { type MemberCompletionContext } from '../completion/memberAccess';
import {
	DIAGNOSTIC_RULES,
	DiagnosticRuleName,
	DiagnosticSeverity,
	normalizeDiagnosticSeverityOverride,
} from './ruleMetadata';

export type {
	AnalyzeModuleOptions,
	DiagnosticSeverityOverrides,
	VbaCreateProcedureStubData,
	VbaDiagnostic,
	VbaDiagnosticData,
	VbaMissingRequiredArgumentPlaceholderData,
} from './analysisContext';
import { isObjectModuleKind } from './analysisContext';
import type {
	AnalyzeModuleOptions,
	DiagnosticSeverityOverrides,
	VbaDiagnostic,
	VbaDiagnosticData,
} from './analysisContext';
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
} from './rules/duplicates';
import {
	checkDimInitializer,
	checkFixedLengthStringBounds,
	checkInvalidAsTypeNames,
	checkInvalidIdentifierStarts,
	checkModuleDeclarationsAfterProcedures,
	checkModuleDeclarationsInProcedureBodies,
	checkModuleLevelStatementsOutsideProcedures,
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
	incompleteMemberAccess,
	isNonUnaryBinaryOperator,
} from './rules/expressions';
import {
	checkDuplicateLabels,
	checkElseBranchOrder,
	checkExitStatements,
	checkForEachLoopTypes,
	checkStatementContext,
	checkUndefinedLabels,
} from './rules/controlFlow';

/** Resolves the effective severity of a rule, or undefined when switched off. */
function severityOf(
	rule: DiagnosticRuleName,
	overrides: DiagnosticSeverityOverrides | undefined,
): DiagnosticSeverity | undefined {
	const meta = DIAGNOSTIC_RULES[rule];
	const override = normalizeDiagnosticSeverityOverride(meta.code, overrides?.[meta.code]);
	if (override === 'off') {
		return undefined;
	}
	return override ?? meta.defaultSeverity;
}

/**
 * Analyzes one VBA module source and returns its active diagnostics.
 * Never throws: any internal failure yields an empty list.
 */
export function analyzeModule(
	source: string,
	opts: AnalyzeModuleOptions = {},
): VbaDiagnostic[] {
	try {
		return runRules(source, opts);
	} catch {
		return [];
	}
}

export function incompleteExpressionEditSpan(
	source: string,
	offset: number,
): Span | undefined {
	const line = physicalLineSpanAtOffset(source, offset);
	const statement = activeStatementSpanOnLine(source, line, offset);
	return incompleteMemberAccess(source, statement, { includeLeadingDot: true })?.span
		?? trailingBinaryOperatorEditSpan(source, statement, offset)
		?? unmatchedOpenParenEditSpan(source, statement, offset);
}

function runRules(
	source: string,
	opts: AnalyzeModuleOptions,
): VbaDiagnostic[] {
	const out: VbaDiagnostic[] = [];
	const moduleName = opts.moduleName ?? 'Module';
	const moduleKind = opts.moduleKind ?? 'standard';
	const overrides = opts.severityOverrides;

	const push = (
		rule: DiagnosticRuleName,
		message: string,
		span: Span,
		data?: VbaDiagnosticData,
	): void => {
		const severity = severityOf(rule, overrides);
		if (!severity) {
			return;
		}
		const meta = DIAGNOSTIC_RULES[rule];
		out.push({
			code: meta.code,
			message,
			severity,
			span,
			specReference: meta.specReference,
			...(data ? { data } : {}),
		});
	};

	const mod = opts.parsedModule ?? parseModule(source);
	const symbols = buildModuleSymbols(moduleName, moduleKind, source, {
		conditionalCompilation: opts.conditionalCompilation,
		parsedModule: mod,
	});
	const activity = createConditionalActivityTracker(mod, opts.conditionalCompilation);
	const memberCtx = diagnosticMemberCompletionContext(opts, source, mod);

	checkUnterminatedStrings(source, push);
	checkInvalidLineContinuations(source, push);
	checkDuplicateProcedures(symbols.root.children ?? [], push);
	checkDuplicateDeclarations(symbols.root.children ?? [], push);
	checkDuplicateModuleMembers(symbols.root.children ?? [], push);
	checkDuplicateEnumMembers(source, mod, activity, push);
	checkAmbiguousEnumMemberReferences(
		source,
		mod,
		symbols,
		activity,
		moduleName,
		opts.knownProcedures,
		opts.projectProcedures,
		opts.projectClassMembers,
		opts.projectVisibleSymbols,
		push,
	);
	checkConstAssignment(source, mod, symbols, activity, opts.projectVisibleSymbols, push);
	checkOptionExplicit(source, mod, activity, push);
	checkUndeclaredVariables(
		source,
		mod,
		symbols,
		activity,
		opts.knownIdentifiers,
		opts.projectProcedures,
		opts.projectClassMembers,
		opts.projectVisibleSymbols,
		push,
	);
	checkOptionPlacement(source, mod, activity, push);
	checkProcedureHeader(source, mod, activity, push);
	checkInvalidIdentifierStarts(source, mod, activity, push);
	checkModuleDeclarationsInProcedureBodies(source, mod, activity, push);
	checkModuleDeclarationsAfterProcedures(source, mod, activity, push);
	checkModuleLevelStatementsOutsideProcedures(source, mod, activity, push);
	checkReservedDeclarationNames(source, mod, activity, push);
	checkPropertySetterValueParameters(source, mod, activity, opts, push);
	checkPropertyAccessorSignatures(source, mod, activity, push);
	checkParameterOrder(source, mod, activity, push);
	checkParameterDefaultValues(source, mod, activity, memberCtx, push);
	checkUnbalancedParens(source, push);
	checkInvalidExpressionSyntax(source, mod, symbols, opts.projectVisibleSymbols, activity, push);
	checkDivisionByZeroExpressions(
		source,
		mod,
		symbols,
		opts.projectIntegerConstants,
		opts.projectVisibleSymbols,
		activity,
		push,
	);
	checkDimInitializer(source, mod, activity, push);
	checkInvalidRedimTargets(source, mod, symbols, opts.projectVisibleSymbols, activity, push);
	checkRedimImpossibleBounds(source, mod, activity, push);
	checkRedimPreserveDimensions(source, mod, activity, push);
	checkUnallocatedDynamicArrayAccess(source, mod, activity, push);
	checkEraseTargets(source, mod, symbols, opts.projectVisibleSymbols, activity, push);
	checkTypeDeclarationCharacterAsClause(mod, activity, push);
	checkUnexpectedDeclarationTokens(source, mod, activity, push);
	checkFixedLengthStringBounds(source, mod, activity, push);
	checkObjectModulePublicMembers(source, mod, moduleKind, activity, push);
	checkEventDeclarationModuleKind(source, mod, moduleKind, activity, push);
	checkWithEventsDeclarations(source, mod, moduleKind, activity, push);
	checkFriendDeclarations(source, mod, moduleKind, activity, push);
	checkImplementsStatementPlacement(source, mod, moduleKind, activity, push);
	checkRaiseEventTargets(source, mod, activity, push);
	checkDeclarePtrSafeForWin64(source, mod, opts.conditionalCompilation, activity, push);
	checkEventHandlerModuleScope(source, mod, moduleName, moduleKind, opts.documentType, activity, push);
	checkInvalidAsTypeNames(source, mod, activity, opts, push);
	checkCallParens(source, mod, symbols, opts.projectProcedures, opts.projectVisibleSymbols, memberCtx, activity, push);
	checkExpressionCallParens(source, mod, symbols, opts.projectProcedures, opts.projectVisibleSymbols, activity, push);
	checkSetAssignments(source, mod, symbols, opts.projectVisibleSymbols, memberCtx, activity, push);
	checkExitStatements(source, mod, activity, push);
	checkDuplicateLabels(source, mod, activity, push);
	checkUndefinedLabels(source, mod, activity, push);
	checkElseBranchOrder(source, mod, activity, push);
	checkStatementContext(source, mod, activity, push);
	checkForEachLoopTypes(mod, symbols, opts, activity, push);
	checkArrayBoundIntrinsicArguments(source, mod, symbols, opts.projectVisibleSymbols, activity, push);
	checkScalarMemberAccess(source, mod, symbols, opts.projectVisibleSymbols, activity, push);
	checkObjectVariableNotSet(source, mod, symbols, memberCtx, activity, push);
	checkMemberNotFound(source, mod, memberCtx, activity, push);
	checkNonCallableCallStatement(
		source,
		mod,
		symbols,
		activity,
		opts.knownProcedures,
		opts.projectVisibleSymbols,
		push,
	);
	checkArgumentCount(
		source,
		mod,
		symbols,
		opts.projectProcedures,
		opts.projectVisibleSymbols,
		memberCtx,
		activity,
		push,
	);
	checkArgumentTypes(
		source,
		mod,
		symbols,
		opts.projectProcedures,
		opts.projectVisibleSymbols,
		memberCtx,
		activity,
		push,
	);
	checkRuntimeArgumentValues(
		source,
		mod,
		symbols,
		opts.projectProcedures,
		opts.projectIntegerConstants,
		opts.projectVisibleSymbols,
		activity,
		push,
	);
	checkRuntimeConversionValues(source, mod, symbols, opts.projectVisibleSymbols, activity, push);
	checkAssignmentTypes(source, mod, symbols, opts.projectVisibleSymbols, memberCtx, activity, push);
	checkMissingReturnAssignments(source, mod, symbols, opts.projectProcedures, activity, push);
	if (opts.knownProcedures) {
		checkUnknownCallStatement(
			source,
			mod,
			symbols,
			activity,
			opts.knownProcedures,
			opts.projectVisibleSymbols,
			push,
		);
	}

	return out;
}

function diagnosticMemberCompletionContext(
	opts: AnalyzeModuleOptions,
	source: string,
	mod: ModuleNode,
): MemberCompletionContext {
	const ctx: MemberCompletionContext = {
		projectClassMembers: opts.projectClassMembers,
		allowSetAssignmentRefinement: false,
		model: opts.hostModel,
		// Audit #1: member resolution used to re-parse the module and re-lex the
		// source prefix once per dotted reference. Hand it the per-pass AST and
		// the shared full-source token stream instead.
		parsedModule: mod,
		sourceTokens: tokenizeCached(source).filter((t) => t.kind !== 'comment'),
	};
	const meProjectType = meProjectTypeFor(opts.moduleName, opts.moduleKind);
	if (meProjectType) {
		ctx.meProjectType = meProjectType;
	}
	const meType = meHostTypeFor(opts.moduleName, opts.moduleKind);
	if (meType) {
		ctx.meType = meType;
	}
	return ctx;
}

function meProjectTypeFor(
	moduleName: string | undefined,
	moduleKind: ModuleSymbolKind | undefined,
): string | undefined {
	return moduleName && isObjectModuleKind(moduleKind) ? moduleName : undefined;
}

function meHostTypeFor(
	moduleName: string | undefined,
	moduleKind: ModuleSymbolKind | undefined,
): string | undefined {
	if (!moduleName || moduleKind !== 'document') {
		return undefined;
	}
	return moduleName.toLowerCase() === 'thisworkbook' ? 'Excel.Workbook' : undefined;
}

function activeStatementSpanOnLine(source: string, line: Span, offset: number): Span {
	const safeOffset = Math.max(0, Math.min(offset, source.length));
	const toks = statementTokens(source, line);
	let depth = 0;
	let start = line.start;
	let end = line.end;
	for (const tok of toks) {
		if (tok.kind === 'punctuation') {
			if (tok.rawText === '(') {
				depth++;
			} else if (tok.rawText === ')') {
				depth = Math.max(0, depth - 1);
			}
			continue;
		}
		if (tok.kind !== 'colon' || depth !== 0) {
			continue;
		}
		const colon = absoluteSpan(line, tok);
		if (safeOffset <= colon.start) {
			end = colon.start;
			break;
		}
		start = colon.end;
	}
	return { start, end };
}

function trailingBinaryOperatorEditSpan(
	source: string,
	span: Span,
	offset: number,
): Span | undefined {
	const toks = statementTokens(source, span);
	const last = toks[toks.length - 1];
	if (!isNonUnaryBinaryOperator(last)) {
		return undefined;
	}
	const active = absoluteSpan(span, last);
	if (offset < active.start) {
		return undefined;
	}
	const cursorTail = source.slice(active.end, Math.min(offset, span.end));
	return /^[ \t]*$/.test(cursorTail) ? active : undefined;
}

function unmatchedOpenParenEditSpan(
	source: string,
	span: Span,
	offset: number,
): Span | undefined {
	const stack: VbaToken[] = [];
	for (const tok of statementTokens(source, span)) {
		if (tok.kind !== 'punctuation') {
			continue;
		}
		if (tok.rawText === '(') {
			stack.push(tok);
		} else if (tok.rawText === ')' && stack.length > 0) {
			stack.pop();
		}
	}
	const firstUnmatched = stack[0];
	if (!firstUnmatched) {
		return undefined;
	}
	const active = absoluteSpan(span, firstUnmatched);
	return offset >= active.start ? active : undefined;
}
