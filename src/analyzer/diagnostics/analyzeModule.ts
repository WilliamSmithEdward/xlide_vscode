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

import { tokenize, tokenizeCached } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import {
	parseVbaIntegerLiteral,
	resolveRawIntegerConstants,
	type IntegerConstantLookup,
} from '../constants/integerConstantExpression';
import { walkStraightLineBody } from './dataflow';
import {
	absoluteSpan,
	activeModuleMembers,
	bareAssignmentTarget,
	declaredNameSpan,
	firstTokenSpan,
	forEachProcedureBodyLine,
	forEachStatement,
	forEachVariableGroup,
	isInactiveNode,
	localsPassedAsCallArguments,
	matchParenFrom,
	physicalLineSpanAtOffset,
	pluralizeCount,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
	topLevelOperatorIndex,
} from './walker';
import {
	callableAcceptsZeroArguments,
	splitArgSlots,
	type CallableTypeSignature,
} from './callExtraction';
import {
	collectBodyLiteralIntegerConstants,
	collectModuleLiteralIntegerConstants,
	foldIntegerExpressionTokens,
} from './constExpr';
import {
	bareCallableSourceShadowed,
	callableSignatureFor,
	callableTypeSignaturesFor,
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	declaredTypeForSourceBinding,
	isKnownScalarType,
	normalizeType,
	resolveExactMemberCompletion,
	runtimeCallableSourceShadowed,
	scopedIntegerConstantLookup,
	sourceNameScopeFor,
	typeEnvironmentFor,
	type DeclaredValueShape,
	type SourceDeclaredShape,
	type SourceDeclaredTypeResolver,
	type SourceNameScope,
} from './typeInference';
import {
	resolveRuntimeFunction,
	runtimeAllowsExplicitCall,
} from '../runtime/vbaRuntime';
import {
	explicitCallStatementArgumentWithoutParens,
	explicitCallStatementTarget,
	standaloneEmptyParenthesizedCallStatement,
} from '../call/callContext';
import type {
	BodyNode,
	ForBlockNode,
	ModuleNode,
	ProcedureNode,
	Span,
	StatementNode,
	VariableDeclNode,
	VariableGroupNode,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import {
	conditionalCompilerConstants,
	createConditionalActivityTracker,
	type ConditionalActivityTracker,
	type ConditionalCompilationEnvironment,
} from '../conditional/conditionalCompilation';
import {
	collectProcedureLabelDeclarations,
	collectProcedureLabelReferences,
	collectProcedureLabels,
} from '../flow/procedureLabels';
import type {
	ModuleSymbolKind,
	VbaProcedureSignature,
	VbaSymbol,
} from '../symbols/symbolModel';
import { qualifiedProcedureKey } from '../symbols/symbolModel';
import { type BareIdentifierContext } from '../symbols/nameResolution';
import { type MemberCompletionContext } from '../completion/memberAccess';
import { resolveTypeName } from '../completion/typeCompletion';
import {
	eventHandlerDocumentTypeForContext,
	eventHandlerProcedureForName,
	type EventHandlerDocumentType,
} from '../completion/eventHandlers';
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
import {
	isObjectModuleKind,
	procedureSymbolFor,
} from './analysisContext';
import type {
	AnalyzeModuleOptions,
	DiagnosticSeverityOverrides,
	PushFn,
	VbaDiagnostic,
	VbaDiagnosticData,
} from './analysisContext';
import {
	isBareOrVbaQualifiedIntrinsicCall,
	scanConditionalCompilationBranchOrder,
} from './rules/shared';
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

/**
 * Rule: every parenthesis must be matched within its logical statement. VBA has
 * no cross-statement parentheses (a `(` is closed before the line ends unless a
 * `_` line-continuation joins the next physical line, which the lexer already
 * folds into trivia), so an open `(` left dangling at a statement boundary, or a
 * `)` with no matching `(`, is always the VBE "Expected: )" / "Syntax error".
 *
 * The scan walks the whole module's token stream, tracking paren depth and
 * resetting at each logical-statement boundary (a `newline` token or a depth-0
 * `:` statement separator). Only literal `(`/`)` punctuation tokens count -
 * parentheses inside strings, comments, date literals, and `[bracketed]` names
 * are distinct token kinds, so they can never create a false positive. At most
 * one diagnostic is reported per statement.
 */
function checkUnbalancedParens(source: string, push: PushFn): void {
	const toks = tokenizeCached(source);
	let depth = 0;
	const openOffsets: number[] = [];
	let flagged = false;

	const flush = (): void => {
		if (!flagged && depth > 0) {
			const off = openOffsets[0];
			push(
				'unbalancedParens',
				"Unbalanced parentheses: a ')' is missing.",
				{ start: off, end: off + 1 },
			);
		}
		depth = 0;
		openOffsets.length = 0;
		flagged = false;
	};

	for (const tok of toks) {
		if (tok.kind === 'newline') {
			flush();
			continue;
		}
		if (tok.kind === 'colon' && depth === 0) {
			flush();
			continue;
		}
		if (tok.kind !== 'punctuation') {
			continue;
		}
		if (tok.rawText === '(') {
			depth++;
			openOffsets.push(tok.start);
		} else if (tok.rawText === ')') {
			if (depth === 0) {
				if (!flagged) {
					push(
						'unbalancedParens',
						"Unbalanced parentheses: an unexpected ')' was found.",
						{ start: tok.start, end: tok.end },
					);
					flagged = true;
				}
			} else {
				depth--;
				openOffsets.pop();
			}
		}
	}
	flush();
}

function checkArrayBoundIntrinsicArguments(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			for (const hit of arrayBoundScalarArguments(
				source,
				stmt.span,
				shapes,
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'expression',
				),
			)) {
				push(
					'arrayBoundRequiresArray',
					`${hit.functionName} requires an array argument, but '${hit.name}' is declared As ${hit.asType}.`,
					hit.span,
				);
			}
		}, activity);
	}
}

interface ArrayBoundScalarArgument {
	functionName: string;
	name: string;
	span: Span;
	asType: string;
}

function arrayBoundScalarArguments(
	source: string,
	span: Span,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveShape?: (name: string) => SourceDeclaredShape,
): ArrayBoundScalarArgument[] {
	const toks = statementTokens(source, span);
	const hits: ArrayBoundScalarArgument[] = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const functionName = tokenName(toks[i]);
		const lower = functionName?.toLowerCase();
		if (lower !== 'lbound' && lower !== 'ubound') {
			continue;
		}
		if (toks[i + 1]?.rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const inner = toks.slice(i + 2, close);
		if (inner.length === 0) {
			continue;
		}
		const split = splitArgSlots(inner, span.start);
		const firstSlot = split.slots[0] ?? [];
		if (firstSlot.length !== 1) {
			continue;
		}
		const argName = tokenName(firstSlot[0]);
		if (!argName) {
			continue;
		}
		const resolvedShape = resolveShape?.(argName);
		const shape = resolvedShape?.resolved
			? resolvedShape.shape
			: shapes.get(argName.toLowerCase());
		if (!shape || shape.isArray || !shape.asType) {
			continue;
		}
		const normalized = normalizeType(shape.asType);
		if (!normalized || !isKnownScalarType(normalized)) {
			continue;
		}
		hits.push({
			functionName: functionName!,
			name: argName,
			span: split.spans[0] ?? { start: span.start + firstSlot[0].start, end: span.start + firstSlot[0].end },
			asType: shape.asType,
		});
	}
	return hits;
}

interface RedimBlockedDeclaration {
	name: string;
	span: Span;
	kind: 'fixedArray' | 'scalar';
}

interface RedimDimension {
	key?: string;
	lowerKey?: string;
	lowerValue?: number;
	upperValue?: number;
	span: Span;
}

interface RedimTarget {
	name: string;
	span: Span;
	preserve: boolean;
	dimensions: RedimDimension[];
}

/**
 * Rule: ReDim can allocate dynamic arrays, but it cannot resize a variable that
 * was already declared as a scalar or as a fixed-size array.
 */
function checkInvalidRedimTargets(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleDeclarations = redimBlockedDeclarationsForModule(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const localDeclarations = redimBlockedDeclarationsForBody(member.body, activity);
		const localNames = declarationNamesForBody(member.body, activity);
		const procSym = procedureSymbolFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			for (const target of redimTargets(source, stmt.span)) {
				const lower = target.name.toLowerCase();
				const resolvedShape = declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					target.name,
					'assignmentTarget',
				);
				const declaration = resolvedShape.resolved
					? redimBlockedDeclarationForShape(target.name, target.span, resolvedShape.shape)
					: localDeclarations.get(lower) ??
						(localNames.has(lower) ? undefined : moduleDeclarations.get(lower));
				if (!declaration) {
					continue;
				}
				if (declaration.kind === 'scalar') {
					push(
						'scalarRedim',
						`Scalar variable '${target.name}' cannot be resized with ReDim; declare it as a dynamic array first.`,
						target.span,
					);
					continue;
				}
				push(
					'fixedArrayRedim',
					`Fixed-size array '${target.name}' cannot be resized with ReDim.`,
					target.span,
				);
			}
		}, activity);
	}
}

function redimBlockedDeclarationForShape(
	name: string,
	span: Span,
	shape: DeclaredValueShape | undefined,
): RedimBlockedDeclaration | undefined {
	if (!shape) {
		return undefined;
	}
	if (!shape.isArray) {
		if (isVariantLikeRedimTargetType(shape.asType)) {
			return undefined;
		}
		return { name, span, kind: 'scalar' };
	}
	if (shape.isFixedArray) {
		return { name, span, kind: 'fixedArray' };
	}
	return undefined;
}

function isVariantLikeRedimTargetType(asType: string | undefined): boolean {
	return !asType || normalizeType(asType) === 'variant';
}

function redimBlockedDeclarationsForModule(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): Map<string, RedimBlockedDeclaration> {
	const declarations = new Map<string, RedimBlockedDeclaration>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			addRedimBlockedDeclarations(member, declarations);
		}
	}
	return declarations;
}

function redimBlockedDeclarationsForBody(
	body: BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Map<string, RedimBlockedDeclaration> {
	const declarations = new Map<string, RedimBlockedDeclaration>();
	forEachVariableGroup(body, (group) => addRedimBlockedDeclarations(group, declarations), activity);
	return declarations;
}

function declarationNamesForBody(
	body: BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Set<string> {
	const names = new Set<string>();
	forEachVariableGroup(body, (group) => {
		for (const decl of group.declarations) {
			names.add(decl.name.toLowerCase());
		}
	}, activity);
	return names;
}

function addRedimBlockedDeclarations(
	group: VariableGroupNode,
	out: Map<string, RedimBlockedDeclaration>,
): void {
	for (const decl of group.declarations) {
		const kind = redimBlockedDeclarationKind(decl);
		if (!kind) {
			continue;
		}
		const lower = decl.name.toLowerCase();
		if (!out.has(lower)) {
			out.set(lower, { name: decl.name, span: decl.span, kind });
		}
	}
}

function redimBlockedDeclarationKind(decl: VariableDeclNode): RedimBlockedDeclaration['kind'] | undefined {
	if (!decl.isArray) {
		if (isVariantLikeRedimTargetType(decl.asType)) {
			return undefined;
		}
		return 'scalar';
	}
	return decl.arrayBounds ? 'fixedArray' : undefined;
}

function redimTargets(source: string, span: Span): Array<{ name: string; span: Span }> {
	return redimStatementTargets(source, span).map((target) => ({
		name: target.name,
		span: target.span,
	}));
}

function redimStatementTargets(source: string, span: Span): RedimTarget[] {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'redim') {
		return [];
	}
	const preserve = tokenText(toks[1]) === 'preserve';
	const start = preserve ? 2 : 1;
	const out: RedimTarget[] = [];
	for (const group of splitTopLevelTokenGroups(toks.slice(start), ',')) {
		const target = redimTargetFromGroup(span, group, preserve);
		if (target) {
			out.push(target);
		}
	}
	return out;
}

function redimTargetFromGroup(
	base: Span,
	group: readonly VbaToken[],
	preserve: boolean,
): RedimTarget | undefined {
	const content = group.filter((tok) => tok.kind !== 'comment');
	const nameTok = content[0];
	const name = tokenName(nameTok);
	if (!name || !nameTok) {
		return undefined;
	}
	const dimensions: RedimDimension[] = [];
	if (content[1]?.rawText === '(') {
		const close = matchParenFrom(content, 1);
		if (close > 1) {
			for (const part of splitTopLevelTokenGroups(content.slice(2, close), ',')) {
				const dimTokens = part.filter((tok) => tok.kind !== 'comment');
				if (dimTokens.length === 0) {
					continue;
				}
				const bound = comparableArrayBoundKey(dimTokens);
				dimensions.push({
					key: bound.key,
					lowerKey: bound.lowerKey,
					lowerValue: bound.lowerValue,
					upperValue: bound.upperValue,
					span: tokenGroupSpan(base, dimTokens),
				});
			}
		}
	}
	return {
		name,
		span: absoluteSpan(base, nameTok),
		preserve,
		dimensions,
	};
}

function comparableArrayBoundKey(
	toks: readonly VbaToken[],
): { key?: string; lowerKey?: string; lowerValue?: number; upperValue?: number } {
	const toIndex = toks.findIndex((tok) => tokenText(tok) === 'to');
	if (toIndex > 0) {
		const lower = comparableArrayBoundExpression(toks.slice(0, toIndex));
		const upper = comparableArrayBoundExpression(toks.slice(toIndex + 1));
		return {
			key: lower.key && upper.key ? `${lower.key}to${upper.key}` : undefined,
			lowerKey: lower.key,
			lowerValue: lower.value,
			upperValue: upper.value,
		};
	}
	const upper = comparableArrayBoundExpression(toks);
	return { key: upper.key, upperValue: upper.value };
}

function comparableArrayBoundExpression(toks: readonly VbaToken[]): { key?: string; value?: number } {
	return {
		key: comparableArrayBoundExpressionKey(toks),
		value: comparableArrayBoundExpressionValue(toks),
	};
}

function comparableArrayBoundExpressionKey(toks: readonly VbaToken[]): string | undefined {
	const parts: string[] = [];
	for (const tok of toks) {
		const word = tokenText(tok);
		if (
			tok.kind === 'integerLiteral' ||
			tok.rawText === '+' ||
			tok.rawText === '-' ||
			word === 'to'
		) {
			parts.push(word || tok.rawText.toLowerCase());
			continue;
		}
		return undefined;
	}
	return parts.length > 0 ? parts.join('') : undefined;
}

function comparableArrayBoundExpressionValue(toks: readonly VbaToken[]): number | undefined {
	let value = 0;
	let sign = 1;
	let expectingValue = true;
	let sawValue = false;
	for (const tok of toks) {
		if (expectingValue) {
			if (tok.rawText === '+' || tok.rawText === '-') {
				sign *= tok.rawText === '-' ? -1 : 1;
				continue;
			}
			if (tok.kind !== 'integerLiteral') {
				return undefined;
			}
			const parsed = parseVbaIntegerLiteral(tok.rawText);
			if (parsed === undefined) {
				return undefined;
			}
			const next = value + sign * parsed;
			if (!Number.isSafeInteger(next)) {
				return undefined;
			}
			value = next;
			sign = 1;
			expectingValue = false;
			sawValue = true;
			continue;
		}
		if (tok.rawText === '+' || tok.rawText === '-') {
			sign = tok.rawText === '-' ? -1 : 1;
			expectingValue = true;
			continue;
		}
		return undefined;
	}
	return sawValue && !expectingValue ? value : undefined;
}

/**
 * Rule: ReDim lower bounds must not be greater than their upper bounds.
 * This only reports explicit, literal-style `lower To upper` dimensions.
 */
function checkRedimImpossibleBounds(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleDeclarations = redimBlockedDeclarationsForModule(mod, activity);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const localDeclarations = redimBlockedDeclarationsForBody(member.body, activity);
		const localNames = declarationNamesForBody(member.body, activity);
		forEachStatement(member.body, (stmt) => {
			for (const target of redimStatementTargets(source, stmt.span)) {
				const lowerName = target.name.toLowerCase();
				const blockedDeclaration = localDeclarations.get(lowerName) ??
					(localNames.has(lowerName) ? undefined : moduleDeclarations.get(lowerName));
				if (blockedDeclaration) {
					continue;
				}
				target.dimensions.forEach((dimension, index) => {
					if (
						dimension.lowerValue === undefined ||
						dimension.upperValue === undefined ||
						dimension.lowerValue <= dimension.upperValue
					) {
						return;
					}
					push(
						'redimImpossibleBounds',
						`ReDim lower bound ${dimension.lowerValue} is greater than upper bound ${dimension.upperValue} for dimension ${index + 1} of '${target.name}'; this will raise Run-time error '9': Subscript out of range.`,
						dimension.span,
					);
				});
			}
		}, activity);
	}
}

/**
 * Rule: ReDim Preserve may only resize the last dimension of an already
 * allocated dynamic array. This tracks simple, active ReDim shapes in a
 * conservative per-body flow so nested branch updates do not leak outward.
 */
function checkRedimPreserveDimensions(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		checkRedimPreserveDimensionsInBody(source, member.body, new Map(), activity, push);
	}
}

function checkRedimPreserveDimensionsInBody(
	source: string,
	body: BodyNode[],
	initialShapes: ReadonlyMap<string, RedimTarget>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const shapes = new Map(initialShapes);
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			for (const target of redimStatementTargets(source, node.span)) {
				if (target.preserve) {
					const previous = shapes.get(target.name.toLowerCase());
					const reason = previous
						? redimPreserveDimensionMismatch(previous, target)
						: undefined;
					if (reason) {
						push(
							'redimPreserveDimensionChange',
							`ReDim Preserve can only resize the last dimension of '${target.name}'. ${reason}`,
							target.span,
						);
					}
				}
				if (target.dimensions.length > 0) {
					shapes.set(target.name.toLowerCase(), target);
				}
			}
			continue;
		}
		if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			checkRedimPreserveDimensionsInBody(
				source,
				(node as { body: BodyNode[] }).body,
				shapes,
				activity,
				push,
			);
		}
	}
}

function redimPreserveDimensionMismatch(
	previous: RedimTarget,
	current: RedimTarget,
): string | undefined {
	if (
		previous.dimensions.length > 0 &&
		current.dimensions.length > 0 &&
		previous.dimensions.length !== current.dimensions.length
	) {
		return `Previous ReDim has ${pluralizeCount(previous.dimensions.length, 'dimension')}, but this ReDim Preserve has ${current.dimensions.length}.`;
	}
	const comparableCount = Math.min(previous.dimensions.length, current.dimensions.length) - 1;
	for (let i = 0; i < comparableCount; i++) {
		const before = previous.dimensions[i]?.key;
		const after = current.dimensions[i]?.key;
		if (before && after && before !== after) {
			return `Dimension ${i + 1} changes before the final dimension.`;
		}
	}
	const finalIndex = Math.min(previous.dimensions.length, current.dimensions.length) - 1;
	if (finalIndex >= 0) {
		const beforeLower = previous.dimensions[finalIndex]?.lowerKey;
		const afterLower = current.dimensions[finalIndex]?.lowerKey;
		if (beforeLower && afterLower && beforeLower !== afterLower) {
			return `The lower bound of dimension ${finalIndex + 1} changes under Preserve.`;
		}
	}
	return undefined;
}

interface DynamicArrayDeclaration {
	name: string;
	span: Span;
}

type DynamicArrayAllocationState = 'unallocated' | 'allocated' | 'unknown';

/**
 * Rule: a local dynamic array declared as `Dim values() As T` has no storage
 * until ReDim allocates it. This tracks only straight-line local state; nested
 * runtime blocks and helper calls make the state unknown instead of guessed.
 */
function checkUnallocatedDynamicArrayAccess(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const arrays = localDynamicArrayDeclarationsForBody(member.body, activity);
		if (arrays.size === 0) {
			continue;
		}
		const state = new Map<string, DynamicArrayAllocationState>();
		for (const lower of arrays.keys()) {
			state.set(lower, 'unallocated');
		}
		walkStraightLineBody(member.body, (node) => isInactiveNode(activity, node), {
			onStatement: (stmt) =>
				checkUnallocatedDynamicArrayAccessStatement(source, stmt, arrays, state, push),
			touchesInStatement: (stmt) => dynamicArrayTouchesInStatement(source, stmt, arrays),
			demoteToUnknown: (lower) => {
				state.set(lower, 'unknown');
			},
		});
	}
}

function checkUnallocatedDynamicArrayAccessStatement(
	source: string,
	stmt: StatementNode,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
	state: Map<string, DynamicArrayAllocationState>,
	push: PushFn,
): void {
	const redimmed = redimStatementTargets(source, stmt.span);
	if (redimmed.length > 0) {
		for (const target of redimmed) {
			const lower = target.name.toLowerCase();
			if (arrays.has(lower) && target.dimensions.length > 0) {
				state.set(lower, 'allocated');
			}
		}
		return;
	}
	const erased = eraseStatementSimpleTargets(source, stmt.span);
	if (erased.size > 0) {
		for (const lower of erased) {
			if (arrays.has(lower)) {
				state.set(lower, 'unallocated');
			}
		}
		return;
	}
	for (const hit of unallocatedDynamicArrayIndexAccesses(source, stmt.span, arrays, state)) {
		push(
			'unallocatedDynamicArrayAccess',
			`Dynamic array '${hit.name}' is not allocated before indexed access. This will raise Run-time error '9': Subscript out of range.`,
			hit.span,
		);
	}
	for (const hit of unallocatedDynamicArrayBoundCalls(source, stmt.span, arrays, state)) {
		push(
			'unallocatedDynamicArrayAccess',
			`Dynamic array '${hit.name}' is not allocated before ${hit.functionName}. This will raise Run-time error '9': Subscript out of range.`,
			hit.span,
		);
	}
	const assignment = bareAssignmentTarget(source, stmt.span);
	const assignmentLower = assignment?.name.toLowerCase();
	if (assignmentLower && arrays.has(assignmentLower)) {
		state.set(assignmentLower, 'unknown');
	}
	for (const lower of localsPassedAsCallArguments(source, stmt.span, arrays)) {
		if (state.get(lower) === 'unallocated') {
			state.set(lower, 'unknown');
		}
	}
}

function localDynamicArrayDeclarationsForBody(
	body: readonly BodyNode[],
	activity: ConditionalActivityTracker | undefined,
): Map<string, DynamicArrayDeclaration> {
	const out = new Map<string, DynamicArrayDeclaration>();
	forEachVariableGroup(body as BodyNode[], (group) => {
		if (group.isConst || group.modifier === 'Static') {
			return;
		}
		for (const decl of group.declarations) {
			if (!decl.isArray || decl.arrayBounds) {
				continue;
			}
			const lower = decl.name.toLowerCase();
			if (!out.has(lower)) {
				out.set(lower, { name: decl.name, span: decl.span });
			}
		}
	}, activity);
	return out;
}

function unallocatedDynamicArrayIndexAccesses(
	source: string,
	span: Span,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
	state: ReadonlyMap<string, DynamicArrayAllocationState>,
): Array<{ name: string; span: Span }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const out: Array<{ name: string; span: Span }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (toks[i + 1].rawText !== '(' || toks[i - 1]?.rawText === '.') {
			continue;
		}
		const name = tokenName(toks[i]);
		const lower = name?.toLowerCase();
		if (!name || !lower || !arrays.has(lower) || state.get(lower) !== 'unallocated') {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close <= i + 1) {
			continue;
		}
		out.push({
			name,
			span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
		});
	}
	return out;
}

function unallocatedDynamicArrayBoundCalls(
	source: string,
	span: Span,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
	state: ReadonlyMap<string, DynamicArrayAllocationState>,
): Array<{ functionName: string; name: string; span: Span }> {
	const toks = statementTokens(source, span);
	const out: Array<{ functionName: string; name: string; span: Span }> = [];
	for (let i = 0; i < toks.length - 2; i++) {
		const functionName = tokenName(toks[i]);
		const lowerFunction = functionName?.toLowerCase();
		if (lowerFunction !== 'lbound' && lowerFunction !== 'ubound') {
			continue;
		}
		if (toks[i + 1]?.rawText !== '(' || !isBareOrVbaQualifiedIntrinsicCall(toks, i)) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const split = splitArgSlots(toks.slice(i + 2, close), span.start);
		const firstSlot = split.slots[0] ?? [];
		if (firstSlot.length !== 1) {
			continue;
		}
		const name = tokenName(firstSlot[0]);
		const lower = name?.toLowerCase();
		if (!name || !lower || !arrays.has(lower) || state.get(lower) !== 'unallocated') {
			continue;
		}
		out.push({
			functionName: functionName!,
			name,
			span: split.spans[0] ?? { start: span.start + firstSlot[0].start, end: span.start + firstSlot[0].end },
		});
	}
	return out;
}

function dynamicArrayTouchesInStatement(
	source: string,
	stmt: StatementNode,
	arrays: ReadonlyMap<string, DynamicArrayDeclaration>,
): Set<string> {
	const out = new Set<string>();
	for (const target of redimStatementTargets(source, stmt.span)) {
		const lower = target.name.toLowerCase();
		if (arrays.has(lower)) {
			out.add(lower);
		}
	}
	for (const lower of eraseStatementSimpleTargets(source, stmt.span)) {
		if (arrays.has(lower)) {
			out.add(lower);
		}
	}
	const assignment = bareAssignmentTarget(source, stmt.span);
	const assignmentLower = assignment?.name.toLowerCase();
	if (assignmentLower && arrays.has(assignmentLower)) {
		out.add(assignmentLower);
	}
	for (const lower of localsPassedAsCallArguments(source, stmt.span, arrays)) {
		out.add(lower);
	}
	return out;
}

function eraseStatementSimpleTargets(source: string, span: Span): Set<string> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'erase') {
		return new Set();
	}
	const out = new Set<string>();
	for (const group of splitTopLevelTokenGroups(toks.slice(1), ',')) {
		const content = group.filter((tok) => tok.kind !== 'comment');
		if (content.length !== 1) {
			continue;
		}
		const name = tokenName(content[0]);
		if (name) {
			out.add(name.toLowerCase());
		}
	}
	return out;
}

/**
 * Rule: Erase targets must be variable/array target names, not arbitrary
 * expressions. This intentionally stays syntax-shaped: array-ness/type
 * resolution is a separate binder-backed slice.
 */
function checkEraseTargets(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			for (const hit of invalidEraseTargets(source, stmt.span)) {
				push(
					'invalidEraseTarget',
					`Erase target must be a variable or array name, not an arbitrary expression.`,
					hit.span,
				);
			}
			for (const hit of eraseScalarTargets(
				source,
				stmt.span,
				shapes,
				(name) => declaredShapeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'assignmentTarget',
				),
			)) {
				push(
					'eraseRequiresArray',
					`Erase target '${hit.name}' must be an array or Variant, but it is declared As ${hit.asType}.`,
					hit.span,
				);
			}
		}, activity);
	}
}

function invalidEraseTargets(source: string, span: Span): Array<{ span: Span }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'erase') {
		return [];
	}
	const out: Array<{ span: Span }> = [];
	for (const group of splitTopLevelTokenGroups(toks.slice(1), ',')) {
		const content = group.filter((tok) => tok.kind !== 'comment');
		if (content.length === 0) {
			continue;
		}
		if (eraseTargetLooksVariableLike(content)) {
			continue;
		}
		out.push({ span: tokenGroupSpan(span, content) });
	}
	return out;
}

function eraseScalarTargets(
	source: string,
	span: Span,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	resolveShape?: (name: string) => SourceDeclaredShape,
): Array<{ name: string; span: Span; asType: string }> {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'erase') {
		return [];
	}
	const out: Array<{ name: string; span: Span; asType: string }> = [];
	for (const group of splitTopLevelTokenGroups(toks.slice(1), ',')) {
		const content = group.filter((tok) => tok.kind !== 'comment');
		if (content.length !== 1) {
			continue;
		}
		const name = tokenName(content[0]);
		if (!name) {
			continue;
		}
		const resolvedShape = resolveShape?.(name);
		const shape = resolvedShape?.resolved
			? resolvedShape.shape
			: shapes.get(name.toLowerCase());
		if (!shape || shape.isArray || !shape.asType) {
			continue;
		}
		const normalized = normalizeType(shape.asType);
		if (!normalized || normalized === 'variant') {
			continue;
		}
		if (normalized === 'object' || isKnownScalarType(normalized)) {
			out.push({ name, span: tokenGroupSpan(span, content), asType: shape.asType });
		}
	}
	return out;
}

function eraseTargetLooksVariableLike(toks: readonly VbaToken[]): boolean {
	if (!tokenName(toks[0])) {
		return false;
	}
	if (toks.some((tok) => ERASE_EXPRESSION_OPERATORS.has(tok.rawText))) {
		return false;
	}
	if (toks[0]?.rawText === '(') {
		return false;
	}
	return true;
}

const ERASE_EXPRESSION_OPERATORS = new Set([
	'+',
	'-',
	'*',
	'/',
	'\\',
	'^',
	'&',
	'=',
	'<',
	'>',
	'<=',
	'>=',
	'<>',
]);

function tokenGroupSpan(base: Span, toks: readonly VbaToken[]): Span {
	const first = toks[0];
	const last = toks[toks.length - 1];
	return {
		start: base.start + (first?.start ?? 0),
		end: base.start + (last?.end ?? 0),
	};
}

function splitTopLevelTokenGroups(
	toks: readonly VbaToken[],
	separator: string,
): VbaToken[][] {
	const groups: VbaToken[][] = [];
	let current: VbaToken[] = [];
	let depth = 0;
	for (const tok of toks) {
		if (tok.rawText === '(') {
			depth++;
		} else if (tok.rawText === ')') {
			depth = Math.max(0, depth - 1);
		}
		if (depth === 0 && tok.rawText === separator) {
			groups.push(current);
			current = [];
			continue;
		}
		current.push(tok);
	}
	groups.push(current);
	return groups;
}

/**
 * Rule: object modules cannot expose certain public declarations as object
 * members. VBE reports one compile error family for public constants,
 * fixed-length strings, arrays, user-defined types, and Declare statements in
 * class/document/UserForm modules.
 */
function checkObjectModulePublicMembers(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!isObjectModuleKind(moduleKind)) {
		return;
	}

	const report = (kind: string, span: Span): void => {
		push(
			'objectModulePublicMember',
			`Public ${kind} are not allowed as Public members of object modules; VBE Compile rejects this declaration.`,
			span,
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && isPublicModifier(member.modifier)) {
			for (const decl of member.declarations) {
				const span = declaredNameSpan(source, decl.span, decl.name);
				if (member.isConst) {
					report('constants', span);
				} else if (decl.isArray) {
					report('arrays', span);
				} else if (decl.fixedLength !== undefined) {
					report('fixed-length strings', span);
				}
			}
			continue;
		}

		if (member.kind === 'Type' && isPublicModifier(member.visibility)) {
			report('user-defined types', declaredNameSpan(source, member.span, member.name));
			continue;
		}

		if (member.kind === 'Declare' && isPublicModifier(member.visibility)) {
			report('Declare statements', declaredNameSpan(source, member.span, member.name));
		}
	}
}

/**
 * Rule: `Event` declarations are object-module declarations. Standard modules
 * can contain ordinary procedures that look like handlers, but not `Event`
 * declarations themselves.
 */
function checkEventDeclarationModuleKind(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (isObjectModuleKind(moduleKind)) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Event') {
			continue;
		}
		push(
			'eventDeclarationModuleKind',
			`Event declaration '${member.name}' is only valid in class, document, or UserForm modules.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

/**
 * Rule: `WithEvents` is a module-level object-module variable declarator. The
 * parser exposes the relevant settled facts directly: module kind, local vs
 * module declaration context, `As New`, and array declarators.
 */
function checkWithEventsDeclarations(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (message: string, span: Span): void => {
		push('withEventsDeclaration', message, span);
	};
	const inspect = (group: VariableGroupNode, insideProcedure: boolean): void => {
		if (!group.withEvents || isInactiveNode(activity, group)) {
			return;
		}
		for (const decl of group.declarations) {
			const nameSpan = declaredNameSpan(source, decl.span, decl.name);
			if (insideProcedure) {
				report(
					`WithEvents variable '${decl.name}' must be declared at module level.`,
					nameSpan,
				);
				continue;
			}
			if (!isObjectModuleKind(moduleKind)) {
				report(
					`WithEvents variable '${decl.name}' is only valid in class, document, or UserForm modules.`,
					nameSpan,
				);
				continue;
			}
			if (decl.isNew) {
				report(
					`WithEvents variable '${decl.name}' cannot be declared As New.`,
					nameSpan,
				);
			}
			if (decl.isArray) {
				report(
					`WithEvents variable '${decl.name}' cannot be an array.`,
					nameSpan,
				);
			}
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspect(member, false);
			continue;
		}
		if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, (group) => inspect(group, true), activity);
		}
	}
}

/**
 * Rule: `Friend` is procedure visibility for object modules. It is not a
 * module-variable modifier and it is not valid for standard-module procedures.
 */
function checkFriendDeclarations(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			if (hasFriendModifier(member.modifiers) && !isObjectModuleKind(moduleKind)) {
				push(
					'friendDeclaration',
					`Friend procedure '${member.name}' is only valid in class, document, or UserForm modules.`,
					friendKeywordSpan(source, member.span),
				);
			}
			continue;
		}
		if (member.kind !== 'VariableGroup' || member.modifier.toLowerCase() !== 'friend') {
			continue;
		}
		push(
			'friendDeclaration',
			'Friend can only modify procedure declarations, not variables.',
			friendKeywordSpan(source, member.span),
		);
	}
}

function hasFriendModifier(modifiers: readonly string[]): boolean {
	return modifiers.some((modifier) => modifier.toLowerCase() === 'friend');
}

function friendKeywordSpan(source: string, span: Span): Span {
	const tok = statementTokensAfterLeadingLabel(source, span)
		.find((token) => tokenText(token) === 'friend');
	return tok ? absoluteSpan(span, tok) : firstTokenSpan(source, span);
}

/**
 * Rule: `Implements` is an object-module declaration-section statement. This
 * intentionally validates only placement/module-kind facts, leaving interface
 * member completeness to a later project-binder pass.
 */
function checkImplementsStatementPlacement(
	source: string,
	mod: ModuleNode,
	moduleKind: ModuleSymbolKind,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let procedureSeen = false;
	const reportModuleKind = (hit: ImplementsStatementHit): void => {
		push(
			'implementsStatementPlacement',
			`Implements statement '${hit.name}' is only valid in class, document, or UserForm modules.`,
			hit.span,
		);
	};
	const reportProcedurePlacement = (hit: ImplementsStatementHit): void => {
		push(
			'implementsStatementPlacement',
			`Implements statement '${hit.name}' must appear in the module declaration section before any procedure.`,
			hit.span,
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			procedureSeen = true;
			forEachStatement(member.body, (stmt) => {
				const hit = implementsStatementHit(source, stmt.span);
				if (hit) {
					reportProcedurePlacement(hit);
				}
			}, activity);
			continue;
		}
		if (member.kind !== 'Statement') {
			continue;
		}
		const hit = implementsStatementHit(source, member.span);
		if (!hit) {
			continue;
		}
		if (!isObjectModuleKind(moduleKind)) {
			reportModuleKind(hit);
			continue;
		}
		if (procedureSeen) {
			reportProcedurePlacement(hit);
		}
	}
}

interface ImplementsStatementHit {
	name: string;
	span: Span;
}

function implementsStatementHit(source: string, span: Span): ImplementsStatementHit | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (tokenText(toks[0]) !== 'implements') {
		return undefined;
	}
	const firstName = tokenName(toks[1]);
	if (!firstName) {
		return undefined;
	}
	let name = firstName;
	let endIndex = 1;
	while (toks[endIndex + 1]?.rawText === '.') {
		const part = tokenName(toks[endIndex + 2]);
		if (!part) {
			break;
		}
		name += `.${part}`;
		endIndex += 2;
	}
	return {
		name,
		span: {
			start: span.start + toks[1].start,
			end: span.start + toks[endIndex].end,
		},
	};
}

/**
 * Rule: `RaiseEvent` names an Event declared by the containing module. Event
 * signature/arity checks remain deferred to the richer event-binding slice.
 */
function checkRaiseEventTargets(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const events = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Event' && member.name) {
			events.add(member.name.toLowerCase());
		}
	}

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachProcedureBodyLine(source, member, (lineSpan) => {
			if (activity?.isInactive(lineSpan)) {
				return;
			}
			const hit = raiseEventTargetHit(source, lineSpan);
			if (!hit || events.has(hit.name.toLowerCase())) {
				return;
			}
			push(
				'raiseEventUndeclaredEvent',
				`Event '${hit.name}' is not declared in this module, so it cannot be raised with RaiseEvent.`,
				hit.span,
			);
		});
	}
}

function raiseEventTargetHit(
	source: string,
	span: Span,
): { name: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	const start = raiseEventStatementStartIndex(toks);
	if (start < 0) {
		return undefined;
	}
	const nameTok = toks[start + 1];
	if (!nameTok) {
		return undefined;
	}
	const name = tokenName(nameTok);
	if (!name) {
		return undefined;
	}
	return {
		name,
		span: {
			start: span.start + nameTok.start,
			end: span.start + nameTok.end,
		},
	};
}

function raiseEventStatementStartIndex(toks: readonly VbaToken[]): number {
	let start = 0;
	if (toks.length > 1 && /^\d+$/.test(toks[0].rawText)) {
		start = 1;
	} else if (
		toks.length > 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		start = 2;
	}
	return tokenText(toks[start]) === 'raiseevent' ? start : -1;
}

function checkDeclarePtrSafeForWin64(
	source: string,
	mod: ModuleNode,
	conditionalCompilation: ConditionalCompilationEnvironment | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!conditionalValueTruthy(conditionalCompilerConstants(conditionalCompilation).get('win64'))) {
		return;
	}
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Declare' || member.ptrSafe) {
			continue;
		}
		push(
			'declareMissingPtrSafe',
			`Declare statement '${member.name}' must include PtrSafe when compiling for 64-bit Office.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function conditionalValueTruthy(value: unknown): boolean {
	if (typeof value === 'boolean') {
		return value;
	}
	if (typeof value === 'number') {
		return value !== 0;
	}
	return typeof value === 'string' && value.length > 0;
}

function checkEventHandlerModuleScope(
	source: string,
	mod: ModuleNode,
	moduleName: string,
	moduleKind: ModuleSymbolKind,
	documentType: EventHandlerDocumentType | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const actualDocumentType = eventHandlerDocumentTypeForContext({
		moduleName,
		moduleKind,
		documentType,
	});
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure' || member.procKind !== 'Sub') {
			continue;
		}
		const event = eventHandlerProcedureForName(member.name);
		if (!event) {
			continue;
		}
		if (actualDocumentType === event.documentType) {
			continue;
		}
		const moduleDescription =
			moduleKind === 'document'
				? `${describeEventDocumentType(actualDocumentType)} document module`
				: `${moduleKind} module`;
		push(
			'eventHandlerWrongModule',
			`'${event.name}' matches a ${event.owner} event handler, but this ${moduleDescription} is not where Excel wires that event. It will behave like an ordinary procedure here.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function describeEventDocumentType(
	documentType: EventHandlerDocumentType | undefined,
): string {
	switch (documentType) {
		case 'workbook':
			return 'workbook';
		case 'worksheet':
			return 'worksheet';
		case 'chart':
			return 'chart';
		default:
			return 'unknown';
	}
}

function isPublicModifier(value: string | undefined): boolean {
	return value?.toLowerCase() === 'public';
}

/**
 * Rule: a `Call` statement must wrap its arguments in parentheses. After the
 * `Call` keyword the callee chain (identifier, then any run of `.member` or
 * `(...)` groups) is consumed; any token left over is an unparenthesised
 * argument - the VBE "Expected: (" error. Unbalanced parentheses are left to the
 * dedicated rule.
 */
function checkCallParens(
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
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		forEachStatement(member.body, (stmt) => {
			const invalidCallTarget = invalidExplicitCallTarget(source, stmt.span, moduleSignatures, sourceNames);
			if (invalidCallTarget) {
				push(
					'invalidExplicitCallTarget',
					`'${invalidCallTarget.name}' cannot be used as the target of an explicit Call statement.`,
					invalidCallTarget.span,
				);
				return;
			}
			const at = explicitCallStatementArgumentWithoutParens(source, stmt.span);
			if (at) {
				push(
					'callRequiresParens',
					'A Call statement requires parentheses around its argument list.',
					at,
				);
			}
			const bare = implicitParenthesizedBareCallableCall(source, stmt.span, moduleSignatures, sourceNames);
			if (bare) {
				push(
					'callStatementForbidsParens',
					bareCallForbidsParensMessage(bare.name, moduleSignatures, sourceNames),
					bare.span,
				);
			}
			const implicit = implicitParenthesizedMemberCall(source, stmt.span, memberCtx);
			if (implicit) {
				push(
					'callStatementForbidsParens',
					'Standalone zero-argument member calls cannot use empty parentheses unless they are prefixed with Call or used in an expression.',
					implicit.span,
				);
			}
		}, activity);
	}
}

function bareCallForbidsParensMessage(
	name: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
): string {
	const runtime = !moduleSignatures.has(name.toLowerCase()) &&
		!runtimeCallableSourceShadowed(name, sourceNames)
		? resolveRuntimeFunction(name)
		: undefined;
	if (runtime && !runtimeAllowsExplicitCall(runtime)) {
		return `Standalone '${runtime.name}()' cannot use empty parentheses in statement context; use '${runtime.name}' as a statement or use it in an expression.`;
	}
	return 'Standalone zero-argument procedure calls cannot use empty parentheses unless they are prefixed with Call or used in an expression.';
}

function invalidExplicitCallTarget(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
): { name: string; span: Span } | undefined {
	const target = explicitCallStatementTarget(source, span);
	if (!target) {
		return undefined;
	}
	if (
		moduleSignatures.has(target.name.toLowerCase()) ||
		runtimeCallableSourceShadowed(target.name, sourceNames)
	) {
		return undefined;
	}
	const runtime = resolveRuntimeFunction(target.name);
	if (!runtime || runtimeAllowsExplicitCall(runtime)) {
		return undefined;
	}
	return { name: runtime.name, span: target.span };
}

function checkInvalidExpressionSyntax(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			const incompleteMember = incompleteMemberAccess(source, stmt.span, {
				scalarTypes: env,
				resolveScalarType: (name) => declaredTypeForSourceBinding(
					symbols,
					procSym,
					projectVisibleSymbols,
					name,
					'memberReceiver',
				),
			});
			if (incompleteMember) {
				push(
					'invalidExpressionSyntax',
					"Incomplete member access: type a member name after '.'.",
					incompleteMember.span,
				);
				return;
			}
			const unsupportedQuestion = unsupportedQuestionMarkOperator(source, stmt.span);
			if (unsupportedQuestion) {
				push(
					'invalidExpressionSyntax',
					"VBA does not support the '?' conditional operator in code modules; use If...Then...Else, or IIf(...) only when both branches are safe to evaluate.",
					unsupportedQuestion.span,
				);
				return;
			}
			const hit = invalidOperatorSequence(source, stmt.span);
			if (hit) {
				push(
					'invalidExpressionSyntax',
					`Invalid operator sequence '${hit.text}'; this will fail to compile as a syntax error.`,
					hit.span,
				);
			}
		}, activity);
	}
}

const NON_UNARY_BINARY_OPERATORS = new Set([
	'*',
	'/',
	'\\',
	'^',
	'&',
	'=',
	'<',
	'>',
	'<=',
	'>=',
	'<>',
	':=',
	'like',
	'is',
	'and',
	'or',
	'xor',
	'eqv',
	'imp',
	'mod',
]);

function invalidOperatorSequence(
	source: string,
	span: Span,
): { text: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	for (let i = 0; i < toks.length; i++) {
		if (!isNonUnaryBinaryOperator(toks[i])) {
			continue;
		}
		let end = i;
		while (isNonUnaryBinaryOperator(toks[end + 1])) {
			end++;
		}
		if (end > i) {
			const first = toks[i];
			const last = toks[end];
			return {
				text: source.slice(span.start + first.start, span.start + last.end),
				span: { start: span.start + first.start, end: span.start + last.end },
			};
		}
		if (i === toks.length - 1) {
			return {
				text: toks[i].rawText,
				span: absoluteSpan(span, toks[i]),
			};
		}
	}
	return undefined;
}

function unsupportedQuestionMarkOperator(
	source: string,
	span: Span,
): { span: Span } | undefined {
	const question = statementTokens(source, span).find(
		(tok) => tok.kind === 'operator' && tok.rawText === '?',
	);
	return question ? { span: absoluteSpan(span, question) } : undefined;
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

function incompleteMemberAccess(
	source: string,
	span: Span,
	options: {
		includeLeadingDot?: boolean;
		scalarTypes?: ReadonlyMap<string, string>;
		resolveScalarType?: SourceDeclaredTypeResolver;
	} = {},
): { span: Span } | undefined {
	const toks = statementTokens(source, span);
	for (let i = 0; i < toks.length; i++) {
		const tok = toks[i];
		if (tok.rawText !== '.') {
			continue;
		}
		if (i === 0 && !options.includeLeadingDot) {
			continue;
		}
		const next = toks[i + 1];
		if (next && tokenName(next)) {
			continue;
		}
		const receiverName = i > 0 ? tokenName(toks[i - 1]) : undefined;
		if (receiverName) {
			const resolvedType = options.resolveScalarType?.(receiverName);
			const asType = resolvedType?.resolved
				? resolvedType.asType
				: options.scalarTypes?.get(receiverName.toLowerCase());
			const normalized = normalizeType(asType);
			if (normalized && isKnownScalarType(normalized)) {
				continue;
			}
		}
		return { span: absoluteSpan(span, tok) };
	}
	return undefined;
}

function isNonUnaryBinaryOperator(tok: VbaToken | undefined): boolean {
	if (!tok || tok.kind !== 'operator') {
		return false;
	}
	return NON_UNARY_BINARY_OPERATORS.has(tokenText(tok));
}

function checkDivisionByZeroExpressions(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const projectConstants = resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map());
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, projectConstants);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procedureConstants = new Map(moduleConstants);
		collectBodyLiteralIntegerConstants(member.body, procedureConstants, activity);
		const procSym = procedureSymbolFor(symbols, member);
		const constants = scopedIntegerConstantLookup(
			procedureConstants,
			symbols,
			procSym,
			projectVisibleSymbols,
		);
		forEachStatement(member.body, (stmt) => {
			for (const hit of divisionByZeroDivisors(source, stmt.span, constants)) {
				push(
					'divisionByZero',
					`Expression uses '${hit.operator}' with a zero divisor. This will raise Run-time error '11': Division by zero.`,
					hit.span,
				);
			}
		}, activity);
	}
}

function divisionByZeroDivisors(
	source: string,
	span: Span,
	constants: IntegerConstantLookup,
): Array<{ operator: string; span: Span }> {
	const toks = statementTokens(source, span);
	const hits: Array<{ operator: string; span: Span }> = [];
	for (let i = 0; i < toks.length; i++) {
		const operator = divisionByZeroOperatorLabel(toks[i]);
		if (!operator) {
			continue;
		}
		const divisor = zeroDivisorToken(source, span, toks, i + 1, constants);
		if (divisor) {
			hits.push({ operator, span: absoluteTokenGroupSpan(span, divisor) });
		}
	}
	return hits;
}

function divisionByZeroOperatorLabel(tok: VbaToken | undefined): string | undefined {
	const text = tokenText(tok);
	if (text === '/' || text === '\\') {
		return text;
	}
	return text === 'mod' ? 'Mod' : undefined;
}

function zeroDivisorToken(
	source: string,
	span: Span,
	toks: VbaToken[],
	start: number,
	constants: IntegerConstantLookup,
): VbaToken[] | undefined {
	const first = toks[start];
	if (!first) {
		return undefined;
	}
	if (first.rawText === '(') {
		const close = matchParenFrom(toks, start);
		if (close < 0) {
			return undefined;
		}
		return zeroDivisorExpression(source, span, toks, start + 1, close, constants);
	}
	if (
		first.kind === 'operator' &&
		(first.rawText === '+' || first.rawText === '-')
	) {
		const signed = zeroDivisorAtomTokenGroup(toks, start + 1, constants);
		return signed ? [first, ...signed] : undefined;
	}
	return zeroDivisorAtomTokenGroup(toks, start, constants);
}

function zeroDivisorExpression(
	source: string,
	span: Span,
	toks: VbaToken[],
	start: number,
	endExclusive: number,
	constants: IntegerConstantLookup,
): VbaToken[] | undefined {
	if (start >= endExclusive) {
		return undefined;
	}
	const folded = foldIntegerExpressionTokens(source, span, toks, start, endExclusive, constants);
	if (folded === 0) {
		return toks.slice(start, endExclusive);
	}
	if (toks[start]?.rawText === '(') {
		const close = matchParenFrom(toks, start);
		if (close === endExclusive - 1) {
			return zeroDivisorExpression(source, span, toks, start + 1, close, constants);
		}
	}
	if (
		endExclusive === start + 2 &&
		toks[start]?.kind === 'operator' &&
		(toks[start].rawText === '+' || toks[start].rawText === '-') &&
		isZeroDivisorAtom(toks[start + 1], constants)
	) {
		return [toks[start], toks[start + 1]];
	}
	if (endExclusive === start + 1 && isZeroDivisorAtom(toks[start], constants)) {
		return [toks[start]];
	}
	return undefined;
}

function zeroDivisorAtomTokenGroup(
	toks: readonly VbaToken[],
	start: number,
	constants: IntegerConstantLookup,
): VbaToken[] | undefined {
	const first = toks[start];
	const firstName = first ? tokenName(first) : undefined;
	const member = toks[start + 2];
	const memberName = member ? tokenName(member) : undefined;
	if (firstName && toks[start + 1]?.rawText === '.' && memberName) {
		return constants.get(`${firstName}.${memberName}`.toLowerCase()) === 0
			? [first, toks[start + 1], member]
			: undefined;
	}
	return isZeroDivisorAtom(first, constants) ? [first] : undefined;
}

function isZeroDivisorAtom(
	tok: VbaToken | undefined,
	constants: IntegerConstantLookup,
): boolean {
	if (isZeroNumericLiteral(tok)) {
		return true;
	}
	const name = tok ? tokenName(tok) : undefined;
	return name !== undefined && constants.get(name.toLowerCase()) === 0;
}

function isZeroNumericLiteral(tok: VbaToken | undefined): boolean {
	if (!tok || (tok.kind !== 'integerLiteral' && tok.kind !== 'floatLiteral')) {
		return false;
	}
	const normalized = tok.rawText
		.replace(/[!#@%&^]$/, '')
		.replace(/[dD]/g, 'E');
	const hex = /^&[hH]([0-9A-Fa-f]+)$/.exec(normalized);
	if (hex) {
		return Number.parseInt(hex[1], 16) === 0;
	}
	const octal = /^&[oO]([0-7]+)$/.exec(normalized);
	if (octal) {
		return Number.parseInt(octal[1], 8) === 0;
	}
	if (!/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalized)) {
		return false;
	}
	return Number(normalized) === 0;
}

function absoluteTokenGroupSpan(base: Span, toks: readonly VbaToken[]): Span {
	return { start: base.start + toks[0].start, end: base.start + toks[toks.length - 1].end };
}

/**
 * Rule: when a Function is used inside an expression, its argument list must be
 * parenthesized (`x = Foo(1, 2)`). The parenless form (`Foo 1, 2`) is only a
 * call-statement form and becomes a VBE syntax error after `=`.
 */
function checkExpressionCallParens(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const functions = expressionCallableFunctionNames(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const sourceNames = sourceNameScopeFor(symbols, member, projectVisibleSymbols);
		forEachStatement(member.body, (stmt) => {
			const hit = parenlessExpressionCall(source, stmt.span, functions, sourceNames);
			if (hit) {
				push(
					'expressionCallRequiresParens',
					`Function call arguments in an expression must be enclosed in parentheses: use '${hit.name}(...)'.`,
					hit.span,
				);
			}
		}, activity);
	}
}

interface ExpressionCallableFunctions {
	bare: Set<string>;
	qualified: Set<string>;
}

function expressionCallableFunctionNames(
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
): ExpressionCallableFunctions {
	const bare = new Set<string>();
	const qualified = new Set<string>();
	for (const member of symbols.root.children ?? []) {
		if (member.kind === 'function' || member.kind === 'propertyGet') {
			bare.add(member.name.toLowerCase());
		}
	}
	for (const [key, candidates] of projectProcedures ?? []) {
		if (candidates.length !== 1 || candidates[0].kind !== 'function') {
			continue;
		}
		if (key.includes('.')) {
			qualified.add(key);
		} else if (!bare.has(key)) {
			bare.add(key);
		}
	}
	return { bare, qualified };
}

function parenlessExpressionCall(
	source: string,
	span: Span,
	functions: ExpressionCallableFunctions,
	sourceNames?: SourceNameScope,
): { name: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	if (toks.length === 0 || isNonAssignmentStatementLeader(tokenText(toks[0]))) {
		return undefined;
	}
	const eq = topLevelOperatorIndex(toks, '=');
	if (eq < 0) {
		return undefined;
	}

	for (let i = eq + 1; i < toks.length - 1; i++) {
		const tok = toks[i];
		const name = tokenName(tok);
		if (!name || !isExpressionCallableAt(toks, i, name, functions, sourceNames)) {
			continue;
		}
		if (i > eq + 1 && toks[i - 1].rawText === '.') {
			const qualifier = tokenName(toks[i - 2]);
			if (!qualifier || !functions.qualified.has(qualifiedProcedureKey(qualifier, name))) {
				continue; // object member calls need receiver typing before we can be precise
			}
		}
		const next = toks[i + 1];
		if (!isParenlessArgumentStart(next)) {
			continue;
		}
		const gap = source.slice(span.start + tok.end, span.start + next.start);
		if (!/\s/.test(gap)) {
			continue;
		}
		return {
			name,
			span: { start: span.start + tok.start, end: span.start + tok.end },
		};
	}
	return undefined;
}

function isExpressionCallableAt(
	toks: readonly VbaToken[],
	index: number,
	name: string,
	functions: ExpressionCallableFunctions,
	sourceNames?: SourceNameScope,
): boolean {
	if (index > 1 && toks[index - 1].rawText === '.') {
		const qualifier = tokenName(toks[index - 2]);
		return qualifier
			? functions.qualified.has(qualifiedProcedureKey(qualifier, name))
			: false;
	}
	if (index > 0 && toks[index - 1].rawText === '.') {
		return false;
	}
	if (bareCallableSourceShadowed(name, sourceNames)) {
		return false;
	}
	if (functions.bare.has(name.toLowerCase())) {
		return true;
	}
	if (runtimeCallableSourceShadowed(name, sourceNames)) {
		return false;
	}
	return resolveRuntimeFunction(name)?.kind === 'function';
}

function isParenlessArgumentStart(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return false;
	}
	switch (tok.kind) {
		case 'identifier':
		case 'bracketedIdentifier':
		case 'integerLiteral':
		case 'floatLiteral':
		case 'stringLiteral':
		case 'dateLiteral':
			return true;
		case 'keyword':
			return !isInfixExpressionKeyword(tok.rawText);
		default:
			return false;
	}
}

function isInfixExpressionKeyword(text: string): boolean {
	switch (text.toLowerCase()) {
		case 'and':
		case 'or':
		case 'xor':
		case 'eqv':
		case 'imp':
		case 'is':
		case 'mod':
			return true;
		default:
			return false;
	}
}

function isNonAssignmentStatementLeader(word: string): boolean {
	switch (word) {
		case 'if':
		case 'elseif':
		case 'for':
		case 'do':
		case 'loop':
		case 'while':
		case 'select':
		case 'case':
			return true;
		default:
			return false;
	}
}

function implicitParenthesizedBareCallableCall(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): { name: string; span: Span } | undefined {
	const call = standaloneEmptyParenthesizedCallStatement(source, span);
	if (!call || call.isMember) {
		return undefined;
	}
	const signature = callableSignatureFor(call.name, moduleSignatures, sourceNames);
	if (!signature || !callableAcceptsZeroArguments(signature)) {
		return undefined;
	}
	return {
		name: call.name,
		span: call.span,
	};
}

function implicitParenthesizedMemberCall(
	source: string,
	span: Span,
	memberCtx: MemberCompletionContext,
): { name: string; span: Span } | undefined {
	const call = standaloneEmptyParenthesizedCallStatement(source, span);
	if (!call || !call.isMember) {
		return undefined;
	}
	if (
		call.startsWithLeadingDot &&
		!resolveExactMemberCompletion(source, call.name, call.calleeEndOffset, memberCtx)
	) {
		return undefined;
	}
	return { name: call.name, span: call.span };
}

/** Index of the `)` matching the `(` at `open`, or -1 if unbalanced. */
/**
 * Rule: an `Exit Sub` / `Exit Function` / `Exit Property` must match the kind of
 * the procedure that encloses it (the three Property accessors all map to
 * `Property`). `Exit Do` / `Exit For` are loop exits and are ignored here.
 */
function checkExitStatements(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const expected = expectedExitWord(member.procKind);
		const label = enclosingProcLabel(member.procKind);
		forEachStatement(member.body, (stmt) => {
			const hit = exitTarget(source, stmt.span);
			if (hit && hit.word !== expected) {
				push(
					'exitWrongProcedure',
					`'Exit ${hit.word}' is not valid inside a ${label}; use 'Exit ${expected}'.`,
					hit.span,
				);
			}
		}, activity);
	}
}

/** Maps a procedure kind to the keyword its `Exit` statement must use. */
function expectedExitWord(kind: ProcedureNode['procKind']): string {
	if (kind === 'Sub') {
		return 'Sub';
	}
	if (kind === 'Function') {
		return 'Function';
	}
	return 'Property';
}

/** Human label for a procedure kind, for diagnostic messages. */
function enclosingProcLabel(kind: ProcedureNode['procKind']): string {
	if (kind === 'Sub') {
		return 'Sub';
	}
	if (kind === 'Function') {
		return 'Function';
	}
	return 'Property procedure';
}

/** If a statement is `Exit Sub|Function|Property`, returns the word and span. */
function exitTarget(
	source: string,
	span: Span,
): { word: string; span: Span } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (toks.length < 2 || toks[0].rawText.toLowerCase() !== 'exit') {
		return undefined;
	}
	const w = toks[1].rawText.toLowerCase();
	let word: string;
	if (w === 'sub') {
		word = 'Sub';
	} else if (w === 'function') {
		word = 'Function';
	} else if (w === 'property') {
		word = 'Property';
	} else {
		return undefined; // Exit Do / Exit For etc.
	}
	return {
		word,
		span: { start: span.start + toks[0].start, end: span.start + toks[1].end },
	};
}

/**
 * Rule: procedure-local control-flow labels must exist in the same procedure.
 * This covers the deterministic VBE "Label not defined" cases without letting
 * labels leak across procedures or across inactive conditional-compilation code.
 */
function checkUndefinedLabels(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const labels = collectProcedureLabels(source, member, activity);
		for (const ref of collectProcedureLabelReferences(source, member, activity)) {
			if (!labels.has(ref.key)) {
				push(
					'undefinedLabel',
					`Label '${ref.text}' is not defined in procedure '${member.name}'.`,
					ref.span,
				);
			}
		}
	}
}

/**
 * Rule: procedure-local labels must be unique within the same procedure. This
 * catches duplicate named labels and normalized decimal line labels.
 */
function checkDuplicateLabels(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const seen = new Set<string>();
		for (const label of collectProcedureLabelDeclarations(source, member, activity)) {
			if (!seen.has(label.key)) {
				seen.add(label.key);
				continue;
			}
			push(
				'duplicateLabel',
				`Label '${label.text}' is already defined in procedure '${member.name}'.`,
				label.span,
			);
		}
	}
}

interface StatementContext {
	forDepth: number;
	doDepth: number;
	withDepth: number;
	selectDepth: number;
}

/**
 * Rules that depend on where a statement appears in the block tree:
 * `If` requires `Then`, `Case` belongs to `Select Case`, a leading `.member`
 * requires `With`, and loop exits require their matching loop.
 */
function checkStatementContext(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const root: StatementContext = {
		forDepth: 0,
		doDepth: 0,
		withDepth: 0,
		selectDepth: 0,
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Statement') {
			continue;
		} else if (member.kind === 'Procedure') {
			checkContextBody(source, member.body, root, activity, push);
		}
	}
}

function checkContextBody(
	source: string,
	body: BodyNode[],
	ctx: StatementContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		switch (node.kind) {
			case 'Statement':
				checkContextStatement(source, node, ctx, push);
				break;
			case 'ForBlock':
				checkForNextControlVariable(source, node, activity, push);
				checkContextBody(
					source,
					node.body,
					{ ...ctx, forDepth: ctx.forDepth + 1 },
					activity,
					push,
				);
				break;
			case 'DoBlock':
				checkContextBody(
					source,
					node.body,
					{ ...ctx, doDepth: ctx.doDepth + 1 },
					activity,
					push,
				);
				break;
			case 'WithBlock':
				checkContextBody(
					source,
					node.body,
					{ ...ctx, withDepth: ctx.withDepth + 1 },
					activity,
					push,
				);
				break;
			case 'SelectBlock':
				checkContextBody(
					source,
					node.body,
					{ ...ctx, selectDepth: ctx.selectDepth + 1 },
					activity,
					push,
				);
				break;
			case 'IfBlock':
			case 'WhileBlock':
				checkContextBody(source, node.body, ctx, activity, push);
				break;
			case 'ConditionalDirective':
			case 'VariableGroup':
				break;
		}
	}
}

function checkForNextControlVariable(
	source: string,
	node: ForBlockNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (
		!node.controlVariable ||
		!node.controlVariableSpan ||
		!node.nextVariable ||
		!node.nextVariableSpan ||
		isInactiveNode(activity, { span: node.controlVariableSpan }) ||
		isInactiveNode(activity, { span: node.nextVariableSpan })
	) {
		return;
	}
	if (node.controlVariable.toLowerCase() === node.nextVariable.toLowerCase()) {
		return;
	}
	push(
		'nextVariableMismatch',
		`Next variable '${node.nextVariable}' does not match active For control variable '${node.controlVariable}'.`,
		node.nextVariableSpan,
	);
}

function checkForEachLoopTypes(
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const shapes = declarationShapeEnvironmentFor(symbols, member);
		const procSym = procedureSymbolFor(symbols, member);
		const resolveShape = (name: string, context: BareIdentifierContext): SourceDeclaredShape =>
			declaredShapeForSourceBinding(
				symbols,
				procSym,
				opts.projectVisibleSymbols,
				name,
				context,
			);
		checkForEachLoopTypesInBody(member.body, shapes, opts, activity, push, resolveShape);
	}
}

function checkForEachLoopTypesInBody(
	body: BodyNode[],
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveShape?: (name: string, context: BareIdentifierContext) => SourceDeclaredShape,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if ('body' in node && Array.isArray(node.body)) {
			if (node.kind === 'ForBlock') {
				checkForEachControlVariableType(node, shapes, opts, activity, push, resolveShape);
				checkForEachSourceType(node, shapes, opts, activity, push, resolveShape);
			}
			checkForEachLoopTypesInBody(node.body, shapes, opts, activity, push, resolveShape);
		}
	}
}

function checkForEachControlVariableType(
	node: ForBlockNode,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveShape?: (name: string, context: BareIdentifierContext) => SourceDeclaredShape,
): void {
	if (
		!node.each ||
		!node.controlVariable ||
		!node.controlVariableSpan ||
		isInactiveNode(activity, { span: node.controlVariableSpan })
	) {
		return;
	}
	const resolvedShape = resolveShape?.(node.controlVariable, 'assignmentTarget');
	const shape = resolvedShape?.resolved
		? resolvedShape.shape
		: shapes.get(node.controlVariable.toLowerCase());
	if (!shape) {
		return;
	}
	const problem = forEachControlVariableTypeProblem(shape, opts);
	if (!problem) {
		return;
	}
	push(
		'forEachControlVariableType',
		`For Each control variable '${node.controlVariable}' must be Variant or Object, but ${problem}.`,
		node.controlVariableSpan,
	);
}

function checkForEachSourceType(
	node: ForBlockNode,
	shapes: ReadonlyMap<string, DeclaredValueShape>,
	opts: AnalyzeModuleOptions,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
	resolveShape?: (name: string, context: BareIdentifierContext) => SourceDeclaredShape,
): void {
	if (
		!node.each ||
		!node.sourceExpression ||
		!node.sourceExpressionSpan ||
		isInactiveNode(activity, { span: node.sourceExpressionSpan })
	) {
		return;
	}
	const sourceName = simpleForEachSourceName(node.sourceExpression);
	if (!sourceName) {
		return;
	}
	const resolvedShape = resolveShape?.(sourceName, 'expression');
	const shape = resolvedShape?.resolved
		? resolvedShape.shape
		: shapes.get(sourceName.toLowerCase());
	if (!shape) {
		return;
	}
	const problem = forEachSourceTypeProblem(shape, opts);
	if (!problem) {
		return;
	}
	push(
		'forEachSourceType',
		`For Each source '${sourceName}' must be a collection object or array, but ${problem}.`,
		node.sourceExpressionSpan,
	);
}

function forEachControlVariableTypeProblem(
	shape: DeclaredValueShape,
	opts: AnalyzeModuleOptions,
): string | undefined {
	if (shape.isArray) {
		return 'it is an array variable';
	}
	if (!shape.asType) {
		return undefined;
	}
	const resolved = resolveTypeName(shape.asType, {
		projectTypes: opts.projectTypes,
		model: opts.hostModel,
	});
	if (resolved?.kind === 'userType') {
		return `it is declared As user-defined Type '${shape.asType}'`;
	}
	if (resolved?.kind === 'enum') {
		return `it is declared As Enum '${shape.asType}'`;
	}
	if (resolved && resolved.kind !== 'primitive') {
		return undefined;
	}
	const normalized = normalizeType(shape.asType);
	if (!normalized || normalized === 'variant' || normalized === 'object') {
		return undefined;
	}
	if (isKnownScalarType(normalized)) {
		return `it is declared As ${shape.asType}`;
	}
	return undefined;
}

function forEachSourceTypeProblem(
	shape: DeclaredValueShape,
	opts: AnalyzeModuleOptions,
): string | undefined {
	if (shape.isArray || !shape.asType) {
		return undefined;
	}
	const resolved = resolveTypeName(shape.asType, {
		projectTypes: opts.projectTypes,
		model: opts.hostModel,
	});
	if (resolved && resolved.kind !== 'primitive') {
		return undefined;
	}
	const normalized = normalizeType(shape.asType);
	if (!normalized || normalized === 'variant' || normalized === 'object') {
		return undefined;
	}
	if (isKnownScalarType(normalized)) {
		return `it is declared As ${shape.asType}`;
	}
	return undefined;
}

function simpleForEachSourceName(sourceExpression: string): string | undefined {
	const toks = statementTokens(sourceExpression, { start: 0, end: sourceExpression.length });
	return toks.length === 1 ? tokenName(toks[0]) : undefined;
}

function checkElseBranchOrder(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	checkConditionalCompilationElseBranchOrder(source, mod, push);
	checkIfBlockElseBranchOrder(source, mod, activity, push);
}

function checkConditionalCompilationElseBranchOrder(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	for (const issue of scanConditionalCompilationBranchOrder(mod).issues) {
		if (issue.kind === 'elseifAfterElse') {
			push(
				'elseBranchOrder',
				"'#ElseIf' cannot appear after '#Else' in the same conditional-compilation block.",
				conditionalDirectiveKeywordSpan(source, issue.directive),
			);
		} else {
			push(
				'elseBranchOrder',
				"Only one '#Else' branch is allowed in a conditional-compilation block.",
				conditionalDirectiveKeywordSpan(source, issue.directive),
			);
		}
	}
}

function checkIfBlockElseBranchOrder(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			checkIfBlockElseBranchOrderInBody(source, member.body, activity, push);
		}
	}
}

function checkIfBlockElseBranchOrderInBody(
	source: string,
	body: BodyNode[],
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'IfBlock') {
			checkSingleIfBlockElseBranchOrder(source, node, activity, push);
		}
		if ('body' in node && Array.isArray(node.body)) {
			checkIfBlockElseBranchOrderInBody(source, node.body, activity, push);
		}
	}
}

function checkSingleIfBlockElseBranchOrder(
	source: string,
	node: { body: BodyNode[] },
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let seenElse = false;
	for (const child of node.body) {
		if (isInactiveNode(activity, child)) {
			continue;
		}
		if (child.kind !== 'Statement') {
			continue;
		}
		const toks = statementTokensAfterLeadingLabel(source, child.span);
		const first = toks[0];
		const word = first ? tokenText(first) : undefined;
		if (word === 'elseif' && seenElse) {
			push(
				'elseBranchOrder',
				"'ElseIf' cannot appear after 'Else' in the same If block.",
				absoluteSpan(child.span, first),
			);
		} else if (word === 'else') {
			if (seenElse) {
				push(
					'elseBranchOrder',
					"Only one 'Else' branch is allowed in an If block.",
					absoluteSpan(child.span, first),
				);
			}
			seenElse = true;
		}
	}
}

function conditionalDirectiveKeywordSpan(
	source: string,
	directive: { span: Span },
): Span {
	const tokens = tokenize(source.slice(directive.span.start, directive.span.end))
		.filter((token) => token.kind !== 'comment' && token.kind !== 'newline');
	const marker = tokens[0];
	const keyword = tokens[1];
	if (marker?.kind === 'directive' && keyword) {
		return {
			start: directive.span.start + marker.start,
			end: directive.span.start + keyword.end,
		};
	}
	return directive.span;
}

function checkContextStatement(
	source: string,
	stmt: StatementNode,
	ctx: StatementContext,
	push: PushFn,
): void {
	const toks = statementTokensAfterLeadingLabel(source, stmt.span);
	const first = toks[0];
	if (!first) {
		return;
	}
	const w0 = tokenText(first);

	if (w0 === 'if' && !toks.some((t) => tokenText(t) === 'then')) {
		push(
			'ifMissingThen',
			"If statement is missing 'Then'.",
			absoluteSpan(stmt.span, first),
		);
	}

	if (w0 === 'case' && ctx.selectDepth === 0) {
		push(
			'caseOutsideSelect',
			"'Case' can only appear inside a 'Select Case' block.",
			absoluteSpan(stmt.span, first),
		);
	}

	if (first.rawText === '.' && ctx.withDepth === 0) {
		push(
			'memberAccessOutsideWith',
			"A statement that starts with '.' must be inside a With block.",
			absoluteSpan(stmt.span, first),
		);
	}
	const leadingMember = toks[1];
	if (first.rawText === '.' && ctx.withDepth > 0 && (!leadingMember || !tokenName(leadingMember))) {
		push(
			'invalidExpressionSyntax',
			"Incomplete member access: type a member name after '.'.",
			absoluteSpan(stmt.span, first),
		);
	}

	if (w0 === 'exit') {
		const target = toks[1];
		const targetWord = tokenText(target);
		if (target && targetWord === 'for' && ctx.forDepth === 0) {
			push(
				'exitOutsideBlock',
				"'Exit For' can only appear inside a For loop.",
				exitPhraseSpan(stmt.span, first, target),
			);
		} else if (target && targetWord === 'do' && ctx.doDepth === 0) {
			push(
				'exitOutsideBlock',
				"'Exit Do' can only appear inside a Do loop.",
				exitPhraseSpan(stmt.span, first, target),
			);
		}
	}
}

function exitPhraseSpan(base: Span, first: VbaToken, target: VbaToken): Span {
	return { start: base.start + first.start, end: base.start + target.end };
}
