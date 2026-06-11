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
	absoluteSpan,
	activeModuleMembers,
	forEachStatement,
	isInactiveNode,
	physicalLineSpanAtOffset,
	statementTokens,
	statementTokensAfterLeadingLabel,
	tokenName,
	tokenText,
} from './walker';
import {
	declarationShapeEnvironmentFor,
	declaredShapeForSourceBinding,
	isKnownScalarType,
	normalizeType,
	type DeclaredValueShape,
	type SourceDeclaredShape,
} from './typeInference';
import type {
	BodyNode,
	ForBlockNode,
	ModuleNode,
	ProcedureNode,
	Span,
	StatementNode,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import {
	createConditionalActivityTracker,
	type ConditionalActivityTracker,
} from '../conditional/conditionalCompilation';
import {
	collectProcedureLabelDeclarations,
	collectProcedureLabelReferences,
	collectProcedureLabels,
} from '../flow/procedureLabels';
import type { ModuleSymbolKind } from '../symbols/symbolModel';
import { type BareIdentifierContext } from '../symbols/nameResolution';
import { type MemberCompletionContext } from '../completion/memberAccess';
import { resolveTypeName } from '../completion/typeCompletion';
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
import { scanConditionalCompilationBranchOrder } from './rules/shared';
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
