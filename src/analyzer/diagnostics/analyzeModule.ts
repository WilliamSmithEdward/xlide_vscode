// Active VBA diagnostics engine (MS-VBAL Phase 5).
//
// `analyzeModule` runs the high-confidence semantic rules from the rule
// catalogue over one module's source and returns offset-based diagnostics. It
// is pure (no `vscode`): the editor layer converts spans to ranges and severity
// names to the VS Code enum. Structural block-balance checking stays in
// `src/vbaStructuralAnalysis.ts` (analyzeVbaStructure); this engine adds the semantic rules on
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

import { tokenize } from '../lexer/tokenize';
import type { VbaToken } from '../lexer/tokenKinds';
import { isReservedIdentifier } from '../lexer/keywordTable';
import { VBA_IDENTIFIER_NAME_RE } from '../../vbaStructuralAnalysis';
import {
	getHostMembers,
	resolveHostAlias,
	resolveHostConstant,
	resolveHostGlobal,
} from '../host/hostModel';
import type { HostObjectModel } from '../host/excelObjectModel';
import {
	resolveRuntimeConstant,
	resolveRuntimeFunction,
	resolveRuntimeObject,
	runtimeAllowsExplicitCall,
	type VbaRuntimeFunction,
} from '../runtime/vbaRuntime';
import {
	bareCallStatementTarget as callStatementTarget,
	explicitCallStatementArgumentWithoutParens,
	explicitCallStatementTarget,
	standaloneEmptyParenthesizedCallStatement,
} from '../call/callContext';
import type {
	BodyNode,
	EnumNode,
	ModuleMember,
	ModuleNode,
	ParameterNode,
	ProcedureNode,
	Span,
	StatementNode,
	TypeFieldNode,
	VariableDeclNode,
	VariableGroupNode,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { parseFixedLengthStringType } from '../parser/fixedLengthString';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import {
	conditionalCompilerConstants,
	createConditionalActivityTracker,
	type ConditionalActivityTracker,
	type ConditionalCompilationEnvironment,
} from '../conditional/conditionalCompilation';
import {
	collectProcedureLabelReferences,
	collectProcedureLabels,
} from '../flow/procedureLabels';
import type {
	ModuleSymbolKind,
	VbaProjectClassMembers,
	VbaProcedureSignature,
	VbaSymbol,
} from '../symbols/symbolModel';
import {
	isBareCallableKind,
	isProcedureKind,
	procedureParamsFromSymbol,
	qualifiedProcedureKey,
} from '../symbols/symbolModel';
import {
	resolveMemberCompletions,
	resolveMemberSurfaceAt,
	type MemberCompletion,
	type MemberCompletionContext,
} from '../completion/memberAccess';
import {
	isCreatableTypeCompletion,
	resolveTypeName,
	type ProjectTypeName,
	type TypeCompletionKind,
} from '../completion/typeCompletion';
import {
	eventHandlerDocumentTypeForContext,
	eventHandlerProcedureForName,
	type EventHandlerDocumentType,
} from '../completion/eventHandlers';
import {
	collectTypeNameReferences,
	typeReferenceLookupName,
	type TypeNameReferenceKind,
} from '../semantic/typeSemanticTokens';
import {
	DIAGNOSTIC_RULES,
	DiagnosticRuleName,
	DiagnosticSeverity,
	DiagnosticSeverityOverride,
	normalizeDiagnosticSeverityOverride,
} from './ruleMetadata';

/** A single diagnostic produced by the analyzer (offset-based). */
export interface VbaDiagnostic {
	/** Stable rule code (see DIAGNOSTIC_RULES). */
	code: string;
	/** Human-readable message. */
	message: string;
	/** Effective severity (after any user override). */
	severity: DiagnosticSeverity;
	/** Source span in UTF-16 offsets. */
	span: Span;
	/** MS-VBAL (or other) reference for the rule, when known. */
	specReference?: string;
	/** Optional structured data for deterministic editor actions. */
	data?: VbaDiagnosticData;
}

export interface VbaMissingRequiredArgumentPlaceholderData {
	parameterName: string;
	edit: {
		span: Span;
		newText: string;
	};
}

export interface VbaCreateProcedureStubData {
	procedureName: string;
	edit: {
		span: Span;
		newText: string;
	};
}

export interface VbaDiagnosticData {
	missingRequiredArgumentPlaceholder?: VbaMissingRequiredArgumentPlaceholderData;
	createProcedureStub?: VbaCreateProcedureStubData;
}

/** Per-rule severity overrides keyed by stable diagnostic code; `'off'` disables an allowed rule. */
export type DiagnosticSeverityOverrides = Partial<Record<string, DiagnosticSeverityOverride>>;

/** Inputs for {@link analyzeModule}. */
export interface AnalyzeModuleOptions {
	/** VB component name (used only for symbol container labels). */
	moduleName?: string;
	/** Workbook-project role of the module. */
	moduleKind?: ModuleSymbolKind;
	/** Workbook/document subtype for Excel document modules when known. */
	documentType?: EventHandlerDocumentType;
	/** Optional per-rule severity overrides keyed by stable diagnostic code. */
	severityOverrides?: DiagnosticSeverityOverrides;
	/**
	 * Lowercased procedure names callable as bare identifiers from the current
	 * module (from ProjectIndex.visibleProcedureNames). Required for the
	 * unknown-call statement rule; when omitted, that cross-module rule does not
	 * run so single-module analysis never false-positives on another module.
	 */
	knownProcedures?: ReadonlySet<string>;
	/**
	 * Lowercased bare identifiers visible from the current module (from
	 * ProjectIndex.visibleIdentifierNames). Required for the Option Explicit
	 * undeclared-assignment rule so single-module analysis never guesses about
	 * exported globals in other modules.
	 */
	knownIdentifiers?: ReadonlySet<string>;
	/**
	 * Exported Sub/Function/Declare signatures across the workbook project, grouped by
	 * lowercased procedure name. When omitted, type and arity validation remain
	 * single-module only.
	 */
	projectProcedures?: ReadonlyMap<string, readonly VbaProcedureSignature[]>;
	/** Source-declared workbook object members and UDT fields visible to this module. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/** Source-declared workbook type names visible to this module. */
	projectTypes?: readonly ProjectTypeName[];
	/**
	 * Source-backed module-level symbols visible as bare identifiers from this
	 * module. Used by call-target diagnostics to distinguish "not found" from
	 * "found, but not callable" across modules.
	 */
	projectVisibleSymbols?: readonly VbaSymbol[];
	/** Lowercased visible declaration names known not to be type names. */
	knownNonTypeNames?: ReadonlySet<string>;
	/**
	 * Raw integer constant expressions exported from other visible project modules.
	 * Used as a conservative base for deterministic runtime-value diagnostics.
	 */
	projectIntegerConstants?: ReadonlyMap<string, string | undefined>;
	/** Host object model metadata. Defaults to Excel's curated non-exhaustive model. */
	hostModel?: HostObjectModel;
	/**
	 * Conditional-compilation constants for deterministic branch filtering. Branches
	 * that remain unknown are still analyzed; only proven-inactive code is skipped.
	 */
	conditionalCompilation?: ConditionalCompilationEnvironment;
}

/** Counts double-quote characters; an odd count means the string is unterminated. */
function countQuotes(text: string): number {
	let n = 0;
	for (const ch of text) {
		if (ch === '"') {
			n++;
		}
	}
	return n;
}

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

export function incompleteMemberAccessEditSpan(
	source: string,
	offset: number,
): Span | undefined {
	const line = physicalLineSpanAtOffset(source, offset);
	const statement = activeStatementSpanOnLine(source, line, offset);
	return incompleteMemberAccess(source, statement, { includeLeadingDot: true })?.span;
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

	const mod = parseModule(source);
	const symbols = buildModuleSymbols(moduleName, moduleKind, source, {
		conditionalCompilation: opts.conditionalCompilation,
	});
	const activity = createConditionalActivityTracker(mod, opts.conditionalCompilation);
	const memberCtx = diagnosticMemberCompletionContext(opts);

	checkUnterminatedStrings(source, push);
	checkInvalidLineContinuations(source, push);
	checkDuplicateProcedures(symbols.root.children ?? [], push);
	checkDuplicateDeclarations(symbols.root.children ?? [], push);
	checkDuplicateModuleMembers(symbols.root.children ?? [], push);
	checkConstAssignment(source, mod, symbols, activity, push);
	checkOptionExplicit(source, mod, activity, push);
	checkUndeclaredVariables(
		source,
		mod,
		symbols,
		activity,
		opts.knownIdentifiers,
		opts.projectProcedures,
		opts.projectClassMembers,
		push,
	);
	checkOptionPlacement(source, mod, activity, push);
	checkProcedureHeader(source, mod, activity, push);
	checkInvalidIdentifierStarts(source, mod, activity, push);
	checkModuleDeclarationsInProcedureBodies(source, mod, activity, push);
	checkModuleDeclarationsAfterProcedures(source, mod, activity, push);
	checkReservedDeclarationNames(source, mod, activity, push);
	checkPropertySetterValueParameters(source, mod, activity, push);
	checkParameterOrder(source, mod, activity, push);
	checkParameterDefaultValues(source, mod, activity, push);
	checkUnbalancedParens(source, push);
	checkInvalidExpressionSyntax(source, mod, symbols, activity, push);
	checkDivisionByZeroExpressions(source, mod, opts.projectIntegerConstants, activity, push);
	checkDimInitializer(source, mod, activity, push);
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
	checkCallParens(source, mod, symbols, opts.projectProcedures, memberCtx, activity, push);
	checkExpressionCallParens(source, mod, symbols, opts.projectProcedures, activity, push);
	checkSetAssignments(source, mod, symbols, memberCtx, activity, push);
	checkExitStatements(source, mod, activity, push);
	checkUndefinedLabels(source, mod, activity, push);
	checkStatementContext(source, mod, activity, push);
	checkScalarMemberAccess(source, mod, symbols, activity, push);
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
		memberCtx,
		activity,
		push,
	);
	checkArgumentTypes(
		source,
		mod,
		symbols,
		opts.projectProcedures,
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
		activity,
		push,
	);
	checkAssignmentTypes(source, mod, symbols, memberCtx, activity, push);
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

type PushFn = (
	rule: DiagnosticRuleName,
	message: string,
	span: Span,
	data?: VbaDiagnosticData,
) => void;

function isInactiveNode(
	activity: ConditionalActivityTracker | undefined,
	node: { span: Span },
): boolean {
	return activity?.isInactive(node.span) ?? false;
}

function activeModuleMembers(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): readonly ModuleMember[] {
	if (!activity) {
		return mod.members;
	}
	return mod.members.filter((member) => !isInactiveNode(activity, member));
}

function diagnosticMemberCompletionContext(
	opts: AnalyzeModuleOptions,
): MemberCompletionContext {
	const ctx: MemberCompletionContext = {
		projectClassMembers: opts.projectClassMembers,
		allowSetAssignmentRefinement: false,
		model: opts.hostModel,
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

function isObjectModuleKind(moduleKind: ModuleSymbolKind | undefined): boolean {
	return moduleKind === 'class' || moduleKind === 'document' || moduleKind === 'userform';
}

/** Rule: a string literal with an odd number of quotes is never closed. */
function checkUnterminatedStrings(source: string, push: PushFn): void {
	for (const tok of tokenize(source)) {
		if (tok.kind === 'stringLiteral' && countQuotes(tok.rawText) % 2 === 1) {
			push(
				'unterminatedString',
				'Unterminated string literal.',
				{ start: tok.start, end: tok.end },
			);
		}
	}
}

/**
 * Rule: VBA line-continuation trivia is strictly `1*WSC "_" line-terminator`.
 * A likely continuation underscore with trailing text/comment, or without the
 * required whitespace before it, is a settled compile-time syntax error.
 */
function checkInvalidLineContinuations(source: string, push: PushFn): void {
	let lineStart = 0;
	while (lineStart < source.length) {
		let lineEnd = lineStart;
		while (lineEnd < source.length && source[lineEnd] !== '\r' && source[lineEnd] !== '\n') {
			lineEnd++;
		}
		checkInvalidLineContinuationOnLine(source, lineStart, lineEnd, push);
		if (lineEnd >= source.length) {
			break;
		}
		lineStart = source[lineEnd] === '\r' && source[lineEnd + 1] === '\n'
			? lineEnd + 2
			: lineEnd + 1;
	}
}

function checkInvalidLineContinuationOnLine(
	source: string,
	lineStart: number,
	lineEnd: number,
	push: PushFn,
): void {
	const commentStart = physicalLineCommentStart(source, lineStart, lineEnd) ?? lineEnd;
	const codeLast = lastNonWscOffset(source, lineStart, commentStart);
	if (codeLast === undefined) {
		return;
	}
	const visibleLineEnd = lastNonWscOffset(source, lineStart, lineEnd);
	const spanEnd = visibleLineEnd === undefined ? lineEnd : visibleLineEnd + 1;

	for (const underscore of underscoresOutsideStrings(source, lineStart, commentStart)) {
		const prev = source[underscore - 1];
		const next = source[underscore + 1];
		const prevIsWsc = underscore > lineStart && isVbaWsc(prev);
		const nextStartsIdentifier = next !== undefined && isIdentifierPartChar(next);
		const hasTrailingText = firstNonWscOffset(source, underscore + 1, lineEnd) !== undefined;

		if (prevIsWsc && hasTrailingText && !nextStartsIdentifier) {
			push(
				'invalidLineContinuation',
				"Line continuation '_' must be the final non-whitespace character on the physical line.",
				{ start: underscore, end: Math.max(underscore + 1, spanEnd) },
			);
			return;
		}

		if (
			underscore === codeLast &&
			lineEnd < source.length &&
			!prevIsWsc &&
			!isIdentifierPartChar(prev)
		) {
			push(
				'invalidLineContinuation',
				"Line continuation '_' must be preceded by whitespace.",
				{ start: underscore, end: underscore + 1 },
			);
			return;
		}
	}
}

function physicalLineCommentStart(
	source: string,
	lineStart: number,
	lineEnd: number,
): number | undefined {
	let inString = false;
	let statementStart = true;
	for (let i = lineStart; i < lineEnd; i++) {
		const ch = source[i];
		if (inString) {
			if (ch === '"') {
				if (source[i + 1] === '"') {
					i++;
				} else {
					inString = false;
				}
			}
			continue;
		}
		if (ch === '"') {
			inString = true;
			statementStart = false;
			continue;
		}
		if (ch === "'") {
			return i;
		}
		if (isVbaWsc(ch)) {
			continue;
		}
		if (ch === ':') {
			statementStart = true;
			continue;
		}
		if (statementStart && startsRemComment(source, i, lineEnd)) {
			return i;
		}
		statementStart = false;
	}
	return undefined;
}

function underscoresOutsideStrings(
	source: string,
	start: number,
	end: number,
): number[] {
	const offsets: number[] = [];
	let inString = false;
	for (let i = start; i < end; i++) {
		const ch = source[i];
		if (inString) {
			if (ch === '"') {
				if (source[i + 1] === '"') {
					i++;
				} else {
					inString = false;
				}
			}
			continue;
		}
		if (ch === '"') {
			inString = true;
			continue;
		}
		if (ch === '_') {
			offsets.push(i);
		}
	}
	return offsets;
}

function startsRemComment(source: string, offset: number, end: number): boolean {
	return offset + 3 <= end &&
		source.slice(offset, offset + 3).toLowerCase() === 'rem' &&
		!isIdentifierPartChar(source[offset + 3]);
}

function firstNonWscOffset(
	source: string,
	start: number,
	end: number,
): number | undefined {
	for (let i = start; i < end; i++) {
		if (!isVbaWsc(source[i])) {
			return i;
		}
	}
	return undefined;
}

function lastNonWscOffset(
	source: string,
	start: number,
	end: number,
): number | undefined {
	for (let i = end - 1; i >= start; i--) {
		if (!isVbaWsc(source[i])) {
			return i;
		}
	}
	return undefined;
}

function isVbaWsc(ch: string | undefined): boolean {
	return ch === '\t' || ch === '\u0019' || ch === ' ' || ch === '\u3000';
}

function isIdentifierPartChar(ch: string | undefined): boolean {
	return ch !== undefined && /[A-Za-z0-9_]/.test(ch);
}

/**
 * Rule: a procedure name may name at most one Sub/Function, OR a set of distinct
 * Property accessors (one Get, one Let, one Set). Any other repeat is the VBA
 * "Ambiguous name detected" compile error.
 */
function checkDuplicateProcedures(members: VbaSymbol[], push: PushFn): void {
	const groups = new Map<string, VbaSymbol[]>();
	for (const sym of members) {
		if (!isProcedureKind(sym.kind)) {
			continue;
		}
		const key = sym.name.toLowerCase();
		(groups.get(key) ?? groups.set(key, []).get(key)!).push(sym);
	}

	for (const group of groups.values()) {
		if (group.length < 2) {
			continue;
		}
		let valueProcSeen = false;
		const accessorSeen = new Set<string>();
		for (const sym of group) {
			const isProperty =
				sym.kind === 'propertyGet' ||
				sym.kind === 'propertyLet' ||
				sym.kind === 'propertySet';
			let conflict = false;
			if (!isProperty) {
				conflict = valueProcSeen || accessorSeen.size > 0;
				valueProcSeen = true;
			} else {
				conflict = valueProcSeen || accessorSeen.has(sym.kind);
				accessorSeen.add(sym.kind);
			}
			if (conflict) {
				push(
					'duplicateProcedure',
					`Ambiguous name detected: '${sym.name}' is already declared in this module.`,
					sym.nameSpan,
				);
			}
		}
	}
}

/**
 * Rule: within one procedure, a name may be declared once across its parameters,
 * local Dim/Static variables, and local Const declarations. Repeats are the VBA
 * "Duplicate declaration in current scope" error. Procedure scope is flat in VBA
 * (no block scope), so locals from different branches still collide.
 */
function checkDuplicateDeclarations(members: VbaSymbol[], push: PushFn): void {
	for (const proc of members) {
		if (!isProcedureKind(proc.kind)) {
			continue;
		}
		const seen = new Set<string>();
		for (const child of proc.children ?? []) {
			if (
				child.kind !== 'parameter' &&
				child.kind !== 'localVariable' &&
				child.kind !== 'constant'
			) {
				continue;
			}
			const key = child.name.toLowerCase();
			if (seen.has(key)) {
				push(
					'duplicateDeclaration',
					`Duplicate declaration in current scope: '${child.name}'.`,
					child.nameSpan,
				);
			} else {
				seen.add(key);
			}
		}
	}
}

/** Rule: a module-level variable or constant declared more than once. */
function checkDuplicateModuleMembers(members: VbaSymbol[], push: PushFn): void {
	const seen = new Set<string>();
	for (const sym of members) {
		if (sym.kind !== 'moduleVariable' && sym.kind !== 'constant') {
			continue;
		}
		const key = sym.name.toLowerCase();
		if (seen.has(key)) {
			push(
				'duplicateModuleMember',
				`Duplicate declaration: '${sym.name}' is already declared at module level.`,
				sym.nameSpan,
			);
		} else {
			seen.add(key);
		}
	}
}

/**
 * Rule: assigning to a constant is illegal. High-confidence form only - the
 * left-hand side must be a bare identifier (no member access, no index) that
 * resolves to a Const declared at module level or in the enclosing procedure.
 */
function checkConstAssignment(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleConsts = new Set<string>();
	for (const sym of symbols.root.children ?? []) {
		if (sym.kind === 'constant') {
			moduleConsts.add(sym.name.toLowerCase());
		}
	}

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = (symbols.root.children ?? []).find(
			(s) => isProcedureKind(s.kind) && s.fullSpan.start === member.span.start,
		);
		const localConsts = new Set<string>();
		for (const child of procSym?.children ?? []) {
			if (child.kind === 'constant') {
				localConsts.add(child.name.toLowerCase());
			}
		}
		const inScope = (lower: string): boolean =>
			localConsts.has(lower) || moduleConsts.has(lower);
		forEachStatement(member.body, (stmt) => {
			const hit = bareAssignmentTarget(source, stmt.span);
			if (hit && inScope(hit.name.toLowerCase())) {
				push(
					'constAssignment',
					`Cannot assign to constant '${hit.name}'.`,
					hit.span,
				);
			}
		});
	}
}

/** Walks every StatementNode in a body, descending into nested blocks. */
function forEachStatement(
	body: BodyNode[],
	visit: (stmt: StatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			forEachStatement(node.body, visit, activity);
		}
	}
}

/**
 * If the statement spanning `span` is a simple assignment to a bare identifier
 * (`name = ...` or `Let name = ...`), returns that identifier and its span;
 * otherwise undefined. `Set` (object) assignments and any left-hand side with a
 * `.` or `(` are excluded so only true scalar-name assignments are considered.
 */
function bareAssignmentTarget(
	source: string,
	span: Span,
): { name: string; span: Span; valueTokens: VbaToken[] } | undefined {
	const toks = statementTokens(source, span);
	let i = firstExecutableTokenIndex(toks);
	// Skip an explicit `Let`; bail on `Set` (object assignment).
	if (toks[i] && toks[i].kind === 'keyword') {
		const kw = toks[i].rawText.toLowerCase();
		if (kw === 'set') {
			return undefined;
		}
		if (kw === 'let') {
			i++;
		}
	}
	const nameTok = toks[i];
	if (!nameTok || nameTok.kind !== 'identifier') {
		return undefined; // first token must be a plain identifier LHS
	}
	const next = toks[i + 1];
	if (!next || next.kind !== 'operator' || next.rawText !== '=') {
		return undefined; // not `name =` (excludes `.`, `(`, `<=`, `<>`, comparisons)
	}
	return {
		name: nameTok.rawText,
		span: { start: span.start + nameTok.start, end: span.start + nameTok.end },
		valueTokens: toks.slice(i + 2),
	};
}

function memberAssignmentTarget(
	source: string,
	span: Span,
): {
	member: string;
	label: string;
	memberSpan: Span;
	valueTokens: VbaToken[];
	usesSet: boolean;
} | undefined {
	const toks = statementTokens(source, span);
	let i = firstExecutableTokenIndex(toks);
	const usesSet = tokenText(toks[i]) === 'set';
	if (usesSet || tokenText(toks[i]) === 'let') {
		i++;
	}
	const eq = topLevelOperatorIndex(toks.slice(i), '=');
	if (eq < 0) {
		return undefined;
	}
	const equalsIndex = i + eq;
	const lhs = toks.slice(i, equalsIndex);
	if (lhs.length < 2) {
		return undefined;
	}
	const memberTok = lhs[lhs.length - 1];
	if (!tokenName(memberTok) || lhs[lhs.length - 2]?.rawText !== '.') {
		return undefined;
	}
	if (lhs.some((tok) => tok.kind === 'operator' && tok.rawText === '=')) {
		return undefined;
	}
	return {
		member: tokenName(memberTok)!,
		label: source
			.slice(span.start + lhs[0].start, span.start + memberTok.end)
			.trim(),
		memberSpan: {
			start: span.start + memberTok.start,
			end: span.start + memberTok.end,
		},
		valueTokens: toks.slice(equalsIndex + 1),
		usesSet,
	};
}

function resolveExactMemberCompletion(
	source: string,
	memberName: string,
	memberEndOffset: number,
	memberCtx: MemberCompletionContext,
): MemberCompletion | undefined {
	return resolveMemberCompletions(source, memberEndOffset, memberCtx).find(
		(member) => member.name.toLowerCase() === memberName.toLowerCase(),
	);
}

function checkMemberNotFound(
	source: string,
	mod: ModuleNode,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			for (const ref of memberAccessReferences(source, stmt.span)) {
				const surface = resolveExhaustiveMemberSurface(
					source,
					ref.dotEndOffset,
					memberCtx,
				);
				if (!surface) {
					continue;
				}
				if (
					surface.members.some(
						(candidate) =>
							candidate.name.toLowerCase() === ref.member.toLowerCase(),
					)
				) {
					continue;
				}
				push(
					'memberNotFound',
					`Method or data member not found: '${surface.owner}.${ref.member}'.`,
					ref.memberSpan,
				);
			}
		}, activity);
	}
}

function memberAccessReferences(
	source: string,
	span: Span,
): { member: string; memberSpan: Span; dotEndOffset: number }[] {
	const toks = statementTokens(source, span);
	const out: { member: string; memberSpan: Span; dotEndOffset: number }[] = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (toks[i].rawText !== '.') {
			continue;
		}
		const member = tokenName(toks[i + 1]);
		if (!member) {
			continue;
		}
		out.push({
			member,
			memberSpan: {
				start: span.start + toks[i + 1].start,
				end: span.start + toks[i + 1].end,
			},
			dotEndOffset: span.start + toks[i].end,
		});
	}
	return out;
}

function resolveExhaustiveMemberSurface(
	source: string,
	dotEndOffset: number,
	memberCtx: MemberCompletionContext,
): { owner: string; members: MemberCompletion[] } | undefined {
	const surface = resolveMemberSurfaceAt(source, dotEndOffset, memberCtx);
	if (!surface?.exhaustive) {
		return undefined;
	}
	return { owner: surface.owner, members: surface.members };
}

/**
 * Rule: a *call statement* whose callee is a bare (non-member) identifier - the
 * lone-identifier form `DoStartup`, the parenless-argument form `MsgBox "hi"` /
 * `Foo 1, 2`, or the explicit `Call DoWork` / `Call Foo(1, 2)` form - is a call
 * to a Sub/Function of that name. When the name resolves to nothing the VBE
 * raises "Sub or Function not defined".
 *
 * A name is considered resolved when it matches any project procedure, a name
 * declared in the current module (procedures, module variables/consts, types,
 * enums and their members, Declares), a parameter/local/const of the enclosing
 * procedure, a VBA runtime function/statement, or a host global / Application
 * member (Excel exposes Application's members in the global scope). The callee
 * detection ({@link callStatementTarget}) deliberately ignores assignments,
 * member calls, line labels, and the bare `Name(...)` indexed/implicit-member
 * form so those never produce a false positive.
 */
function checkUnknownCallStatement(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	knownProcedures: ReadonlySet<string>,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): void {
	// Names visible module-wide: every module-level declaration (including a
	// live, not-yet-saved procedure) plus enum member names.
	const moduleNames = new Set<string>();
	for (const sym of symbols.root.children ?? []) {
		moduleNames.add(sym.name.toLowerCase());
		if (sym.kind === 'enum') {
			for (const member of sym.children ?? []) {
				moduleNames.add(member.name.toLowerCase());
			}
		}
	}

	// Excel injects Application's members into the global scope, so a bare call
	// may legitimately bind to one of them (Calculate, Volatile, Evaluate, ...).
	const appType = resolveHostGlobal('Application');
	const appMembers = new Set(
		(appType ? getHostMembers(appType) : []).map((mm) => mm.name.toLowerCase()),
	);
	const projectNonCallableNames = projectVisibleNonCallableNames(projectVisibleSymbols);

	const isKnown = (name: string, locals: ReadonlySet<string>): boolean => {
		const lower = name.toLowerCase();
		return (
			knownProcedures.has(lower) ||
			moduleNames.has(lower) ||
			projectNonCallableNames.has(lower) ||
			locals.has(lower) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveRuntimeObject(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = (symbols.root.children ?? []).find(
			(s) => isProcedureKind(s.kind) && s.fullSpan.start === member.span.start,
		);
		const locals = new Set<string>();
		for (const child of procSym?.children ?? []) {
			locals.add(child.name.toLowerCase());
		}
		forEachStatement(member.body, (stmt) => {
			const hit = callStatementTarget(source, stmt.span);
			if (hit && !isKnown(hit.name, locals)) {
				const call = extractCall(source, stmt.span);
				push(
					'unknownCallStatement',
					`Sub or Function not defined: '${hit.name}'.`,
					hit.span,
					call && call.nameSpan.start === hit.span.start && call.nameSpan.end === hit.span.end
						? createProcedureStubData(source, call)
						: undefined,
				);
			}
		}, activity);
	}
}

function createProcedureStubData(
	source: string,
	call: CallArguments,
): VbaDiagnosticData | undefined {
	if (!isGeneratedStubIdentifier(call.name)) {
		return undefined;
	}
	const params = generatedStubParameters(call);
	if (!params) {
		return undefined;
	}
	const eol = detectSourceEol(source);
	const leading = source.length === 0
		? ''
		: `${source.endsWith('\n') || source.endsWith('\r') ? '' : eol}${endsWithBlankPhysicalLine(source) ? '' : eol}`;
	const text = `${leading}Private Sub ${call.name}(${params.join(', ')})${eol}End Sub${eol}`;
	return {
		createProcedureStub: {
			procedureName: call.name,
			edit: {
				span: { start: source.length, end: source.length },
				newText: text,
			},
		},
	};
}

function generatedStubParameters(call: CallArguments): string[] | undefined {
	if (call.slots.some((slot) => slot.length === 0)) {
		return undefined;
	}
	const named = call.slots.map((slot) => isNamedSlot(slot));
	if (named.some(Boolean) && !named.every(Boolean)) {
		return undefined;
	}
	const used = new Set<string>();
	const params: string[] = [];
	for (let i = 0; i < call.slots.length; i++) {
		const name = named[i]
			? generatedNamedArgumentParameterName(call.slots[i])
			: `arg${i + 1}`;
		if (!name || used.has(name.toLowerCase())) {
			return undefined;
		}
		used.add(name.toLowerCase());
		params.push(`ByVal ${name} As Variant`);
	}
	return params;
}

function generatedNamedArgumentParameterName(slot: VbaToken[]): string | undefined {
	const raw = slot[0]?.rawText;
	if (!raw || raw.startsWith('[')) {
		return undefined;
	}
	return isGeneratedStubIdentifier(raw) ? raw : undefined;
}

function isGeneratedStubIdentifier(name: string): boolean {
	return VBA_IDENTIFIER_NAME_RE.test(name) && !isReservedIdentifier(name);
}

function detectSourceEol(source: string): string {
	return source.includes('\r\n') ? '\r\n' : '\n';
}

function endsWithBlankPhysicalLine(source: string): boolean {
	return /(?:\r\n|\r|\n)[ \t]*(?:\r\n|\r|\n)$/.test(source);
}

/**
 * Rule: call statements must target a callable declaration. VBE Compile rejects
 * a bare non-callable statement (`testStr`), argument-bearing form
 * (`testStr "hello"`), and explicit `Call testStr` as call-shaped statements.
 */
function checkNonCallableCallStatement(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	knownProcedures: ReadonlySet<string> | undefined,
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	push: PushFn,
): void {
	const moduleNonCallables = moduleNonCallableSymbols(symbols);
	const projectNonCallables = projectNonCallableSymbols(
		projectVisibleSymbols,
		knownProcedures,
	);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = (symbols.root.children ?? []).find(
			(s) => isProcedureKind(s.kind) && s.fullSpan.start === member.span.start,
		);
		const localNonCallables = new Map<string, VbaSymbol>();
		for (const child of procSym?.children ?? []) {
			if (isNonCallableSymbol(child)) {
				localNonCallables.set(child.name.toLowerCase(), child);
			}
		}
		forEachStatement(member.body, (stmt) => {
			const call = extractCall(source, stmt.span);
			if (!call) {
				return;
			}
			const lower = call.name.toLowerCase();
			const target = localNonCallables.get(lower) ??
				moduleNonCallables.get(lower) ??
				projectNonCallables.get(lower);
			if (!target) {
				return;
			}
			if (callTargetFeedsMemberAccess(source, stmt.span, call)) {
				return;
			}
			push(
				'nonCallableCallStatement',
				`Cannot call '${call.name}' because it resolves to ${symbolKindLabel(target)}, not a Sub or Function.`,
				call.nameSpan,
			);
		}, activity);
	}
}

function callTargetFeedsMemberAccess(source: string, span: Span, call: CallArguments): boolean {
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	const relCalleeStart = call.nameSpan.start - span.start;
	const calleeIdx = toks.findIndex((t) => t.start === relCalleeStart);
	if (calleeIdx < 0 || toks[calleeIdx + 1]?.rawText !== '(') {
		return false;
	}
	const close = matchParenFrom(toks, calleeIdx + 1);
	return close >= 0 && toks[close + 1]?.rawText === '.';
}

function projectNonCallableSymbols(
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
	knownProcedures: ReadonlySet<string> | undefined,
): Map<string, VbaSymbol> {
	const out = new Map<string, VbaSymbol>();
	const ambiguous = new Set<string>();
	for (const sym of projectVisibleSymbols ?? []) {
		if (!isNonCallableSymbol(sym)) {
			continue;
		}
		const key = sym.name.toLowerCase();
		if (knownProcedures?.has(key) || ambiguous.has(key)) {
			continue;
		}
		if (out.has(key)) {
			out.delete(key);
			ambiguous.add(key);
			continue;
		}
		out.set(key, sym);
	}
	return out;
}

function projectVisibleNonCallableNames(
	projectVisibleSymbols: readonly VbaSymbol[] | undefined,
): Set<string> {
	const out = new Set<string>();
	for (const sym of projectVisibleSymbols ?? []) {
		if (isNonCallableSymbol(sym)) {
			out.add(sym.name.toLowerCase());
		}
	}
	return out;
}

function moduleNonCallableSymbols(symbols: ReturnType<typeof buildModuleSymbols>): Map<string, VbaSymbol> {
	const out = new Map<string, VbaSymbol>();
	const callableNames = new Set(
		(symbols.root.children ?? [])
			.filter((sym) => isProcedureKind(sym.kind) || sym.kind === 'declare')
			.map((sym) => sym.name.toLowerCase()),
	);
	for (const sym of symbols.root.children ?? []) {
		if (isNonCallableSymbol(sym) && !callableNames.has(sym.name.toLowerCase())) {
			out.set(sym.name.toLowerCase(), sym);
		}
		if (sym.kind === 'enum') {
			for (const child of sym.children ?? []) {
				if (!callableNames.has(child.name.toLowerCase())) {
					out.set(child.name.toLowerCase(), child);
				}
			}
		}
	}
	return out;
}

function isNonCallableSymbol(sym: VbaSymbol): boolean {
	return (
		sym.kind === 'parameter' ||
		sym.kind === 'localVariable' ||
		sym.kind === 'moduleVariable' ||
		sym.kind === 'constant' ||
		sym.kind === 'enum' ||
		sym.kind === 'enumMember' ||
		sym.kind === 'type'
	);
}

function symbolKindLabel(sym: VbaSymbol): string {
	switch (sym.kind) {
		case 'parameter':
			return 'a parameter';
		case 'localVariable':
			return 'a local variable';
		case 'moduleVariable':
			return 'a module variable';
		case 'constant':
			return 'a constant';
		case 'enum':
			return 'an enum type';
		case 'enumMember':
			return 'an enum member';
		case 'type':
			return 'a user-defined type';
		default:
			return 'a non-callable declaration';
	}
}

/** Access/storage modifiers that may lead a procedure declaration. */
const PROC_MODIFIERS = new Set([
	'public', 'private', 'friend', 'global', 'static',
]);

/**
 * Rule: a procedure header must be `[(modifiers)] Sub|Function|Property Get/Let/Set
 * Name [(params)] [As Type]`. Once the name is read, the only legal next token is
 * `(` (the parameter list) or, for a `Function`/`Property Get`, `As` (the return
 * type). Any other token - most commonly a second word, as in `Sub My Sub`, where
 * the name was meant to contain a space - is the VBE "Expected: (" compile error.
 * Property `Let`/`Set` and `Sub` have no return value, so an `As` right after the
 * name is rejected for them too.
 */
function checkProcedureHeader(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const headerStart = member.span.start;
		const nl = source.indexOf('\n', headerStart);
		const headerEnd = nl === -1 ? member.span.end : nl;
		const toks = tokenize(source.slice(headerStart, headerEnd)).filter(
			(t) => t.kind !== 'comment' && t.kind !== 'newline',
		);

		let i = 0;
		while (i < toks.length && PROC_MODIFIERS.has(toks[i].rawText.toLowerCase())) {
			i++;
		}
		const kw = toks[i]?.rawText.toLowerCase();
		let allowAs = false;
		if (kw === 'function') {
			allowAs = true;
			i++;
		} else if (kw === 'sub') {
			i++;
		} else if (kw === 'property') {
			i++;
			if (toks[i]?.rawText.toLowerCase() === 'get') {
				allowAs = true;
			}
			i++; // skip the accessor (Get/Let/Set)
		} else {
			continue; // not a recognised procedure header
		}

		const nameTok = toks[i];
		if (!nameTok) {
			continue; // malformed in a way the structural analyzer already reports
		}
		if (isDigitStartedToken(nameTok)) {
			continue; // invalid-identifier-start owns the precise declaration-name range
		}
		const next = toks[i + 1];
		if (!next) {
			continue; // `Sub Foo` with no parameter list is legal
		}
		const r = next.rawText;
		if (r === '(' || (allowAs && r.toLowerCase() === 'as')) {
			continue;
		}
		push(
			'invalidProcedureHeader',
			`Unexpected '${r}' after procedure name '${stripHeaderBrackets(nameTok.rawText)}'; a procedure name must be a single identifier.`,
			{ start: headerStart + next.start, end: headerStart + next.end },
		);
	}
}

/** Strips the surrounding `[ ]` from a bracketed identifier, if present. */
function stripHeaderBrackets(text: string): string {
	return text.startsWith('[') && text.endsWith(']')
		? text.slice(1, -1)
		: text;
}

function checkInvalidIdentifierStarts(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (kind: string, hit: InvalidIdentifierStartHit | undefined): void => {
		if (!hit) {
			return;
		}
		push(
			'invalidIdentifierStart',
			`Invalid ${kind} name '${hit.name}': identifiers cannot start with a digit.`,
			hit.span,
		);
	};

	const inspectVariableGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			report('variable', invalidDeclarationIdentifierStart(source, decl.span));
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectVariableGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			report('user-defined type', invalidTypeOrEnumIdentifierStart(source, member.span, 'type'));
			for (const field of member.fields) {
				report('type field', invalidDeclarationIdentifierStart(source, field.span));
			}
			continue;
		}
		if (member.kind === 'Enum') {
			report('enum', invalidTypeOrEnumIdentifierStart(source, member.span, 'enum'));
			for (const enumMember of member.members) {
				report('enum member', invalidDeclarationIdentifierStart(source, enumMember.span));
			}
			continue;
		}
		if (member.kind === 'Declare') {
			report('Declare procedure', invalidDeclareIdentifierStart(source, member.span));
			continue;
		}
		if (member.kind === 'ConditionalDirective') {
			report('conditional compiler constant', invalidConstDirectiveIdentifierStart(source, member.span));
			continue;
		}
		if (member.kind !== 'Procedure') {
			continue;
		}
		report('procedure', invalidProcedureIdentifierStart(source, member));
		for (const param of member.params) {
			report('parameter', invalidParameterIdentifierStart(source, param.span));
		}
		forEachVariableGroup(member.body, inspectVariableGroup, activity);
	}
}

interface InvalidIdentifierStartHit {
	name: string;
	span: Span;
}

function invalidDeclarationIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	return invalidDigitIdentifierAt(source, span, toks, 0);
}

function invalidParameterIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	let i = 0;
	while (isParameterModifier(toks[i])) {
		i++;
	}
	return invalidDigitIdentifierAt(source, span, toks, i);
}

function invalidProcedureIdentifierStart(
	source: string,
	proc: ProcedureNode,
): InvalidIdentifierStartHit | undefined {
	const header = firstLineSpan(source, proc.span);
	const toks = statementTokens(source, header);
	let i = 0;
	while (i < toks.length && PROC_MODIFIERS.has(tokenText(toks[i]))) {
		i++;
	}
	const head = tokenText(toks[i]);
	if (head === 'property') {
		i += 2;
	} else if (head === 'sub' || head === 'function') {
		i++;
	}
	return invalidDigitIdentifierAt(source, header, toks, i);
}

function invalidTypeOrEnumIdentifierStart(
	source: string,
	span: Span,
	keyword: 'type' | 'enum',
): InvalidIdentifierStartHit | undefined {
	const header = firstLineSpan(source, span);
	const toks = statementTokens(source, header);
	let i = 0;
	if (tokenText(toks[i]) === 'public' || tokenText(toks[i]) === 'private') {
		i++;
	}
	if (tokenText(toks[i]) === keyword) {
		i++;
	}
	return invalidDigitIdentifierAt(source, header, toks, i);
}

function invalidDeclareIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	const kindIndex = toks.findIndex(
		(tok) => tokenText(tok) === 'sub' || tokenText(tok) === 'function',
	);
	return invalidDigitIdentifierAt(source, span, toks, kindIndex + 1);
}

function invalidConstDirectiveIdentifierStart(
	source: string,
	span: Span,
): InvalidIdentifierStartHit | undefined {
	const toks = statementTokens(source, span);
	return tokenText(toks[1]) === 'const'
		? invalidDigitIdentifierAt(source, span, toks, 2)
		: undefined;
}

function invalidDigitIdentifierAt(
	source: string,
	base: Span,
	toks: readonly VbaToken[],
	index: number,
): InvalidIdentifierStartHit | undefined {
	const tok = toks[index];
	if (!tok || !isDigitStartedToken(tok)) {
		return undefined;
	}
	const start = base.start + tok.start;
	const end = invalidIdentifierTextEnd(source, start, base.end);
	return {
		name: source.slice(start, end),
		span: { start, end },
	};
}

function isDigitStartedToken(tok: VbaToken): boolean {
	return (tok.kind === 'integerLiteral' || tok.kind === 'floatLiteral') && /^\d/.test(tok.rawText);
}

function invalidIdentifierTextEnd(source: string, start: number, limit: number): number {
	let end = start;
	while (end < limit && isInvalidIdentifierTextChar(source[end])) {
		end++;
	}
	return end;
}

function isInvalidIdentifierTextChar(ch: string | undefined): boolean {
	return ch !== undefined && /[A-Za-z0-9_]/.test(ch);
}

function isParameterModifier(tok: VbaToken | undefined): boolean {
	switch (tokenText(tok)) {
		case 'optional':
		case 'byval':
		case 'byref':
		case 'paramarray':
			return true;
		default:
			return false;
	}
}

const PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS = new Set(['public', 'private', 'friend', 'global']);
const DEFTYPE_KEYWORDS = new Set([
	'defbool',
	'defbyte',
	'defcur',
	'defdate',
	'defdbl',
	'defdec',
	'defint',
	'deflng',
	'deflnglng',
	'deflngptr',
	'defobj',
	'defsng',
	'defstr',
	'defvar',
]);

function checkModuleDeclarationsInProcedureBodies(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const inspectStatement = (stmt: StatementNode): void => {
		const hit = moduleDeclarationStatementInProcedure(source, stmt.span);
		if (!hit) {
			return;
		}
		push(
			'moduleDeclarationInProcedure',
			`${hit.label} must appear in the module declarations section, not inside a procedure.`,
			hit.span,
		);
	};
	const inspectProcedureBody = (procedure: ProcedureNode): void => {
		let sawConditionalDirective = false;
		for (const node of procedure.body) {
			if (node.kind === 'ConditionalDirective') {
				sawConditionalDirective = true;
				continue;
			}
			if (isInactiveNode(activity, node)) {
				continue;
			}
			if (node.kind === 'Statement') {
				if (
					sawConditionalDirective &&
					isAlternativeProcedureHeaderStatement(source, node.span, procedure)
				) {
					continue;
				}
				inspectStatement(node);
				continue;
			}
			if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
				forEachBodyStatement((node as { body: BodyNode[] }).body, inspectStatement, activity);
			}
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			inspectProcedureBody(member);
		}
	}
}

/**
 * Rule: module declarations belong in the declaration section before the first
 * procedure. Multiple procedures may follow each other, but once an active
 * procedure appears, later active module declarations are misplaced.
 */
function checkModuleDeclarationsAfterProcedures(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let procedureSeen = false;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Procedure') {
			procedureSeen = true;
			continue;
		}
		if (!procedureSeen) {
			continue;
		}
		const hit = moduleDeclarationAfterProcedureHit(source, member);
		if (!hit) {
			continue;
		}
		push(
			'moduleDeclarationAfterProcedure',
			`${hit.label} belong in the module declarations section, before procedures.`,
			hit.span,
		);
	}
}

function moduleDeclarationAfterProcedureHit(
	source: string,
	member: ModuleMember,
): { label: string; span: Span } | undefined {
	switch (member.kind) {
		case 'Declare':
			return {
				label: 'Declare statements',
				span: keywordSpan(source, member.span, 'declare'),
			};
		case 'Event':
			return {
				label: 'Event declarations',
				span: keywordSpan(source, member.span, 'event'),
			};
		case 'VariableGroup':
			return {
				label: member.isConst ? 'Const declarations' : 'Module variable declarations',
				span: member.isConst
					? keywordSpan(source, member.span, 'const')
					: firstTokenSpan(source, member.span),
			};
		case 'Type':
			return {
				label: 'Type declarations',
				span: keywordSpan(source, member.span, 'type'),
			};
		case 'Enum':
			return {
				label: 'Enum declarations',
				span: keywordSpan(source, member.span, 'enum'),
			};
		default:
			return undefined;
	}
}

function isAlternativeProcedureHeaderStatement(
	source: string,
	span: Span,
	procedure: ProcedureNode,
): boolean {
	const toks = statementTokensAfterLeadingLabel(source, span);
	let i = leadingDeclarationModifierCount(toks);
	const head = tokenText(toks[i]);
	let kind: ProcedureNode['procKind'] | undefined;
	if (head === 'property') {
		const accessor = tokenText(toks[i + 1]);
		kind =
			accessor === 'get'
				? 'PropertyGet'
				: accessor === 'let'
					? 'PropertyLet'
					: accessor === 'set'
						? 'PropertySet'
						: undefined;
		i += 2;
	} else if (head === 'function') {
		kind = 'Function';
		i += 1;
	} else if (head === 'sub') {
		kind = 'Sub';
		i += 1;
	}
	const name = tokenName(toks[i]);
	return !!kind &&
		kind === procedure.procKind &&
		!!name &&
		name.toLowerCase() === procedure.name.toLowerCase();
}

function leadingDeclarationModifierCount(toks: readonly VbaToken[]): number {
	let i = 0;
	while (PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS.has(tokenText(toks[i]))) {
		i++;
	}
	return i;
}

function moduleDeclarationStatementInProcedure(
	source: string,
	span: Span,
): { label: string; span: Span } | undefined {
	const toks = statementTokensAfterLeadingLabel(source, span);
	const first = toks[0];
	const head = tokenText(first);
	if (!first) {
		return undefined;
	}
	if (head === 'option') {
		return { label: 'Option statements', span: absoluteSpan(span, first) };
	}
	if (head === 'attribute') {
		return { label: 'Attribute statements', span: absoluteSpan(span, first) };
	}
	if (DEFTYPE_KEYWORDS.has(head)) {
		return {
			label: `${first.canonicalText ?? first.rawText} statements`,
			span: absoluteSpan(span, first),
		};
	}
	if (PROCEDURE_BODY_MODULE_DECLARATION_MODIFIERS.has(head)) {
		return {
			label: `${first.canonicalText ?? first.rawText} declarations`,
			span: absoluteSpan(span, first),
		};
	}
	return undefined;
}

function checkReservedDeclarationNames(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const report = (kind: string, hit: NameTokenHit | undefined): void => {
		if (!hit || hit.bracketed || !isReservedIdentifier(hit.name)) {
			return;
		}
		if (kind === 'type field' && hit.name.toLowerCase() === 'type') {
			return;
		}
		push(
			'invalidDeclarationName',
			`Reserved VBA keyword '${hit.name}' cannot be used as a ${kind} name.`,
			hit.span,
		);
	};

	const inspectVariableGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			report('variable', declarationNameHit(source, decl.span, decl.name));
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectVariableGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			report('user-defined type', typeOrEnumNameHit(source, member.span, 'type'));
			for (const field of member.fields) {
				report('type field', declarationNameHit(source, field.span, field.name));
			}
			continue;
		}
		if (member.kind === 'Enum') {
			report('enum', typeOrEnumNameHit(source, member.span, 'enum'));
			for (const enumMember of member.members) {
				report('enum member', declarationNameHit(source, enumMember.span, enumMember.name));
			}
			continue;
		}
		if (member.kind === 'Declare') {
			report('Declare procedure', declareNameHit(source, member.span));
			continue;
		}
		if (member.kind !== 'Procedure') {
			continue;
		}
		report('procedure', procedureNameHit(source, member));
		for (const param of member.params) {
			report('parameter', declarationNameHit(source, param.span, param.name));
		}
		forEachVariableGroup(member.body, inspectVariableGroup, activity);
	}
}

/**
 * Rule: Property Let/Set setters receive the assigned value through the final
 * parameter. A setter with no parameters has no value slot and Property Set
 * value parameters must be object references, not known scalar values.
 */
function checkPropertySetterValueParameters(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (
			member.kind !== 'Procedure' ||
			(member.procKind !== 'PropertyLet' && member.procKind !== 'PropertySet')
		) {
			continue;
		}
		if (member.params.length > 0) {
			if (member.procKind === 'PropertySet') {
				const valueParam = member.params[member.params.length - 1];
				const normalized = normalizeType(valueParam.asType);
				if (normalized && isKnownScalarType(normalized)) {
					push(
						'propertySetScalarValue',
						`Property Set '${member.name}' final value parameter '${valueParam.name}' must be an object reference, but it is declared As ${valueParam.asType}.`,
						declaredNameSpan(source, valueParam.span, valueParam.name),
					);
				}
			}
			continue;
		}
		const label = member.procKind === 'PropertyLet' ? 'Property Let' : 'Property Set';
		push(
			'propertySetterMissingValue',
			`${label} '${member.name}' must include a final value parameter.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

interface NameTokenHit {
	name: string;
	span: Span;
	bracketed: boolean;
}

function declarationNameHit(
	source: string,
	span: Span,
	name: string,
): NameTokenHit | undefined {
	const lower = name.toLowerCase();
	for (const tok of statementTokens(source, span)) {
		const found = tokenName(tok);
		if (found?.toLowerCase() === lower) {
			return nameTokenHit(span, tok, found);
		}
	}
	return undefined;
}

function procedureNameHit(source: string, proc: ProcedureNode): NameTokenHit | undefined {
	const header = firstLineSpan(source, proc.span);
	const toks = statementTokens(source, header);
	let i = 0;
	while (i < toks.length && PROC_MODIFIERS.has(tokenText(toks[i]))) {
		i++;
	}
	const head = tokenText(toks[i]);
	if (head === 'property') {
		i += 2;
	} else if (head === 'sub' || head === 'function') {
		i++;
	}
	const tok = toks[i];
	const name = tok ? tokenName(tok) : undefined;
	return tok && name ? nameTokenHit(header, tok, name) : undefined;
}

function typeOrEnumNameHit(
	source: string,
	span: Span,
	keyword: 'type' | 'enum',
): NameTokenHit | undefined {
	const header = firstLineSpan(source, span);
	const toks = statementTokens(source, header);
	let i = 0;
	if (tokenText(toks[i]) === 'public' || tokenText(toks[i]) === 'private') {
		i++;
	}
	if (tokenText(toks[i]) === keyword) {
		i++;
	}
	const tok = toks[i];
	const name = tok ? tokenName(tok) : undefined;
	return tok && name ? nameTokenHit(header, tok, name) : undefined;
}

function declareNameHit(source: string, span: Span): NameTokenHit | undefined {
	const toks = statementTokens(source, span);
	const kindIndex = toks.findIndex(
		(tok) => tokenText(tok) === 'sub' || tokenText(tok) === 'function',
	);
	const tok = kindIndex >= 0 ? toks[kindIndex + 1] : undefined;
	const name = tok ? tokenName(tok) : undefined;
	return tok && name ? nameTokenHit(span, tok, name) : undefined;
}

function firstLineSpan(source: string, span: Span): Span {
	const nl = source.indexOf('\n', span.start);
	return {
		start: span.start,
		end: nl === -1 ? span.end : Math.min(nl, span.end),
	};
}

function nameTokenHit(base: Span, tok: VbaToken, name: string): NameTokenHit {
	return {
		name,
		span: absoluteSpan(base, tok),
		bracketed: tok.kind === 'bracketedIdentifier',
	};
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
	const toks = tokenize(source);
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

/** A resolved call statement: callee plus its top-level argument slots. */
interface CallArguments {
	/** Callee identifier text. */
	name: string;
	/** Optional module qualifier for `ModuleName.MemberName` calls. */
	qualifier?: string;
	/** Lowercased signature lookup key; defaults to lowercased `name`. */
	lookupKey?: string;
	/** Absolute span of the callee identifier. */
	nameSpan: Span;
	/** True for the explicit `Call name...` form. */
	explicitCall?: boolean;
	/**
	 * Top-level, comma-separated argument groups. An empty list means no
	 * arguments were supplied; an empty inner array is an omitted positional
	 * argument (`Foo 1, , 3`).
	 */
	slots: VbaToken[][];
	/** Absolute spans for each argument slot; empty slots use the separator span. */
	slotSpans?: Span[];
	/** Absolute offset of the statement slice the slot tokens are relative to. */
	sliceStart: number;
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
 */
function checkArgumentCount(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const sameModuleSignatures = sameModuleCallableSignatures(symbols);
	const projectSignatures = uniqueProjectTypeSignatures(projectProcedures);
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const sourceNames = sourceNameScopeFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
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
		}, activity);
	}
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

/**
 * If the statement spanning `span` is a bare call statement, returns the callee
 * and its top-level argument slots; otherwise undefined. Reuses
 * {@link callStatementTarget} for the safe call-detection gating, then peels off
 * the argument region (the parenless tail, or the contents of the `Call`
 * statement's parentheses).
 */
function extractCall(source: string, span: Span): CallArguments | undefined {
	const hit = callStatementTarget(source, span);
	if (!hit) {
		return undefined;
	}
	const sliceStart = span.start;
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	const startIndex = firstExecutableTokenIndex(toks);
	const relCalleeStart = hit.span.start - sliceStart;
	const calleeIdx = toks.findIndex((t) => t.start === relCalleeStart);
	if (calleeIdx < 0) {
		return undefined;
	}

	const explicitCall = tokenText(toks[startIndex]) === 'call';
	const next = toks[calleeIdx + 1];
	let argToks: VbaToken[];
	if (explicitCall) {
		if (next && next.kind === 'punctuation' && next.rawText === '(') {
			// Collect the tokens strictly inside the call's parentheses.
			let depth = 0;
			let closed = false;
			const inner: VbaToken[] = [];
			for (let k = calleeIdx + 1; k < toks.length; k++) {
				const t = toks[k];
				if (t.kind === 'punctuation' && t.rawText === '(') {
					depth++;
					if (depth === 1) {
						continue; // skip the opening paren itself
					}
				} else if (t.kind === 'punctuation' && t.rawText === ')') {
					depth--;
					if (depth === 0) {
						closed = true;
						break;
					}
				}
				if (depth >= 1) {
					inner.push(t);
				}
			}
			if (!closed) {
				return undefined; // unbalanced - the parentheses rule reports this
			}
			argToks = inner;
		} else {
			argToks = []; // `Call Foo` with no parameter list
		}
	} else {
		argToks = toks.slice(calleeIdx + 1);
	}

	const split = argToks.length === 0 ? emptyArgSplit() : splitArgSlots(argToks, sliceStart);
	return {
		name: hit.name,
		nameSpan: hit.span,
		explicitCall,
		slots: split.slots,
		slotSpans: split.spans,
		sliceStart,
	};
}

/**
 * Extracts a module-qualified call statement (`ModuleName.Procedure ...`) only
 * when the project signature map proves that `ModuleName.Procedure` is an
 * exported project procedure. This keeps host/object member calls out of the
 * arity/type validator.
 */
function extractQualifiedCall(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): CallArguments | undefined {
	const sliceStart = span.start;
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	if (toks.length === 0) {
		return undefined;
	}

	let qualifierIdx = firstExecutableTokenIndex(toks);
	const explicitCall = tokenText(toks[qualifierIdx]) === 'call';
	if (explicitCall) {
		qualifierIdx += 1;
	}
	const qualifier = tokenName(toks[qualifierIdx]);
	const dot = toks[qualifierIdx + 1];
	const callee = toks[qualifierIdx + 2];
	const name = callee ? tokenName(callee) : undefined;
	if (!qualifier || dot?.rawText !== '.' || !name) {
		return undefined;
	}
	const lookupKey = qualifiedProcedureKey(qualifier, name);
	if (!moduleSignatures.has(lookupKey)) {
		return undefined;
	}

	const next = toks[qualifierIdx + 3];
	let argToks: VbaToken[];
	if (explicitCall) {
		if (next && next.kind === 'punctuation' && next.rawText === '(') {
			let depth = 0;
			let closed = false;
			const inner: VbaToken[] = [];
			for (let k = qualifierIdx + 3; k < toks.length; k++) {
				const t = toks[k];
				if (t.kind === 'punctuation' && t.rawText === '(') {
					depth++;
					if (depth === 1) {
						continue;
					}
				} else if (t.kind === 'punctuation' && t.rawText === ')') {
					depth--;
					if (depth === 0) {
						closed = true;
						break;
					}
				}
				if (depth >= 1) {
					inner.push(t);
				}
			}
			if (!closed) {
				return undefined;
			}
			argToks = inner;
		} else {
			argToks = [];
		}
	} else {
		if (next?.rawText === '(') {
			return undefined; // expressionCalls handles parenthesized forms.
		}
		if (next) {
			const gap = source.slice(span.start + callee.end, span.start + next.start);
			if (!/\s/.test(gap)) {
				return undefined;
			}
		}
		argToks = toks.slice(qualifierIdx + 3);
	}

	let depth = 0;
	for (let k = qualifierIdx + 3; k < toks.length; k++) {
		const raw = toks[k].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (depth === 0 && raw === '=') {
			return undefined;
		}
	}

	const split = argToks.length === 0 ? emptyArgSplit() : splitArgSlots(argToks, sliceStart);
	return {
		name,
		qualifier,
		lookupKey,
		nameSpan: { start: span.start + callee.start, end: span.start + callee.end },
		explicitCall,
		slots: split.slots,
		slotSpans: split.spans,
		sliceStart,
	};
}

interface ArgSplit {
	slots: VbaToken[][];
	spans: Span[];
}

/** Splits an argument token run into top-level (depth-0) comma-separated slots. */
function splitArgSlots(toks: VbaToken[], sliceStart: number): ArgSplit {
	const slots: VbaToken[][] = [[]];
	const spans: Span[] = [];
	let depth = 0;
	let emptyMarker: VbaToken | undefined;
	const finishSlot = (nextSeparator?: VbaToken): void => {
		const slot = slots[slots.length - 1];
		spans.push(argumentSlotSpan(slot, emptyMarker, nextSeparator, sliceStart));
	};
	for (const t of toks) {
		if (t.kind === 'punctuation' && t.rawText === '(') {
			depth++;
		} else if (t.kind === 'punctuation' && t.rawText === ')') {
			depth--;
		}
		if (t.kind === 'punctuation' && t.rawText === ',' && depth === 0) {
			finishSlot(t);
			slots.push([]);
			emptyMarker = t;
		} else {
			slots[slots.length - 1].push(t);
			emptyMarker = undefined;
		}
	}
	finishSlot();
	return { slots, spans };
}

function emptyArgSplit(): ArgSplit {
	return { slots: [], spans: [] };
}

function argumentSlotSpan(
	slot: VbaToken[],
	emptyMarker: VbaToken | undefined,
	nextSeparator: VbaToken | undefined,
	sliceStart: number,
): Span {
	if (slot.length > 0) {
		return {
			start: sliceStart + slot[0].start,
			end: sliceStart + slot[slot.length - 1].end,
		};
	}
	if (emptyMarker) {
		return { start: sliceStart + emptyMarker.start, end: sliceStart + emptyMarker.end };
	}
	if (nextSeparator) {
		return { start: sliceStart + nextSeparator.start, end: sliceStart + nextSeparator.end };
	}
	return { start: sliceStart, end: sliceStart };
}

/** True if a slot is a named argument (`name := value`). */
function isNamedSlot(slot: VbaToken[]): boolean {
	return (
		slot.length >= 2 &&
		(slot[0].kind === 'identifier' || slot[0].kind === 'bracketedIdentifier') &&
		slot[1].kind === 'operator' &&
		slot[1].rawText === ':='
	);
}

/** Describes a procedure's acceptable argument-count range for a message. */
function describeArity(required: number, max: number): string {
	if (max === Infinity) {
		return `at least ${required} argument${required === 1 ? '' : 's'}`;
	}
	if (required === max) {
		return `${required} argument${required === 1 ? '' : 's'}`;
	}
	return `between ${required} and ${max} arguments`;
}

/**
 * Validates one call's argument list against a procedure's parameters. When the
 * call uses named arguments, each name is checked against the parameter names
 * and the positional count check is skipped (positional/named mixing is too
 * subtle to count safely); otherwise the supplied slot count is checked against
 * the required minimum and the maximum implied by `Optional`/`ParamArray`.
 */
function validateArity(
	source: string,
	sig: CallableTypeSignature,
	call: CallArguments,
	push: PushFn,
): void {
	const displayName = callDisplayName(sig, call);
	const params = sig.params;
	let required = params.length;
	for (let k = 0; k < params.length; k++) {
		if (params[k].optional || params[k].paramArray) {
			required = k;
			break;
		}
	}
	const hasParamArray = params.some((p) => p.paramArray);
	const max = hasParamArray ? Infinity : params.length;

	const named = call.slots.filter(isNamedSlot);
	if (named.length > 0) {
		const paramNames = new Set(
			params.map((p) => stripHeaderBrackets(p.name).toLowerCase()),
		);
		for (const slot of named) {
			const raw = stripHeaderBrackets(slot[0].rawText);
			if (!paramNames.has(raw.toLowerCase())) {
				push(
					'argumentCount',
					`Named argument not found: '${raw}' is not a parameter of '${displayName}'.`,
					{
						start: call.sliceStart + slot[0].start,
						end: call.sliceStart + slot[0].end,
					},
				);
			}
		}
		return; // positional count is not validated alongside named arguments
	}

	for (let i = 0; i < Math.min(call.slots.length, params.length); i++) {
		const param = params[i];
		if (call.slots[i].length === 0 && !param.optional && !param.paramArray) {
			const name = stripHeaderBrackets(param.name);
			const placeholder = omittedArgumentPlaceholderData(source, call, param, i);
			push(
				'argumentCount',
				`Argument not optional: '${name}' is required by '${displayName}'.`,
				call.slotSpans?.[i] ?? call.nameSpan,
				placeholder,
			);
		}
	}

	const n = call.slots.length;
	if (n < required || n > max) {
		const missingParam = n < required ? params[n] : undefined;
		const placeholder = missingParam
			? trailingMissingArgumentPlaceholderData(source, call, missingParam)
			: undefined;
		push(
			'argumentCount',
			`Wrong number of arguments to '${displayName}': expected ${describeArity(required, max)}, but got ${n}.`,
			call.nameSpan,
			placeholder,
		);
	}
}

function callDisplayName(sig: CallableTypeSignature, call: CallArguments): string {
	return call.qualifier ? `${call.qualifier}.${sig.name}` : sig.name;
}

function omittedArgumentPlaceholderData(
	source: string,
	call: CallArguments,
	param: CallableParamType,
	slotIndex: number,
): VbaDiagnosticData | undefined {
	const separator = call.slotSpans?.[slotIndex];
	if (!separator || source.slice(separator.start, separator.end) !== ',') {
		return undefined;
	}
	const parameterName = stripHeaderBrackets(param.name);
	const placeholder = placeholderNameForParameter(parameterName);
	return {
		missingRequiredArgumentPlaceholder: {
			parameterName,
			edit: {
				span: separatorWithFollowingHorizontalSpace(source, separator),
				newText: slotIndex === 0 ? `${placeholder}, ` : `, ${placeholder}`,
			},
		},
	};
}

function trailingMissingArgumentPlaceholderData(
	source: string,
	call: CallArguments,
	param: CallableParamType,
): VbaDiagnosticData | undefined {
	const parameterName = stripHeaderBrackets(param.name);
	const placeholder = placeholderNameForParameter(parameterName);
	if (call.slots.length === 0) {
		const innerSpan = emptyParenthesizedArgumentSpan(source, call);
		if (innerSpan) {
			return missingArgumentPlaceholderData(parameterName, innerSpan, placeholder);
		}
		const newText = call.explicitCall ? `(${placeholder})` : ` ${placeholder}`;
		return missingArgumentPlaceholderData(
			parameterName,
			{ start: call.nameSpan.end, end: call.nameSpan.end },
			newText,
		);
	}
	const lastSpan = call.slotSpans?.[call.slotSpans.length - 1];
	if (!lastSpan) {
		return undefined;
	}
	return missingArgumentPlaceholderData(
		parameterName,
		{ start: lastSpan.end, end: lastSpan.end },
		`, ${placeholder}`,
	);
}

function missingArgumentPlaceholderData(
	parameterName: string,
	span: Span,
	newText: string,
): VbaDiagnosticData {
	return {
		missingRequiredArgumentPlaceholder: {
			parameterName,
			edit: { span, newText },
		},
	};
}

function placeholderNameForParameter(name: string): string {
	const safe = name.replace(/[^A-Za-z0-9_]/g, '_').replace(/^[^A-Za-z_]+/, '');
	return `TODO_${safe || 'Argument'}`;
}

function separatorWithFollowingHorizontalSpace(source: string, separator: Span): Span {
	let end = separator.end;
	while (end < source.length && (source[end] === ' ' || source[end] === '\t')) {
		end++;
	}
	return { start: separator.start, end };
}

function emptyParenthesizedArgumentSpan(source: string, call: CallArguments): Span | undefined {
	let open = call.nameSpan.end;
	while (open < source.length && (source[open] === ' ' || source[open] === '\t')) {
		open++;
	}
	if (source[open] !== '(') {
		return undefined;
	}
	let close = open + 1;
	while (close < source.length && (source[close] === ' ' || source[close] === '\t')) {
		close++;
	}
	if (source[close] !== ')') {
		return undefined;
	}
	return { start: open + 1, end: close };
}

interface CallableParamType {
	name: string;
	type?: string;
	optional: boolean;
	paramArray: boolean;
	byRef?: boolean;
}

interface CallableTypeSignature {
	name: string;
	params: CallableParamType[];
	returnType?: string;
}

interface InferredArgumentType {
	type: string;
	label: string;
	span: Span;
	stringValue?: string;
	numericValue?: number;
	numericText?: string;
}

/**
 * Rule: when both a callable parameter type and an argument type are known, flag
 * high-confidence mismatches. This first slice is deliberately conservative:
 * unknowns and Variant are accepted, and VBA's normal coercions are allowed
 * unless a literal is clearly incompatible (for example `"blah"` for Currency).
 */
function checkArgumentTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
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
		const sourceNames = sourceNameScopeFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			for (const call of expressionCalls(source, stmt.span, moduleSignatures, sourceNames)) {
				validateArgumentTypes(call, env, moduleSignatures, sourceNames, source, memberCtx, push);
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
				);
			}
			const statementCall = extractCall(source, stmt.span);
			const qualifiedStatementCall = statementCall
				? undefined
				: extractQualifiedCall(source, stmt.span, moduleSignatures);
			const effectiveStatementCall = statementCall ?? qualifiedStatementCall;
			if (effectiveStatementCall) {
				validateArgumentTypes(effectiveStatementCall, env, moduleSignatures, sourceNames, source, memberCtx, push);
			}
		}, activity);
	}
}

interface RuntimeArgumentValueSpec {
	canonicalName: 'Left' | 'Right' | 'String' | 'Space' | 'Mid' | 'Replace' | 'InStr' | 'Chr' | 'ChrW';
	parameterName: string;
	argumentIndex: number;
	minimum?: number;
	maximum?: number;
	minimumSlotCount?: number;
	allowNamed?: boolean;
}

interface RuntimeArgumentValueHit {
	displayName: string;
	parameterName: string;
	value: number;
	span: Span;
}

/**
 * Rule: some runtime-library arguments have deterministic value bounds even
 * when the argument type itself is valid. This slice is VBE-oracle-backed for
 * integer bounds on selected string runtime functions, which compile but raise
 * Run-time error 5 when the value is outside the proven range.
 */
function checkRuntimeArgumentValues(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const projectConstants = resolveRawIntegerConstants(projectIntegerConstants ?? new Map(), new Map());
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity, projectConstants);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const procedureConstants = new Map(moduleConstants);
		collectBodyLiteralIntegerConstants(member.body, procedureConstants, activity);
		forEachStatement(member.body, (stmt) => {
			for (const hit of runtimeArgumentValueHits(source, stmt.span, moduleSignatures, env, procedureConstants)) {
				push(
					'runtimeArgumentValue',
					`Argument '${hit.parameterName}' of '${hit.displayName}' is ${hit.value}; this will raise Run-time error '5': Invalid procedure call or argument.`,
					hit.span,
				);
			}
		}, activity);
	}
}

function runtimeArgumentValueHits(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	env: ReadonlyMap<string, string>,
	constants: ReadonlyMap<string, number | undefined>,
): RuntimeArgumentValueHit[] {
	const toks = statementTokens(source, span);
	if (isDeclarationLikeStatement(toks)) {
		return [];
	}
	const hits: RuntimeArgumentValueHit[] = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const call = runtimeArgumentValueCallAt(toks, i, span, moduleSignatures, env);
		if (!call) {
			continue;
		}
		for (const spec of call.specs) {
			const slot = runtimeArgumentValueSlot(call.slots, spec);
			const literal = slot
				? integerArgumentOutsideBounds(source, slot, span.start, spec, constants)
				: undefined;
			if (!literal) {
				continue;
			}
			hits.push({
				displayName: call.displayName,
				parameterName: spec.parameterName,
				value: literal.value,
				span: literal.span,
			});
		}
	}
	return hits;
}

function runtimeArgumentValueCallAt(
	toks: readonly VbaToken[],
	index: number,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	env: ReadonlyMap<string, string>,
): {
	displayName: string;
	specs: readonly RuntimeArgumentValueSpec[];
	slots: VbaToken[][];
} | undefined {
	const name = tokenName(toks[index]);
	if (!name) {
		return undefined;
	}
	const qualifier = index >= 2 && toks[index - 1].rawText === '.'
		? tokenName(toks[index - 2])
		: undefined;
	if (index > 0 && toks[index - 1].rawText === '.' && !qualifier) {
		return undefined;
	}
	if (qualifier && qualifier.toLowerCase() !== 'vba') {
		return undefined;
	}

	let parenIndex = index + 1;
	let suffix = '';
	if (isRuntimeStringFunctionSuffix(toks[parenIndex])) {
		suffix = toks[parenIndex].rawText;
		parenIndex++;
	}
	if (toks[parenIndex]?.rawText !== '(') {
		return undefined;
	}

	const specs = runtimeArgumentValueSpecs(name);
	if (specs.length === 0) {
		return undefined;
	}
	if (suffix && !runtimeArgumentValueAllowsStringSuffix(specs[0].canonicalName)) {
		return undefined;
	}
	const lower = specs[0].canonicalName.toLowerCase();
	if (!qualifier && (moduleSignatures.has(lower) || env.has(lower))) {
		return undefined;
	}

	const close = matchParenFrom(toks, parenIndex);
	if (close < 0) {
		return undefined;
	}
	const inner = toks.slice(parenIndex + 1, close);
	const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
	return {
		displayName: `${specs[0].canonicalName}${suffix}`,
		specs,
		slots: split.slots,
	};
}

function runtimeArgumentValueSpecs(name: string): readonly RuntimeArgumentValueSpec[] {
	switch (name.toLowerCase()) {
		case 'left':
			return [{ canonicalName: 'Left', parameterName: 'Length', argumentIndex: 1, minimum: 0 }];
		case 'right':
			return [{ canonicalName: 'Right', parameterName: 'Length', argumentIndex: 1, minimum: 0 }];
		case 'string':
			return [{ canonicalName: 'String', parameterName: 'Number', argumentIndex: 0, minimum: 0 }];
		case 'space':
			return [{ canonicalName: 'Space', parameterName: 'Number', argumentIndex: 0, minimum: 0 }];
		case 'mid':
			return [
				{ canonicalName: 'Mid', parameterName: 'Start', argumentIndex: 1, minimum: 1 },
				{ canonicalName: 'Mid', parameterName: 'Length', argumentIndex: 2, minimum: 0 },
			];
		case 'replace':
			return [
				{ canonicalName: 'Replace', parameterName: 'Start', argumentIndex: 3, minimum: 1 },
				{ canonicalName: 'Replace', parameterName: 'Count', argumentIndex: 4, minimum: -1 },
			];
		case 'instr':
			return [
				{
					canonicalName: 'InStr',
					parameterName: 'Start',
					argumentIndex: 0,
					minimum: 1,
					minimumSlotCount: 3,
					allowNamed: false,
				},
			];
		case 'chr':
			return [{ canonicalName: 'Chr', parameterName: 'CharCode', argumentIndex: 0, minimum: 0, maximum: 255 }];
		case 'chrw':
			return [{ canonicalName: 'ChrW', parameterName: 'CharCode', argumentIndex: 0, maximum: 65535 }];
		default:
			return [];
	}
}

function runtimeArgumentValueAllowsStringSuffix(name: RuntimeArgumentValueSpec['canonicalName']): boolean {
	return name === 'Left' || name === 'Right' || name === 'String' || name === 'Space' || name === 'Mid';
}

function runtimeArgumentValueSlot(
	slots: readonly VbaToken[][],
	spec: RuntimeArgumentValueSpec,
): VbaToken[] | undefined {
	if (spec.minimumSlotCount !== undefined && slots.length < spec.minimumSlotCount) {
		return undefined;
	}
	let positionalIndex = 0;
	for (const slot of slots) {
		const named = namedArgumentSlot(slot);
		if (named) {
			if (spec.allowNamed === false) {
				continue;
			}
			if (named.name.toLowerCase() === spec.parameterName.toLowerCase()) {
				return named.value;
			}
			continue;
		}
		if (positionalIndex === spec.argumentIndex) {
			return slot;
		}
		positionalIndex++;
	}
	return undefined;
}

function integerArgumentOutsideBounds(
	source: string,
	slot: readonly VbaToken[],
	sliceStart: number,
	spec: RuntimeArgumentValueSpec,
	constants: ReadonlyMap<string, number | undefined>,
): { value: number; span: Span } | undefined {
	const toks = unwrapOuterParens(
		slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline'),
	);
	if (toks.length === 0) {
		return undefined;
	}
	let sign = 1;
	let literal = toks[0];
	let start = literal?.start;
	let literalValue: number | undefined;
	const signedLiteral = toks.length === 2 && (toks[0].rawText === '-' || toks[0].rawText === '+');
	if (signedLiteral) {
		sign = toks[0].rawText === '-' ? -1 : 1;
		literal = toks[1];
		start = toks[0].start;
	}
	if (literal?.kind === 'integerLiteral' && start !== undefined && (toks.length === 1 || signedLiteral)) {
		const rawValue = parseVbaIntegerLiteral(literal.rawText);
		if (rawValue !== undefined) {
			literalValue = sign * rawValue;
		}
	}
	if (literalValue !== undefined) {
		if (integerArgumentValueInBounds(literalValue, spec)) {
			return undefined;
		}
		return {
			value: literalValue,
			span: { start: sliceStart + start!, end: sliceStart + literal.end },
		};
	}

	const expressionValue = evaluateIntegerConstantExpression(
		source.slice(sliceStart + toks[0].start, sliceStart + toks[toks.length - 1].end),
		constants,
	);
	if (expressionValue === undefined || integerArgumentValueInBounds(expressionValue, spec)) {
		return undefined;
	}
	return {
		value: expressionValue,
		span: { start: sliceStart + toks[0].start, end: sliceStart + toks[toks.length - 1].end },
	};
}

function integerArgumentValueInBounds(
	value: number,
	spec: RuntimeArgumentValueSpec,
): boolean {
	if (spec.minimum !== undefined && value < spec.minimum) {
		return false;
	}
	if (spec.maximum !== undefined && value > spec.maximum) {
		return false;
	}
	return true;
}

function isRuntimeStringFunctionSuffix(tok: VbaToken | undefined): boolean {
	return tok?.rawText === '$';
}

function isDeclarationLikeStatement(toks: readonly VbaToken[]): boolean {
	const first = tokenText(toks[0]);
	switch (first) {
		case 'dim':
		case 'static':
		case 'const':
		case 'private':
		case 'public':
		case 'friend':
		case 'declare':
		case 'sub':
		case 'function':
		case 'property':
		case 'type':
		case 'enum':
			return true;
		default:
			return false;
	}
}

function checkAssignmentTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			const assignment = bareAssignmentTarget(source, stmt.span);
			if (!assignment) {
				return;
			}
			const expected = env.get(assignment.name.toLowerCase());
			if (!expected) {
				return;
			}
			if (isKnownObjectAssignmentType(expected, memberCtx)) {
				push(
					'setRequired',
					`Object assignment to '${assignment.name}' requires Set because it is declared as ${expected}.`,
					assignment.span,
				);
				return;
			}
			const stringArithmetic = nonnumericStringArithmeticOperand(
				expected,
				assignment.valueTokens,
				stmt.span.start,
			);
			if (stringArithmetic) {
				push(
					'stringArithmeticCoercion',
					`Assignment to '${assignment.name}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
					stringArithmetic.span,
				);
				return;
			}
			const actual = inferArgumentType(
				assignment.valueTokens,
				stmt.span.start,
				env,
				moduleSignatures,
				sourceNames,
				source,
				memberCtx,
			);
			if (!actual) {
				return;
			}
			const reason = incompatibilityReason(expected, actual);
			if (!reason) {
				return;
			}
			push(
				'assignmentTypeMismatch',
				`Assignment to '${assignment.name}' expects ${expected}, but got ${actual.label}. ${reason}`,
				actual.span,
			);
		}, activity);
		checkMemberAssignmentTypes(
			source,
			member,
			env,
			moduleSignatures,
			sourceNames,
			memberCtx,
			activity,
			push,
		);
	}
}

/**
 * Rule: a Function/Property Get returns through its hidden return variable.
 * Falling through without assigning that variable is legal VBA, but it silently
 * returns the default value. XLIDE only surfaces this for untyped returns, where
 * the implicit Variant fallthrough is more likely to be accidental than an
 * intentional typed default value.
 */
function checkMissingReturnAssignments(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (
			member.kind !== 'Procedure' ||
			(member.procKind !== 'Function' && member.procKind !== 'PropertyGet')
		) {
			continue;
		}
		if (!member.closed) {
			continue;
		}
		if (member.returnType) {
			continue;
		}
		if (procedureHasReturnAssignment(source, member, activity, moduleSignatures)) {
			continue;
		}
		const procLabel = member.procKind === 'PropertyGet' ? 'Property Get' : 'Function';
		push(
			'missingReturnAssignment',
			`${procLabel} '${member.name}' has no return assignment; VBA will return the default value. Assign to '${member.name}' before exit if a value is intended.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function procedureHasReturnAssignment(
	source: string,
	proc: ProcedureNode,
	activity: ConditionalActivityTracker | undefined,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	const lower = proc.name.toLowerCase();
	let found = false;
	forEachStatement(proc.body, (stmt) => {
		if (found) {
			return;
		}
		const bare = bareAssignmentTarget(source, stmt.span);
		if (bare?.name.toLowerCase() === lower) {
			found = true;
			return;
		}
		const set = setAssignmentTarget(source, stmt.span);
		if (set?.name.toLowerCase() === lower) {
			found = true;
			return;
		}
		const call = extractCall(source, stmt.span);
		const qualifiedCall = call
			? undefined
			: extractQualifiedCall(source, stmt.span, moduleSignatures);
		const effectiveCall = call ?? qualifiedCall;
		if (effectiveCall && callPassesNameToByRefParam(effectiveCall, lower, moduleSignatures)) {
			found = true;
		}
	}, activity);
	return found;
}

function callPassesNameToByRefParam(
	call: CallArguments,
	lowerName: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	const sig = callableSignatureForCall(call, moduleSignatures);
	if (!sig) {
		return false;
	}
	let positionalIndex = 0;
	for (const slot of call.slots) {
		const named = namedArgumentSlot(slot);
		let param: CallableParamType | undefined;
		let valueSlot = slot;
		if (named) {
			param = sig.params.find((p) => stripHeaderBrackets(p.name).toLowerCase() === named.name.toLowerCase());
			valueSlot = named.value;
		} else {
			param = sig.params[Math.min(positionalIndex, sig.params.length - 1)];
			positionalIndex++;
		}
		if (!param?.byRef || !singleSlotNameEquals(valueSlot, lowerName)) {
			continue;
		}
		return true;
	}
	return false;
}

function singleSlotNameEquals(slot: readonly VbaToken[], lowerName: string): boolean {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return toks.length === 1 && tokenName(toks[0])?.toLowerCase() === lowerName;
}

function checkMemberAssignmentTypes(
	source: string,
	member: ProcedureNode,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	if (!memberCtx.projectClassMembers || memberCtx.projectClassMembers.length === 0) {
		return;
	}
	forEachStatement(member.body, (stmt) => {
		const assignment = memberAssignmentTarget(source, stmt.span);
		if (!assignment) {
			return;
		}
		const target = resolveExactMemberCompletion(
			source,
			assignment.member,
			assignment.memberSpan.end,
			memberCtx,
		);
		if (!target || target.writable === undefined) {
			return;
		}
		if (target.writable === false) {
			push(
				'readonlyMemberAssignment',
				`Cannot assign to read-only property '${assignment.label}'.`,
				assignment.memberSpan,
			);
			return;
		}
		const expected = target.writeType ?? target.returns;
		if (assignment.usesSet) {
			if (expected && isKnownScalarType(normalizeType(expected) ?? '')) {
				push(
					'setRequiresObject',
					`Set assignment requires an object-valued target, but '${assignment.label}' expects ${expected}.`,
					assignment.memberSpan,
				);
				return;
			}
			const actual = inferArgumentType(
				assignment.valueTokens,
				stmt.span.start,
				env,
				moduleSignatures,
				sourceNames,
				source,
				memberCtx,
			);
			const reason = objectAssignmentIncompatibilityReason(
				expected,
				actual,
				memberCtx,
			);
			if (reason) {
				push(
					'assignmentObjectTypeMismatch',
					`Object assignment to '${assignment.label}' expects ${expected}, but got ${actual?.label}. ${reason}`,
					actual?.span ?? assignment.memberSpan,
				);
			}
			return;
		}
		if (isKnownObjectAssignmentType(expected, memberCtx)) {
			push(
				'setRequired',
				`Object assignment to '${assignment.label}' requires Set because it expects ${expected}.`,
				assignment.memberSpan,
			);
			return;
		}
		if (!expected || normalizeType(expected) === 'object') {
			return;
		}
		const stringArithmetic = nonnumericStringArithmeticOperand(
			expected,
			assignment.valueTokens,
			stmt.span.start,
		);
		if (stringArithmetic) {
			push(
				'stringArithmeticCoercion',
				`Assignment to '${assignment.label}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
				stringArithmetic.span,
			);
			return;
		}
		const actual = inferArgumentType(
			assignment.valueTokens,
			stmt.span.start,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
		);
		if (!actual) {
			return;
		}
		const reason = incompatibilityReason(expected, actual);
		if (!reason) {
			return;
		}
		push(
			'assignmentTypeMismatch',
			`Assignment to '${assignment.label}' expects ${expected}, but got ${actual.label}. ${reason}`,
			actual.span,
		);
	}, activity);
}

function checkSetAssignments(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = buildModuleTypeSignatures(symbols);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		const sourceNames = sourceNameScopeFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			const target = setAssignmentTarget(source, stmt.span);
			if (!target) {
				return;
			}
			const expected = env.get(target.name.toLowerCase());
			const targetType = normalizeType(expected);
			if (!targetType || !isKnownScalarType(targetType)) {
				if (!isKnownObjectAssignmentType(expected, memberCtx)) {
					return;
				}
				const actual = inferArgumentType(
					target.valueTokens,
					stmt.span.start,
					env,
					moduleSignatures,
					sourceNames,
					source,
					memberCtx,
				);
				const reason = objectAssignmentIncompatibilityReason(
					expected,
					actual,
					memberCtx,
				);
				if (reason) {
					push(
						'assignmentObjectTypeMismatch',
						`Object assignment to '${target.name}' expects ${expected}, but got ${actual?.label}. ${reason}`,
						actual?.span ?? target.span,
					);
				}
				return;
			}
			push(
				'setRequiresObject',
				`Set assignment requires an object variable, but '${target.name}' is declared as ${expected}.`,
				target.span,
			);
		}, activity);
	}
}

function checkScalarMemberAccess(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
		for (const member of activeModuleMembers(mod, activity)) {
			if (member.kind !== 'Procedure') {
				continue;
			}
			const env = typeEnvironmentFor(symbols, member);
			forEachStatement(member.body, (stmt) => {
				for (const hit of scalarMemberAccesses(source, stmt.span, env)) {
					push(
						'scalarMemberAccess',
						`Member access on '${hit.name}' is invalid because it is declared as ${hit.asType}. This is a VBE compile error: ${hit.vbeError}.`,
						hit.span,
					);
				}
			}, activity);
		}
}

function scalarMemberAccesses(
	source: string,
	span: Span,
	env: ReadonlyMap<string, string>,
): Array<{ name: string; asType: string; span: Span; vbeError: string }> {
	const toks = statementTokens(source, span);
	const out: Array<{ name: string; asType: string; span: Span; vbeError: string }> = [];
	for (let i = 0; i < toks.length - 1; i++) {
		if (toks[i + 1].rawText !== '.') {
			continue;
		}
		const name = tokenName(toks[i]);
		if (!name) {
			continue;
		}
		const asType = env.get(name.toLowerCase());
		const normalized = normalizeType(asType);
		if (!asType || !normalized || !isKnownScalarType(normalized)) {
			continue;
		}
		const memberName = toks[i + 2] ? tokenName(toks[i + 2]) : undefined;
		out.push({
			name,
			asType,
			vbeError: memberName ? 'Invalid qualifier' : 'Syntax error',
			span: { start: span.start + toks[i].start, end: span.start + toks[i + 1].end },
		});
	}
	return out;
}

function setAssignmentTarget(
	source: string,
	span: Span,
): { name: string; span: Span; valueTokens: VbaToken[] } | undefined {
	const toks = statementTokens(source, span);
	const i = firstExecutableTokenIndex(toks);
	if (tokenText(toks[i]) !== 'set') {
		return undefined;
	}
	const nameTok = toks[i + 1];
	if (!nameTok || nameTok.kind !== 'identifier') {
		return undefined;
	}
	const equals = toks[i + 2];
	if (!equals || equals.kind !== 'operator' || equals.rawText !== '=') {
		return undefined;
	}
	return {
		name: nameTok.rawText,
		span: { start: span.start + nameTok.start, end: span.start + nameTok.end },
		valueTokens: toks.slice(i + 3),
	};
}

function buildModuleTypeSignatures(
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, CallableTypeSignature> {
	const out = new Map<string, CallableTypeSignature>();
	for (const symbol of symbols.root.children ?? []) {
		if (isProcedureKind(symbol.kind) || symbol.kind === 'declare') {
			out.set(symbol.name.toLowerCase(), callableTypeSignatureFromSymbol(symbol));
		}
	}
	return out;
}

function sameModuleCallableSignatures(
	symbols: ReturnType<typeof buildModuleSymbols>,
): Map<string, CallableTypeSignature[]> {
	const out = new Map<string, CallableTypeSignature[]>();
	for (const symbol of symbols.root.children ?? []) {
		if (!isBareCallableKind(symbol.kind)) {
			continue;
		}
		const signature = callableTypeSignatureFromSymbol(symbol);
		const key = signature.name.toLowerCase();
		const arr = out.get(key);
		if (arr) {
			arr.push(signature);
		} else {
			out.set(key, [signature]);
		}
	}
	return out;
}

function callableTypeSignatureFromSymbol(symbol: VbaSymbol): CallableTypeSignature {
	return {
		name: symbol.name,
		params: procedureParamsFromSymbol(symbol, { includePassing: true }).map((p) => ({
			name: stripHeaderBrackets(p.name),
			type: p.type,
			optional: p.optional,
			paramArray: p.paramArray,
			byRef: isByRefProcedureParam(p),
		})),
		returnType: symbol.asType,
	};
}

function callableTypeSignaturesFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
): Map<string, CallableTypeSignature> {
	const out = buildModuleTypeSignatures(symbols);
	for (const [lower, sig] of uniqueProjectTypeSignatures(projectProcedures)) {
		if (!out.has(lower)) {
			out.set(lower, sig);
		}
	}
	return out;
}

function uniqueProjectTypeSignatures(
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
): Map<string, CallableTypeSignature> {
	const out = new Map<string, CallableTypeSignature>();
	if (!projectProcedures) {
		return out;
	}
	for (const [lower, candidates] of projectProcedures) {
		if (candidates.length !== 1) {
			continue;
		}
		const candidate = candidates[0];
		out.set(lower, {
			name: candidate.name,
			params: candidate.params.map((p) => ({
				name: p.name,
				type: p.type,
				optional: p.optional,
				paramArray: p.paramArray,
				byRef: isByRefProcedureParam(p),
			})),
			returnType: candidate.returnType,
		});
	}
	return out;
}

function isByRefProcedureParam(param: { byRef?: boolean; byVal?: boolean; paramArray?: boolean }): boolean {
	if (param.paramArray) {
		return false;
	}
	return param.byRef === true || param.byVal !== true;
}

function typeEnvironmentFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): Map<string, string> {
	const out = new Map<string, string>();
	for (const sym of symbols.root.children ?? []) {
		if (sym.asType && !isProcedureKind(sym.kind)) {
			out.set(sym.name.toLowerCase(), sym.asType);
		}
	}
	const procSym = (symbols.root.children ?? []).find(
		(s) => isProcedureKind(s.kind) && s.fullSpan.start === proc.span.start,
	);
	const returnType = returnAssignmentTypeFor(proc);
	if (returnType) {
		out.set(proc.name.toLowerCase(), returnType);
	}
	for (const child of procSym?.children ?? []) {
		if (child.asType) {
			out.set(child.name.toLowerCase(), child.asType);
		}
	}
	return out;
}

interface SourceNameScope {
	/**
	 * Non-callable names visible at the current expression/call site. These block
	 * bare callable resolution before same-module, project, or runtime signatures.
	 */
	callableShadows: ReadonlySet<string>;
	/**
	 * Any source-backed identifier visible in the current procedure. These block
	 * runtime fallback once source/project callable signatures have not resolved.
	 */
	runtimeShadows: ReadonlySet<string>;
}

function sourceNameScopeFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): SourceNameScope {
	const callableShadows = new Set(moduleNonCallableSymbols(symbols).keys());
	const runtimeShadows = moduleLevelIdentifierNames(symbols);
	const procSym = (symbols.root.children ?? []).find(
		(s) => isProcedureKind(s.kind) && s.fullSpan.start === proc.span.start,
	);
	for (const child of procSym?.children ?? []) {
		const lower = child.name.toLowerCase();
		runtimeShadows.add(lower);
		if (isNonCallableSymbol(child)) {
			callableShadows.add(lower);
		}
	}
	return { callableShadows, runtimeShadows };
}

function bareCallableSourceShadowed(
	name: string,
	sourceNames: SourceNameScope | undefined,
): boolean {
	return sourceNames?.callableShadows.has(name.toLowerCase()) === true;
}

function runtimeCallableSourceShadowed(
	name: string,
	sourceNames: SourceNameScope | undefined,
): boolean {
	return sourceNames?.runtimeShadows.has(name.toLowerCase()) === true;
}

function returnAssignmentTypeFor(proc: ProcedureNode): string | undefined {
	return (proc.procKind === 'Function' || proc.procKind === 'PropertyGet')
		? proc.returnType
		: undefined;
}

function expressionCalls(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallArguments[] {
	const toks = statementTokens(source, span);
	const out: CallArguments[] = [];
	for (let i = 0; i < toks.length - 1; i++) {
		const name = tokenName(toks[i]);
		if (!name || toks[i + 1].rawText !== '(') {
			continue;
		}
		const qualifier =
			i >= 2 && toks[i - 1].rawText === '.'
				? tokenName(toks[i - 2])
				: undefined;
		const lookupKey = qualifier ? qualifiedProcedureKey(qualifier, name) : undefined;
		if (qualifier && !moduleSignatures.has(lookupKey!)) {
			continue; // host/member calls need receiver binding before checking
		}
		if (!qualifier && i > 0 && toks[i - 1].rawText === '.') {
			continue;
		}
		if (
			lookupKey
				? !moduleSignatures.has(lookupKey)
				: !callableSignatureFor(name, moduleSignatures, sourceNames)
		) {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const inner = toks.slice(i + 2, close);
		const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
		out.push({
			name,
			qualifier,
			lookupKey,
			nameSpan: { start: span.start + toks[i].start, end: span.start + toks[i].end },
			slots: split.slots,
			slotSpans: split.spans,
			sliceStart: span.start,
		});
	}
	return out;
}

interface BoundMemberCall {
	call: CallArguments;
	signature: CallableTypeSignature;
}

function memberExpressionCalls(
	source: string,
	span: Span,
	memberCtx: MemberCompletionContext,
): BoundMemberCall[] {
	const toks = statementTokens(source, span);
	const standaloneEmptyCall = standaloneEmptyParenthesizedCallStatement(source, span);
	const out: BoundMemberCall[] = [];
	for (let i = 1; i < toks.length - 1; i++) {
		const name = tokenName(toks[i]);
		if (!name || toks[i - 1]?.rawText !== '.' || toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0) {
			continue;
		}
		const member = resolveExactMemberCompletion(
			source,
			name,
			span.start + toks[i].end,
			memberCtx,
		);
		if (!member?.signature) {
			continue;
		}
		const inner = toks.slice(i + 2, close);
		const callSpan = {
			start: span.start + toks[i].start,
			end: span.start + toks[close].end,
		};
		if (
			standaloneEmptyCall?.isMember &&
			standaloneEmptyCall.span.start === callSpan.start &&
			standaloneEmptyCall.span.end === callSpan.end
		) {
			continue;
		}
		const signature = parseRuntimeDisplaySignature(member.name, member.signature);
		if (isPropertyResultIndexing(member, signature, inner)) {
			continue;
		}
		const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
		out.push({
			signature,
			call: {
				name: member.name,
				nameSpan: { start: callSpan.start, end: span.start + toks[i].end },
				slots: split.slots,
				slotSpans: split.spans,
				sliceStart: span.start,
			},
		});
	}
	return out;
}

function memberStatementCalls(
	source: string,
	span: Span,
	memberCtx: MemberCompletionContext,
): BoundMemberCall[] {
	const toks = statementTokensAfterLeadingLabel(source, span);
	if (toks.length === 0 || topLevelOperatorIndex(toks, '=') >= 0) {
		return [];
	}
	const explicitCall = tokenText(toks[0]) === 'call';
	const chainStart = explicitCall ? 1 : 0;
	if (!tokenName(toks[chainStart]) && toks[chainStart]?.rawText !== '.') {
		return [];
	}
	const out: BoundMemberCall[] = [];
	const firstMemberIndex = toks[chainStart]?.rawText === '.' ? chainStart + 1 : chainStart + 2;
	for (let i = firstMemberIndex; i < toks.length; i++) {
		const name = tokenName(toks[i]);
		if (!name || toks[i - 1]?.rawText !== '.') {
			continue;
		}
		if (!isMemberStatementChainThrough(toks, chainStart, i)) {
			continue;
		}
		const next = toks[i + 1];
		if (next?.rawText === '(') {
			continue; // parenthesized member calls are handled by memberExpressionCalls
		}
		if (explicitCall && next) {
			continue; // Call p.Save arg is a call-requires-parens syntax error
		}
		if (next) {
			const gap = source.slice(span.start + toks[i].end, span.start + next.start);
			if (!/\s/.test(gap) || !isMemberParenlessArgumentStart(next)) {
				continue;
			}
		}
		const member = resolveExactMemberCompletion(
			source,
			name,
			span.start + toks[i].end,
			memberCtx,
		);
		if (!member?.signature) {
			continue;
		}
		const argToks = toks.slice(i + 1);
		const split = argToks.length === 0 ? emptyArgSplit() : splitArgSlots(argToks, span.start);
		out.push({
			signature: parseRuntimeDisplaySignature(member.name, member.signature),
			call: {
				name: member.name,
				nameSpan: { start: span.start + toks[i].start, end: span.start + toks[i].end },
				explicitCall,
				slots: split.slots,
				slotSpans: split.spans,
				sliceStart: span.start,
			},
		});
		break;
	}
	return out;
}

function isPropertyResultIndexing(
	member: MemberCompletion,
	signature: CallableTypeSignature,
	inner: readonly VbaToken[],
): boolean {
	return member.kind === 'property' &&
		signature.params.length === 0 &&
		inner.length > 0;
}

function isMemberStatementChainThrough(
	toks: readonly VbaToken[],
	startIdx: number,
	memberIdx: number,
): boolean {
	if (toks[startIdx]?.rawText === '.') {
		if (!tokenName(toks[startIdx + 1])) {
			return false;
		}
		if (startIdx + 1 === memberIdx) {
			return true;
		}
		return isMemberStatementChainThrough(toks, startIdx + 1, memberIdx);
	}
	if (!tokenName(toks[startIdx])) {
		return false;
	}
	let i = startIdx + 1;
	while (i < toks.length) {
		const raw = toks[i]?.rawText;
		if (raw === '(') {
			const close = matchParenFrom(toks, i);
			if (close < 0 || close >= memberIdx) {
				return false;
			}
			i = close + 1;
			continue;
		}
		if (raw !== '.') {
			return false;
		}
		const nameIdx = i + 1;
		if (!tokenName(toks[nameIdx])) {
			return false;
		}
		if (nameIdx === memberIdx) {
			return true;
		}
		i = nameIdx + 1;
	}
	return false;
}

function isMemberParenlessArgumentStart(tok: VbaToken): boolean {
	if (
		tok.kind === 'identifier' ||
		tok.kind === 'keyword' ||
		tok.kind === 'bracketedIdentifier' ||
		tok.kind === 'stringLiteral' ||
		tok.kind === 'dateLiteral' ||
		tok.kind === 'integerLiteral' ||
		tok.kind === 'floatLiteral'
	) {
		return true;
	}
	return tok.rawText === ',' || tok.rawText === '+' || tok.rawText === '-';
}

function validateArgumentTypes(
	call: CallArguments,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
	source: string,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const sig = callableSignatureForCall(call, moduleSignatures, sourceNames);
	if (!sig || sig.params.length === 0) {
		return;
	}
	validateArgumentTypesForSignature(sig, call, env, moduleSignatures, sourceNames, source, memberCtx, push);
}

function validateArgumentTypesForSignature(
	sig: CallableTypeSignature,
	call: CallArguments,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames: SourceNameScope | undefined,
	source: string,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	if (sig.params.length === 0) {
		return;
	}
	const paramsByName = new Map(
		sig.params.map((p) => [stripHeaderBrackets(p.name).toLowerCase(), p]),
	);
	let positionalIndex = 0;
	for (let i = 0; i < call.slots.length; i++) {
		const named = namedArgumentSlot(call.slots[i]);
		let param: CallableParamType | undefined;
		let valueSlot = call.slots[i];
		if (named) {
			param = paramsByName.get(named.name.toLowerCase());
			valueSlot = named.value;
		} else {
			param = sig.params[Math.min(positionalIndex, sig.params.length - 1)];
			if (!param || (positionalIndex >= sig.params.length && !param.paramArray)) {
				continue;
			}
			positionalIndex++;
		}
		if (!param) {
			continue;
		}
		const expected = param.type;
		if (!expected) {
			continue;
		}
		const byRefMismatch = byRefVariableTypeMismatch(param, valueSlot, call.sliceStart, env);
		if (byRefMismatch) {
			push(
				'byRefArgumentTypeMismatch',
				`ByRef argument '${param.name}' of '${sig.name}' expects ${expected}, but '${byRefMismatch.name}' is declared as ${byRefMismatch.actual}. This is a VBE compile error: ByRef argument type mismatch.`,
				byRefMismatch.span,
			);
			continue;
		}
		const stringArithmetic = nonnumericStringArithmeticOperand(
			expected,
			valueSlot,
			call.sliceStart,
		);
		if (stringArithmetic) {
			push(
				'stringArithmeticCoercion',
				`Argument '${param.name}' of '${sig.name}' expects ${expected}, but this numeric expression contains ${stringArithmetic.label}. This will raise Run-time error '13': Type mismatch.`,
				stringArithmetic.span,
			);
			continue;
		}
		const actual = inferArgumentType(
			valueSlot,
			call.sliceStart,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
		);
		if (!actual) {
			continue;
		}
		const reason = incompatibilityReason(expected, actual);
		if (!reason) {
			continue;
		}
		const rule =
			normalizeType(expected) === 'object'
				? 'argumentObjectTypeMismatch'
				: 'argumentTypeMismatch';
		push(
			rule,
			`Argument '${param.name}' of '${sig.name}' expects ${expected}, but got ${actual.label}. ${reason}`,
			actual.span,
		);
	}
}

function callableSignatureForCall(
	call: CallArguments,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallableTypeSignature | undefined {
	if (call.lookupKey) {
		return moduleSignatures.get(call.lookupKey);
	}
	return callableSignatureFor(call.name, moduleSignatures, sourceNames);
}

function byRefVariableTypeMismatch(
	param: CallableParamType,
	slot: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
): { name: string; actual: string; span: Span } | undefined {
	if (!param.byRef || !param.type) {
		return undefined;
	}
	const expected = normalizeType(param.type);
	if (!expected || !isKnownScalarType(expected)) {
		return undefined;
	}
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	if (toks.length !== 1) {
		return undefined;
	}
	const name = tokenName(toks[0]);
	if (!name) {
		return undefined;
	}
	const actualRaw = env.get(name.toLowerCase());
	const actual = normalizeType(actualRaw);
	if (!actual || !isKnownScalarType(actual) || actual === expected) {
		return undefined;
	}
	return {
		name,
		actual: actualRaw ?? name,
		span: { start: sliceStart + toks[0].start, end: sliceStart + toks[0].end },
	};
}

function namedArgumentSlot(slot: VbaToken[]): { name: string; value: VbaToken[] } | undefined {
	if (!isNamedSlot(slot)) {
		return undefined;
	}
	return {
		name: stripHeaderBrackets(slot[0].rawText),
		value: slot.slice(2),
	};
}

function callableSignatureFor(
	name: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallableTypeSignature | undefined {
	if (bareCallableSourceShadowed(name, sourceNames)) {
		return undefined;
	}
	const user = moduleSignatures.get(name.toLowerCase());
	if (user) {
		return user;
	}
	if (runtimeCallableSourceShadowed(name, sourceNames)) {
		return undefined;
	}
	const runtime = resolveRuntimeFunction(name);
	if (!runtime) {
		return undefined;
	}
	return runtimeTypeSignature(runtime);
}

function runtimeTypeSignature(runtime: VbaRuntimeFunction): CallableTypeSignature {
	if (runtime.params) {
		return {
			name: runtime.name,
			params: runtime.params.map((p) => ({
				name: p.name,
				type: p.type,
				optional: p.optional ?? false,
				paramArray: p.paramArray ?? false,
			})),
			returnType: runtime.returns,
		};
	}
	return parseRuntimeDisplaySignature(runtime.name, runtime.signature, runtime.returns);
}

function runtimeAritySignature(runtime: VbaRuntimeFunction): CallableTypeSignature | undefined {
	if (runtime.params || runtimeSignatureParameterText(runtime.signature) !== undefined) {
		return runtimeTypeSignature(runtime);
	}
	return undefined;
}

function parseRuntimeDisplaySignature(
	name: string,
	signature: string,
	returnType?: string,
): CallableTypeSignature {
	const inner = runtimeSignatureParameterText(signature);
	if (inner === undefined) {
		return { name, params: [], returnType };
	}
	const params = splitSignatureTopLevel(inner)
		.map(parseRuntimeParamType)
		.filter((p): p is CallableParamType => p !== undefined);
	return { name, params, returnType };
}

function runtimeSignatureParameterText(signature: string): string | undefined {
	const open = signature.indexOf('(');
	const close = signature.lastIndexOf(')');
	if (open < 0 || close < open) {
		return undefined;
	}
	return signature.slice(open + 1, close);
}

function parseRuntimeParamType(raw: string): CallableParamType | undefined {
	let text = raw.trim();
	if (!text) {
		return undefined;
	}
	const optional = text.startsWith('[') && text.endsWith(']');
	text = text.replace(/^\[/, '').replace(/\]$/, '').trim();
	const paramArray = /^ParamArray\b/i.test(text);
	text = text.replace(/^ParamArray\b\s*/i, '');
	text = text.replace(/^(?:ByVal|ByRef)\b\s*/i, '');
	text = text.replace(/\s*=\s*.*$/, '').trim();
	const as = /\bAs\s+([A-Za-z_][A-Za-z0-9_]*(?:\(\))?)/i.exec(text);
	const first = /[A-Za-z_][A-Za-z0-9_]*/.exec(text)?.[0];
	if (!first) {
		return undefined;
	}
	return {
		name: first,
		type: as?.[1],
		optional,
		paramArray,
	};
}

function splitSignatureTopLevel(text: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let start = 0;
	for (let i = 0; i < text.length; i++) {
		const c = text[i];
		if (c === '(' || c === '[') {
			depth++;
		} else if (c === ')' || c === ']') {
			depth--;
		} else if (c === ',' && depth === 0) {
			out.push(text.slice(start, i));
			start = i + 1;
		}
	}
	out.push(text.slice(start));
	return out;
}

function inferArgumentType(
	slot: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
): InferredArgumentType | undefined {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return inferExpressionType(toks, sliceStart, env, moduleSignatures, sourceNames, source, memberCtx);
}

function inferExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
): InferredArgumentType | undefined {
	const first = toks[0];
	if (!first) {
		return undefined;
	}
	const unwrapped = unwrapOuterParens(toks);
	if (unwrapped !== toks) {
		return inferExpressionType(unwrapped, sliceStart, env, moduleSignatures, sourceNames, source, memberCtx);
	}
	const signedNumericLiteral = inferSignedNumericLiteral(toks, sliceStart);
	if (signedNumericLiteral) {
		return signedNumericLiteral;
	}
	const concatenation = inferStringConcatenationExpressionType(
		toks,
		sliceStart,
		env,
		moduleSignatures,
		sourceNames,
		source,
		memberCtx,
	);
	if (concatenation) {
		return concatenation;
	}
	const arithmetic = inferArithmeticExpressionType(
		toks,
		sliceStart,
		env,
		moduleSignatures,
		sourceNames,
		source,
		memberCtx,
	);
	if (arithmetic) {
		return arithmetic;
	}
	return inferAtomicExpressionType(toks, sliceStart, env, moduleSignatures, sourceNames, source, memberCtx);
}

function inferSignedNumericLiteral(
	toks: VbaToken[],
	sliceStart: number,
): InferredArgumentType | undefined {
	if (toks.length !== 2 || toks[0].kind !== 'operator') {
		return undefined;
	}
	const sign = toks[0].rawText;
	if (sign !== '+' && sign !== '-') {
		return undefined;
	}
	const literal = toks[1];
	if (literal.kind !== 'integerLiteral') {
		return undefined;
	}
	const value = parseDecimalIntegerLiteral(literal.rawText);
	if (value === undefined) {
		return undefined;
	}
	const signed = sign === '-' ? -value : value;
	const text = `${sign}${literal.rawText}`;
	return {
		type: 'Double',
		label: `numeric literal ${text}`,
		span: { start: sliceStart + toks[0].start, end: sliceStart + literal.end },
		numericValue: signed,
		numericText: text,
	};
}

function inferAtomicExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
): InferredArgumentType | undefined {
	const first = toks[0];
	if (!first) {
		return undefined;
	}
	const span = { start: sliceStart + first.start, end: sliceStart + first.end };
	if (toks.length === 1) {
		switch (first.kind) {
			case 'stringLiteral': {
				const value = stringLiteralValue(first.rawText);
				return { type: 'String', label: `String literal ${first.rawText}`, span, stringValue: value };
			}
			case 'integerLiteral':
			case 'floatLiteral': {
				const numericValue =
					first.kind === 'integerLiteral'
						? parseDecimalIntegerLiteral(first.rawText)
						: undefined;
				return {
					type: 'Double',
					label: `numeric literal ${first.rawText}`,
					span,
					numericValue,
					numericText: first.rawText,
				};
			}
			case 'dateLiteral':
				return { type: 'Date', label: 'Date literal', span };
			case 'keyword': {
				const word = first.rawText.toLowerCase();
				if (word === 'true' || word === 'false') {
					return { type: 'Boolean', label: 'Boolean literal', span };
				}
				if (word === 'nothing') {
					return { type: 'Nothing', label: 'Nothing', span };
				}
				break;
			}
			default:
				break;
		}
	}
	const name = tokenName(first);
	if (name && toks.length === 1) {
		const type = env.get(name.toLowerCase());
		if (type) {
			return { type, label: `${name} As ${type}`, span };
		}
		const sig = parameterlessValueSignature(name, moduleSignatures, sourceNames);
		if (sig?.returnType) {
			return { type: sig.returnType, label: `${name} As ${sig.returnType}`, span };
		}
		return undefined;
	}
	if (tokenText(first) === 'new' && toks.length === 2) {
		const typeName = tokenName(toks[1]);
		if (typeName) {
			return {
				type: typeName,
				label: `New ${typeName}`,
				span: { start: sliceStart + toks[1].start, end: sliceStart + toks[1].end },
			};
		}
	}
	if (name && toks[1]?.rawText === '(') {
		const sig = callableSignatureFor(name, moduleSignatures, sourceNames);
		if (sig?.returnType && matchParenFrom(toks, 1) === toks.length - 1) {
			return { type: sig.returnType, label: `${name}(...) As ${sig.returnType}`, span };
		}
	}
	if (name && toks[1]?.rawText === '.') {
		const member = tokenName(toks[2]);
		if (member && toks.length === 3) {
			const lookupKey = qualifiedProcedureKey(name, member);
			const sig = parameterlessValueSignature(lookupKey, moduleSignatures);
			if (sig?.returnType) {
				return {
					type: sig.returnType,
					label: `${name}.${member} As ${sig.returnType}`,
					span: { start: sliceStart + toks[2].start, end: sliceStart + toks[2].end },
				};
			}
		}
		if (member && toks[3]?.rawText === '(' && matchParenFrom(toks, 3) === toks.length - 1) {
			const lookupKey = qualifiedProcedureKey(name, member);
			const sig = moduleSignatures.get(lookupKey);
			if (sig?.returnType) {
				return {
					type: sig.returnType,
					label: `${name}.${member}(...) As ${sig.returnType}`,
					span: { start: sliceStart + toks[2].start, end: sliceStart + toks[2].end },
				};
			}
		}
	}
	const memberType = source && memberCtx
		? inferMemberExpressionType(source, toks, sliceStart, memberCtx)
		: undefined;
	if (memberType) {
		return memberType;
	}
	return undefined;
}

function inferMemberExpressionType(
	source: string,
	toks: VbaToken[],
	sliceStart: number,
	memberCtx: MemberCompletionContext,
): InferredArgumentType | undefined {
	if (hasTopLevelOperator(toks)) {
		return undefined;
	}
	const resolved = finalMemberTokenInExpression(toks);
	if (!resolved) {
		return undefined;
	}
	const member = resolveExactMemberCompletion(
		source,
		resolved.name,
		sliceStart + resolved.token.end,
		memberCtx,
	);
	if (!member?.returns) {
		return undefined;
	}
	if (!resolved.called && member.kind === 'method' && !memberAcceptsZeroArguments(member)) {
		return undefined;
	}
	const labelStart = toks[0]?.start ?? resolved.token.start;
	const labelEnd = resolved.called ? toks[toks.length - 1].end : resolved.token.end;
	const labelText = source.slice(sliceStart + labelStart, sliceStart + labelEnd).trim();
	return {
		type: member.returns,
		label: `${labelText} As ${member.returns}`,
		span: { start: sliceStart + resolved.token.start, end: sliceStart + resolved.token.end },
	};
}

function finalMemberTokenInExpression(
	toks: readonly VbaToken[],
): { name: string; token: VbaToken; called: boolean } | undefined {
	const last = toks[toks.length - 1];
	if (!last) {
		return undefined;
	}
	if (tokenName(last) && toks[toks.length - 2]?.rawText === '.') {
		return { name: tokenName(last)!, token: last, called: false };
	}
	if (last.rawText !== ')') {
		return undefined;
	}
	const open = matchingOpenParenIndex(toks, toks.length - 1);
	if (open < 2) {
		return undefined;
	}
	const member = toks[open - 1];
	if (!tokenName(member) || toks[open - 2]?.rawText !== '.') {
		return undefined;
	}
	return { name: tokenName(member)!, token: member, called: true };
}

function matchingOpenParenIndex(toks: readonly VbaToken[], close: number): number {
	let depth = 0;
	for (let i = close; i >= 0; i--) {
		const raw = toks[i].rawText;
		if (raw === ')') {
			depth++;
		} else if (raw === '(') {
			depth--;
			if (depth === 0) {
				return i;
			}
		}
	}
	return -1;
}

function hasTopLevelOperator(toks: readonly VbaToken[]): boolean {
	let depth = 0;
	for (const tok of toks) {
		const raw = tok.rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (depth === 0 && tok.kind === 'operator') {
			return true;
		}
	}
	return false;
}

function memberAcceptsZeroArguments(member: MemberCompletion): boolean {
	if (!member.signature) {
		return false;
	}
	return callableAcceptsZeroArguments(parseRuntimeDisplaySignature(member.name, member.signature));
}

function parameterlessValueSignature(
	name: string,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
): CallableTypeSignature | undefined {
	const sig = callableSignatureFor(name, moduleSignatures, sourceNames);
	return sig?.returnType && callableAcceptsZeroArguments(sig) ? sig : undefined;
}

function unwrapOuterParens(toks: VbaToken[]): VbaToken[] {
	if (toks.length < 2 || toks[0].rawText !== '(') {
		return toks;
	}
	const close = matchParenFrom(toks, 0);
	return close === toks.length - 1 ? toks.slice(1, -1) : toks;
}

function inferArithmeticExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
): InferredArgumentType | undefined {
	const parts = splitTopLevelArithmeticOperands(toks);
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const inferred = inferExpressionType(
			part,
			sliceStart,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
		);
		const normalized = normalizeType(inferred?.type);
		if (!normalized || !isNumericType(normalized)) {
			return undefined;
		}
	}
	return {
		type: 'Double',
		label: 'numeric expression',
		span: spanForTokens(toks, sliceStart),
	};
}

function nonnumericStringArithmeticOperand(
	expectedRaw: string,
	slot: VbaToken[],
	sliceStart: number,
): InferredArgumentType | undefined {
	const expected = normalizeType(expectedRaw);
	if (!expected || !isNumericType(expected)) {
		return undefined;
	}
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return findNonnumericStringInArithmeticExpression(toks, sliceStart);
}

function findNonnumericStringInArithmeticExpression(
	toks: VbaToken[],
	sliceStart: number,
): InferredArgumentType | undefined {
	const unwrapped = unwrapOuterParens(toks);
	if (unwrapped !== toks) {
		return findNonnumericStringInArithmeticExpression(unwrapped, sliceStart);
	}
	const parts = splitTopLevelArithmeticOperands(toks);
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const nested = findNonnumericStringInArithmeticExpression(part, sliceStart);
		if (nested) {
			return nested;
		}
		const operand = unwrapOuterParens(part);
		if (operand.length === 1 && operand[0].kind === 'stringLiteral') {
			const value = stringLiteralValue(operand[0].rawText);
			if (isProvablyNonNumericString(value)) {
				return {
					type: 'String',
					label: `nonnumeric string literal ${operand[0].rawText}`,
					span: { start: sliceStart + operand[0].start, end: sliceStart + operand[0].end },
					stringValue: value,
				};
			}
		}
	}
	return undefined;
}

function inferStringConcatenationExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	sourceNames?: SourceNameScope,
	source?: string,
	memberCtx?: MemberCompletionContext,
): InferredArgumentType | undefined {
	const parts = splitTopLevelOperands(toks, '&');
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const inferred = inferExpressionType(
			part,
			sliceStart,
			env,
			moduleSignatures,
			sourceNames,
			source,
			memberCtx,
		);
		const normalized = normalizeType(inferred?.type);
		if (!normalized || !isStringConcatenationOperandType(normalized)) {
			return undefined;
		}
	}
	return {
		type: 'String',
		label: 'string concatenation expression',
		span: spanForTokens(toks, sliceStart),
	};
}

function splitTopLevelArithmeticOperands(toks: VbaToken[]): VbaToken[][] {
	const parts = splitTopLevelOperands(toks, '+', '-', '*', '/', '\\', '^');
	if (parts.length < 2) {
		return [];
	}
	return parts;
}

function splitTopLevelOperands(toks: VbaToken[], ...operators: string[]): VbaToken[][] {
	const allowed = new Set(operators);
	const parts: VbaToken[][] = [];
	let start = 0;
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
			continue;
		}
		if (raw === ')' || raw === ']') {
			depth--;
			continue;
		}
		if (depth !== 0) {
			continue;
		}
		if (toks[i].kind !== 'operator' || !allowed.has(toks[i].rawText)) {
			if (toks[i].kind === 'operator') {
				return [];
			}
			continue;
		}
		if (i === start || i === toks.length - 1) {
			return [];
		}
		parts.push(toks.slice(start, i));
		start = i + 1;
	}
	if (parts.length === 0) {
		return [];
	}
	parts.push(toks.slice(start));
	return parts;
}

function isStringConcatenationOperandType(type: string): boolean {
	return (
		type === 'string' ||
		type === 'boolean' ||
		type === 'date' ||
		isNumericType(type)
	);
}

function spanForTokens(toks: VbaToken[], sliceStart: number): Span {
	const first = toks[0];
	const last = toks[toks.length - 1];
	return { start: sliceStart + first.start, end: sliceStart + last.end };
}

function incompatibilityReason(
	expectedRaw: string,
	actual: InferredArgumentType,
): string | undefined {
	const expected = normalizeType(expectedRaw);
	const actualType = normalizeType(actual.type);
	if (!expected || !actualType || expected === 'variant' || actualType === 'variant') {
		return undefined;
	}
	if (expected === 'object') {
		return actualType === 'nothing' || actualType === 'object' || !isKnownScalarType(actualType)
			? undefined
			: 'An object parameter requires an object value.';
	}
	if (isNumericType(expected)) {
		const overflow = numericLiteralOverflowReason(expected, actual);
		if (overflow) {
			return overflow;
		}
		if (isNumericType(actualType) || actualType === 'boolean') {
			return undefined;
		}
		if (actualType === 'string') {
			return actual.stringValue !== undefined && isProvablyNonNumericString(actual.stringValue)
				? "This string literal cannot be converted to a numeric value. This will raise Run-time error '13': Type mismatch."
				: undefined;
		}
		return undefined;
	}
	if (expected === 'boolean') {
		if (actualType === 'boolean' || isNumericType(actualType)) {
			return undefined;
		}
		if (actualType === 'string') {
			return actual.stringValue !== undefined && isBooleanString(actual.stringValue)
				? undefined
				: "This string literal cannot be converted to Boolean. This will raise Run-time error '13': Type mismatch.";
		}
		return undefined;
	}
	if (expected === 'string') {
		return undefined; // VBA can stringify scalar values; do not warn.
	}
	return undefined;
}

function numericLiteralOverflowReason(
	expected: string,
	actual: InferredArgumentType,
): string | undefined {
	if (actual.numericValue === undefined) {
		return undefined;
	}
	const bounds = numericLiteralBounds(expected);
	if (!bounds) {
		return undefined;
	}
	if (actual.numericValue >= bounds.min && actual.numericValue <= bounds.max) {
		return undefined;
	}
	const literal = actual.numericText ?? String(actual.numericValue);
	return `The numeric literal ${literal} is outside the ${bounds.label} range ${bounds.min} to ${bounds.max}. This will raise Run-time error '6': Overflow.`;
}

function numericLiteralBounds(
	expected: string,
): { min: number; max: number; label: string } | undefined {
	switch (expected) {
		case 'byte':
			return { min: 0, max: 255, label: 'Byte' };
		case 'integer':
			return { min: -32768, max: 32767, label: 'Integer' };
		default:
			return undefined;
	}
}

function normalizeType(type: string | undefined): string | undefined {
	if (!type) {
		return undefined;
	}
	return type
		.replace(/\(\)$/, '')
		.replace(/^vb/i, '')
		.trim()
		.toLowerCase();
}

function isNumericType(type: string): boolean {
	return new Set([
		'byte',
		'integer',
		'long',
		'longlong',
		'longptr',
		'single',
		'double',
		'currency',
		'decimal',
	]).has(type);
}

function isKnownScalarType(type: string): boolean {
	return type === 'string' || type === 'boolean' || type === 'date' || isNumericType(type);
}

function isKnownObjectAssignmentType(
	type: string | undefined,
	memberCtx: MemberCompletionContext,
): boolean {
	return resolveKnownObjectAssignmentType(type, memberCtx) !== undefined;
}

type KnownObjectAssignmentType =
	| { kind: 'generic'; display: string; key: 'object' }
	| { kind: 'host'; display: string; key: string }
	| { kind: 'project'; display: string; key: string; implements: readonly string[] };

function resolveKnownObjectAssignmentType(
	type: string | undefined,
	memberCtx: MemberCompletionContext,
): KnownObjectAssignmentType | undefined {
	if (!type) {
		return undefined;
	}
	const normalized = normalizeType(type);
	if (!normalized || normalized === 'variant') {
		return undefined;
	}
	if (normalized === 'object') {
		return { kind: 'generic', display: type, key: 'object' };
	}
	if (isKnownScalarType(normalized)) {
		return undefined;
	}
	const host = resolveHostAlias(type, memberCtx.model);
	if (host) {
		return { kind: 'host', display: type, key: host.toLowerCase() };
	}
	const simple = simpleTypeNameForAssignment(type);
	if (!simple) {
		return undefined;
	}
	const lower = simple.toLowerCase();
	const matches = (memberCtx.projectClassMembers ?? []).filter(
		(projectType) =>
			projectType.kind !== 'userType' &&
			projectType.kind !== 'standardModule' &&
			projectType.name.toLowerCase() === lower,
	);
	if (matches.length !== 1) {
		return undefined;
	}
	return {
		kind: 'project',
		display: matches[0].name,
		key: lower,
		implements: matches[0].implements ?? [],
	};
}

function simpleTypeNameForAssignment(type: string): string | undefined {
	const trimmed = type.replace(/\(\)$/, '').trim();
	return /^[A-Za-z_][A-Za-z0-9_]*$/.test(trimmed) ? trimmed : undefined;
}

function objectAssignmentIncompatibilityReason(
	expectedRaw: string | undefined,
	actual: InferredArgumentType | undefined,
	memberCtx: MemberCompletionContext,
): string | undefined {
	const expected = resolveKnownObjectAssignmentType(expectedRaw, memberCtx);
	if (!expected || !actual) {
		return undefined;
	}
	const actualType = normalizeType(actual.type);
	if (!actualType || actualType === 'variant' || actualType === 'nothing') {
		return undefined;
	}
	if (isKnownScalarType(actualType)) {
		return 'An object assignment requires an object value.';
	}
	if (expected.kind === 'generic') {
		return undefined;
	}
	const actualObject = resolveKnownObjectAssignmentType(actual.type, memberCtx);
	if (!actualObject) {
		return undefined;
	}
	if (actualObject.kind === 'generic') {
		return undefined;
	}
	if (expected.key === actualObject.key) {
		return undefined;
	}
	if (actualObject.kind === 'project' && implementsObjectType(actualObject, expected)) {
		return undefined;
	}
	return `This object type is not compatible with ${expected.display}.`;
}

function implementsObjectType(
	actual: Extract<KnownObjectAssignmentType, { kind: 'project' }>,
	expected: KnownObjectAssignmentType,
): boolean {
	const expectedNames = new Set([expected.key]);
	const simple = simpleTypeNameForAssignment(expected.display);
	if (simple) {
		expectedNames.add(simple.toLowerCase());
	}
	const expectedLastSegment = expected.key.split('.').pop();
	if (expectedLastSegment) {
		expectedNames.add(expectedLastSegment);
	}
	return actual.implements.some((implemented) => {
		const lower = implemented.toLowerCase();
		return expectedNames.has(lower) || expectedNames.has(`excel.${lower}`);
	});
}

// One-way proof only: strings with digits are left unknown until VBA conversion
// semantics are modeled explicitly.
function isProvablyNonNumericString(value: string): boolean {
	const trimmed = value.trim();
	return trimmed.length > 0 && !/[0-9]/.test(trimmed);
}

function stringLiteralValue(raw: string): string {
	return raw
		.replace(/^"/, '')
		.replace(/"$/, '')
		.replace(/""/g, '"');
}

function isBooleanString(value: string): boolean {
	return /^(true|false|0|-?1)$/i.test(value.trim());
}

/**
 * Rule: a code module that contains real code but no `Option Explicit` lets
 * variables be used without declaration. Empty/attribute-only modules are
 * skipped to avoid noise on blank document modules.
 */
function checkOptionExplicit(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let hasExplicit = false;
	let hasCode = false;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Option' && /^explicit\b/i.test(member.optionText.trim())) {
			hasExplicit = true;
		}
		if (
			member.kind === 'Procedure' ||
			member.kind === 'VariableGroup' ||
			member.kind === 'Type' ||
			member.kind === 'Enum' ||
			member.kind === 'Declare'
		) {
			hasCode = true;
		}
	}
	if (hasExplicit || !hasCode) {
		return;
	}
	// Point at the first physical line so the squiggle sits at the module top.
	const nl = source.search(/\r|\n/);
	const end = nl === -1 ? source.length : nl;
	push(
		'optionExplicitMissing',
		'Option Explicit is not specified; variables can be used without being declared. Add "Option Explicit" to the top of the module.',
		{ start: 0, end },
	);
}

/**
 * Rule: with `Option Explicit`, a variable must be declared before it can be
 * assigned or read. The rule only runs once the caller has supplied the
 * project-visible identifier set, so cross-module globals and enum members do
 * not false-positive.
 */
function checkUndeclaredVariables(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	activity: ConditionalActivityTracker | undefined,
	knownIdentifiers: ReadonlySet<string> | undefined,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
	push: PushFn,
): void {
	if (!hasOptionExplicit(mod, activity) || !knownIdentifiers) {
		return;
	}

	const moduleNames = moduleLevelIdentifierNames(symbols);
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	const appType = resolveHostGlobal('Application');
	const appMembers = new Set(
		(appType ? getHostMembers(appType) : []).map((member) => member.name.toLowerCase()),
	);
	const isKnown = (name: string, locals: ReadonlySet<string>): boolean => {
		const lower = name.toLowerCase();
		return (
			lower === 'vba' ||
			locals.has(lower) ||
			moduleNames.has(lower) ||
			knownIdentifiers.has(lower) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveHostConstant(name) !== undefined ||
			resolveRuntimeConstant(name) !== undefined ||
			resolveRuntimeObject(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const procSym = (symbols.root.children ?? []).find(
			(sym) => isProcedureKind(sym.kind) && sym.fullSpan.start === member.span.start,
		);
		const locals = new Set<string>();
		for (const child of procSym?.children ?? []) {
			locals.add(child.name.toLowerCase());
		}
		forEachUndeclaredReferenceSpan(source, member.body, (span) => {
			const reported = new Set<string>();
			const report = (name: string, span: Span, mode: 'assigning to it' | 'using it'): void => {
				const key = `${span.start}:${span.end}`;
				if (reported.has(key) || isKnown(name, locals)) {
					return;
				}
				reported.add(key);
				push(
					'undeclaredVariable',
					`Variable not defined: '${name}'. Declare it before ${mode}, or remove Option Explicit.`,
					span,
				);
			};
			const scalarTarget = bareAssignmentTarget(source, span);
			const objectTarget = scalarTarget ? undefined : setAssignmentTarget(source, span);
			const target = scalarTarget ?? objectTarget;
			if (target) {
				report(target.name, target.span, 'assigning to it');
			}
			for (const ref of undeclaredReadReferences(
				source,
				span,
				(name) => isKnown(name, locals),
				moduleSignatures,
				projectMembers,
			)) {
				report(ref.name, ref.span, 'using it');
			}
		}, activity);
	}
}

function forEachUndeclaredReferenceSpan(
	source: string,
	body: BodyNode[],
	visit: (span: Span) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			visit(node.span);
		} else if ('body' in node && Array.isArray(node.body)) {
			visit(blockHeaderLineSpan(source, node.span));
			if (node.kind === 'DoBlock') {
				const footer = blockFooterLineSpan(source, node.span);
				if (footer.start > node.span.start) {
					visit(footer);
				}
			}
			forEachUndeclaredReferenceSpan(source, node.body, visit, activity);
		}
	}
}

function blockHeaderLineSpan(source: string, span: Span): Span {
	const nl = firstLineBreakAtOrAfter(source, span.start);
	if (nl < 0 || nl > span.end) {
		return span;
	}
	return { start: span.start, end: nl };
}

function blockFooterLineSpan(source: string, span: Span): Span {
	let start = span.end;
	while (start > span.start && source[start - 1] !== '\n' && source[start - 1] !== '\r') {
		start--;
	}
	return { start, end: span.end };
}

function firstLineBreakAtOrAfter(source: string, start: number): number {
	for (let i = start; i < source.length; i++) {
		const ch = source[i];
		if (ch === '\n' || ch === '\r') {
			return i;
		}
	}
	return -1;
}

function undeclaredReadReferences(
	source: string,
	span: Span,
	isKnown: (name: string) => boolean,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): Array<{ name: string; span: Span }> {
	const toks = statementTokens(source, span);
	const out: Array<{ name: string; span: Span }> = [];
	const skip = undeclaredReferenceSkipIndexes(
		source,
		span,
		toks,
		isKnown,
		moduleSignatures,
		projectMembers,
	);
	for (let i = 0; i < toks.length; i++) {
		if (skip.has(i) || !isPotentialVariableReferenceToken(toks[i])) {
			continue;
		}
		if (toks[i - 1]?.rawText === '.') {
			continue;
		}
		const name = tokenName(toks[i]);
		if (!name || isKnown(name)) {
			continue;
		}
		out.push({
			name,
			span: { start: span.start + toks[i].start, end: span.start + toks[i].end },
		});
	}
	return out;
}

function undeclaredReferenceSkipIndexes(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
	isKnown: (name: string) => boolean,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): Set<number> {
	const skip = new Set<number>();
	if (toks.length === 0) {
		return skip;
	}
	if (toks[1]?.rawText === ':' || isLineLabelOnlyStatement(source, span, toks)) {
		skip.add(0); // line label declaration
	}
	const firstExecutable = firstExecutableTokenIndex(toks);
	if (tokenText(toks[firstExecutable]) === 'implements') {
		for (let i = firstExecutable + 1; i < toks.length; i++) {
			skip.add(i);
		}
		return skip;
	}

	const call = callStatementTarget(source, span);
	if (call) {
		const callIdx = toks.findIndex((tok) => span.start + tok.start === call.span.start);
		if (callIdx >= 0) {
			skip.add(callIdx);
			if (!isKnown(call.name)) {
				// Unknown call targets may be external procedures or unresolved call
				// errors; do not also guess about their argument identifiers.
				for (let i = callIdx + 1; i < toks.length; i++) {
					skip.add(i);
				}
			}
		}
	}

	const assignment = simpleAssignmentLhsIdentifierIndex(toks);
	if (assignment >= 0) {
		skip.add(assignment);
	}

	for (let i = 0; i < toks.length; i++) {
		const word = tokenText(toks[i]);
		if (
			isQualifiedProjectCallableQualifier(toks, i, moduleSignatures) ||
			isQualifiedProjectMemberQualifier(toks, i, projectMembers)
		) {
			skip.add(i);
		}
		if (word === 'new' && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (word === 'is' && hasEarlierTypeOf(toks, i) && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (isLabelReferenceKeyword(word) && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (word === 'raiseevent' && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (word === 'addressof' && isPotentialVariableReferenceToken(toks[i + 1])) {
			skip.add(i + 1);
		}
		if (isNamedArgumentLabel(toks, i)) {
			skip.add(i);
		}
	}

	return skip;
}

function isQualifiedProjectCallableQualifier(
	toks: readonly VbaToken[],
	index: number,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): boolean {
	if (!isPotentialVariableReferenceToken(toks[index]) || toks[index + 1]?.rawText !== '.') {
		return false;
	}
	if (!isPotentialVariableReferenceToken(toks[index + 2])) {
		return false;
	}
	const qualifier = tokenName(toks[index]);
	const member = tokenName(toks[index + 2]);
	if (!qualifier || !member) {
		return false;
	}
	return moduleSignatures.has(qualifiedProcedureKey(qualifier, member));
}

function isQualifiedProjectMemberQualifier(
	toks: readonly VbaToken[],
	index: number,
	projectMembers: readonly VbaProjectClassMembers[] | undefined,
): boolean {
	if (
		!projectMembers ||
		!isPotentialVariableReferenceToken(toks[index]) ||
		toks[index + 1]?.rawText !== '.'
	) {
		return false;
	}
	if (!isPotentialVariableReferenceToken(toks[index + 2])) {
		return false;
	}
	const qualifier = tokenName(toks[index]);
	const member = tokenName(toks[index + 2]);
	if (!qualifier || !member) {
		return false;
	}
	const qualifierLower = qualifier.toLowerCase();
	const memberLower = member.toLowerCase();
	let surface: VbaProjectClassMembers | undefined;
	for (const candidate of projectMembers) {
		if (candidate.name.toLowerCase() !== qualifierLower) {
			continue;
		}
		if (surface) {
			return false;
		}
		surface = candidate;
	}
	if (!surface) {
		return false;
	}
	if (surface.kind === 'standardModule') {
		return true;
	}
	return surface.members.some((candidate) => candidate.name.toLowerCase() === memberLower);
}

function simpleAssignmentLhsIdentifierIndex(toks: readonly VbaToken[]): number {
	let start = firstExecutableTokenIndex(toks);
	if (tokenText(toks[start]) === 'let' || tokenText(toks[start]) === 'set') {
		start++;
	}
	const eq = topLevelOperatorIndex(toks.slice(start), '=');
	if (eq !== 1) {
		return -1;
	}
	const nameTok = toks[start];
	return nameTok && nameTok.kind === 'identifier' ? start : -1;
}

function isLineLabelOnlyStatement(
	source: string,
	span: Span,
	toks: readonly VbaToken[],
): boolean {
	if (toks.length !== 1 || !isPotentialVariableReferenceToken(toks[0])) {
		return false;
	}
	let j = span.start + toks[0].end;
	while (j < source.length && (source[j] === ' ' || source[j] === '\t')) {
		j++;
	}
	return source[j] === ':';
}

function isPotentialVariableReferenceToken(tok: VbaToken | undefined): boolean {
	return tok?.kind === 'identifier' || tok?.kind === 'bracketedIdentifier';
}

function hasEarlierTypeOf(toks: readonly VbaToken[], before: number): boolean {
	for (let i = 0; i < before; i++) {
		if (tokenText(toks[i]) === 'typeof') {
			return true;
		}
	}
	return false;
}

function isLabelReferenceKeyword(word: string): boolean {
	return word === 'goto' || word === 'gosub' || word === 'resume';
}

function isNamedArgumentLabel(toks: readonly VbaToken[], index: number): boolean {
	if (!isPotentialVariableReferenceToken(toks[index])) {
		return false;
	}
	if (toks[index + 1]?.rawText === ':=') {
		return true;
	}
	return toks[index + 1]?.rawText === ':' && toks[index + 2]?.rawText === '=';
}

function hasOptionExplicit(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): boolean {
	return activeModuleMembers(mod, activity).some(
		(member) =>
			member.kind === 'Option' && /^explicit\b/i.test(member.optionText.trim()),
	);
}

function moduleLevelIdentifierNames(
	symbols: ReturnType<typeof buildModuleSymbols>,
): Set<string> {
	const names = new Set<string>();
	for (const sym of symbols.root.children ?? []) {
		names.add(sym.name.toLowerCase());
		if (sym.kind === 'enum') {
			for (const child of sym.children ?? []) {
				names.add(child.name.toLowerCase());
			}
		}
	}
	return names;
}

/**
 * Rule: a variable declaration cannot include an inline initializer. VBA has no
 * VB.NET-style `Dim x As Long = 1`; the `= value` is a syntax error. `Const`
 * legitimately uses `=` and is skipped. Detection walks every non-Const
 * VariableGroup (module level and inside procedure bodies) and looks for a
 * top-level `=` operator in the group's source slice - a declaration list has no
 * other lawful place for a depth-0 `=`.
 */
function checkDimInitializer(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const inspect = (group: VariableGroupNode): void => {
		if (group.isConst) {
			return; // Const requires `=`; not an error.
		}
		const at = topLevelAssignOffset(source, group.span);
		if (at !== undefined) {
			push(
				'dimInitializer',
				'A variable declaration cannot include an initializer in VBA; assign the value in a separate statement.',
				{ start: at, end: at + 1 },
			);
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspect(member);
		} else if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, inspect, activity);
		}
	}
}

/**
 * Rule: once a declaration's `As <type>` clause is complete, another token in
 * the same logical statement must be introduced by real declaration syntax
 * (`=`, `,`, `:`/newline, etc.). A bare identifier after a complete type name,
 * as in `Dim s As String junk`, is VBE Compile `Syntax error`.
 *
 * This rule is intentionally narrow. It validates the token shape around the
 * `As` clause only; broad unknown type-name resolution belongs to the
 * project-wide binder. Recognized fixed-length String suffixes are consumed by
 * the shared suffix parser before trailing-token detection; their literal size
 * bounds are checked by `checkFixedLengthStringBounds`.
 */
function checkUnexpectedDeclarationTokens(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const inspect = (span: Span, allowEquals: boolean): void => {
		const hit = unexpectedTokenAfterDeclarationType(source, span, allowEquals);
		if (!hit) {
			return;
		}
		push(
			'unexpectedDeclarationToken',
			`Unexpected token '${hit.text}' after a complete declaration type; this will fail to compile as a syntax error.`,
			hit.span,
		);
	};

	const inspectGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			inspect(decl.span, true);
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			for (const field of member.fields) {
				inspectTypeField(field, inspect);
			}
			continue;
		}
		if (member.kind === 'Procedure') {
			for (const param of member.params) {
				inspectParameter(param, inspect);
			}
			forEachVariableGroup(member.body, inspectGroup, activity);
		}
	}
}

const FIXED_LENGTH_STRING_MIN = 1;
const FIXED_LENGTH_STRING_MAX = 65526;

/**
 * Rule: fixed-length String sizes must be in VBE's accepted range when the
 * length is a decimal literal or a same-procedure/module Const/Enum member
 * whose value can be reduced to a deterministic integer expression. Broader
 * constant-expression semantics remain deferred.
 */
function checkFixedLengthStringBounds(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleConstants = collectModuleLiteralIntegerConstants(mod, activity);
	const inspectDeclaration = (
		decl: VariableDeclNode | TypeFieldNode,
		constants: ReadonlyMap<string, number | undefined>,
	): void => {
		if (decl.fixedLength === undefined || isInactiveNode(activity, decl)) {
			return;
		}
		const value = resolveFixedLengthStringSize(decl.fixedLength, constants);
		if (value === undefined) {
			return;
		}
		if (value >= FIXED_LENGTH_STRING_MIN && value <= FIXED_LENGTH_STRING_MAX) {
			return;
		}
		push(
			'fixedLengthStringSize',
			`Fixed-length String size must be between ${FIXED_LENGTH_STRING_MIN} and ${FIXED_LENGTH_STRING_MAX} characters; got ${value}.`,
			fixedLengthStringLengthSpan(source, decl.span) ?? decl.span,
		);
	};

	const inspectGroup = (group: VariableGroupNode): void => {
		for (const decl of group.declarations) {
			inspectDeclaration(decl, moduleConstants);
		}
	};

	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspectGroup(member);
			continue;
		}
		if (member.kind === 'Type') {
			for (const field of member.fields) {
				inspectDeclaration(field, moduleConstants);
			}
			continue;
		}
		if (member.kind === 'Procedure') {
			const procedureConstants = new Map(moduleConstants);
			collectBodyLiteralIntegerConstants(member.body, procedureConstants, activity);
			forEachVariableGroup(
				member.body,
				(group) => {
					for (const decl of group.declarations) {
						inspectDeclaration(decl, procedureConstants);
					}
				},
				activity,
			);
		}
	}
}

function collectModuleLiteralIntegerConstants(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	base: ReadonlyMap<string, number | undefined> = new Map(),
): Map<string, number | undefined> {
	const rawConstants = new Map<string, string | undefined>();
	const seen = new Set<string>();
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup' && member.isConst) {
			addRawIntegerConstants(member, rawConstants, seen);
		} else if (member.kind === 'Enum') {
			addRawEnumIntegerConstants(member, rawConstants, seen);
		}
	}
	const resolved = new Map(base);
	for (const [name, value] of resolveRawIntegerConstants(rawConstants, base)) {
		resolved.set(name, value);
	}
	return resolved;
}

function collectBodyLiteralIntegerConstants(
	body: BodyNode[],
	constants: Map<string, number | undefined>,
	activity: ConditionalActivityTracker | undefined,
): void {
	const rawConstants = new Map<string, string | undefined>();
	const seen = new Set<string>();
	collectBodyRawIntegerConstants(body, rawConstants, activity, seen);
	const resolved = resolveRawIntegerConstants(rawConstants, constants);
	for (const [name, value] of resolved) {
		constants.set(name, value);
	}
}

function collectBodyRawIntegerConstants(
	body: BodyNode[],
	rawConstants: Map<string, string | undefined>,
	activity: ConditionalActivityTracker | undefined,
	seen: Set<string>,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'VariableGroup') {
			if (node.isConst) {
				addRawIntegerConstants(node, rawConstants, seen);
			}
		} else if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			collectBodyRawIntegerConstants((node as { body: BodyNode[] }).body, rawConstants, activity, seen);
		}
	}
}

function addRawIntegerConstants(
	group: VariableGroupNode,
	rawConstants: Map<string, string | undefined>,
	seen: Set<string>,
): void {
	for (const decl of group.declarations) {
		const name = normalizeDeclaredConstantName(decl.name);
		if (!name) {
			continue;
		}
		const key = name.toLowerCase();
		if (seen.has(key)) {
			rawConstants.set(key, undefined);
			continue;
		}
		seen.add(key);
		rawConstants.set(key, decl.defaultRaw);
	}
}

function addRawEnumIntegerConstants(
	en: EnumNode,
	rawConstants: Map<string, string | undefined>,
	seen: Set<string>,
): void {
	let previousName: string | undefined;
	for (const member of en.members) {
		const name = normalizeDeclaredConstantName(member.name);
		if (!name) {
			continue;
		}
		const key = name.toLowerCase();
		if (seen.has(key)) {
			rawConstants.set(key, undefined);
			previousName = name;
			continue;
		}
		seen.add(key);
		rawConstants.set(key, member.valueRaw ?? (previousName ? `${previousName} + 1` : '0'));
		previousName = name;
	}
}

function resolveFixedLengthStringSize(
	raw: string,
	constants: ReadonlyMap<string, number | undefined>,
): number | undefined {
	return evaluateIntegerConstantExpression(raw, constants);
}

function resolveRawIntegerConstants(
	rawConstants: ReadonlyMap<string, string | undefined>,
	base: ReadonlyMap<string, number | undefined>,
): Map<string, number | undefined> {
	const resolved = new Map<string, number | undefined>();
	const resolving = new Set<string>();
	const resolve = (name: string): number | undefined => {
		const key = name.toLowerCase();
		if (resolved.has(key)) {
			return resolved.get(key);
		}
		if (!rawConstants.has(key)) {
			return base.get(key);
		}
		if (resolving.has(key)) {
			resolved.set(key, undefined);
			return undefined;
		}
		const raw = rawConstants.get(key);
		if (raw === undefined) {
			resolved.set(key, undefined);
			return undefined;
		}
		resolving.add(key);
		const value = evaluateIntegerConstantExpression(raw, {
			get: resolve,
		});
		resolving.delete(key);
		resolved.set(key, value);
		return value;
	};
	for (const key of rawConstants.keys()) {
		resolve(key);
	}
	return resolved;
}

interface IntegerConstantLookup {
	get(name: string): number | undefined;
}

function evaluateIntegerConstantExpression(
	raw: string,
	constants: IntegerConstantLookup,
): number | undefined {
	const parser = new IntegerConstantExpressionParser(raw, constants);
	return parser.parse();
}

class IntegerConstantExpressionParser {
	private readonly tokens: VbaToken[];
	private index = 0;

	constructor(
		raw: string,
		private readonly constants: IntegerConstantLookup,
	) {
		this.tokens = tokenize(raw).filter((token) => token.kind !== 'comment' && token.kind !== 'newline');
	}

	parse(): number | undefined {
		if (this.tokens.length === 0) {
			return undefined;
		}
		const value = this.expression();
		if (value === undefined || this.current()) {
			return undefined;
		}
		return value;
	}

	private expression(): number | undefined {
		let value = this.term();
		while (value !== undefined) {
			if (this.accept('+')) {
				const right = this.term();
				value = right === undefined ? undefined : safeInteger(value + right);
				continue;
			}
			if (this.accept('-')) {
				const right = this.term();
				value = right === undefined ? undefined : safeInteger(value - right);
				continue;
			}
			break;
		}
		return value;
	}

	private term(): number | undefined {
		let value = this.factor();
		while (value !== undefined) {
			if (!this.accept('*')) {
				break;
			}
			const right = this.factor();
			value = right === undefined ? undefined : safeInteger(value * right);
		}
		return value;
	}

	private factor(): number | undefined {
		if (this.accept('+')) {
			return this.factor();
		}
		if (this.accept('-')) {
			const value = this.factor();
			return value === undefined ? undefined : safeInteger(-value);
		}
		if (this.accept('(')) {
			const value = this.expression();
			return value !== undefined && this.accept(')') ? value : undefined;
		}
		const token = this.current();
		if (!token) {
			return undefined;
		}
		if (token.kind === 'integerLiteral') {
			this.index++;
			return parseVbaIntegerLiteral(token.rawText);
		}
		const qualifiedName = this.qualifiedConstantName();
		if (qualifiedName) {
			return this.constants.get(qualifiedName.toLowerCase());
		}
		const name = normalizeFixedLengthConstantName(token.rawText);
		if (name) {
			this.index++;
			return this.constants.get(name.toLowerCase());
		}
		return undefined;
	}

	private qualifiedConstantName(): string | undefined {
		const qualifier = this.current();
		const dot = this.tokens[this.index + 1];
		const member = this.tokens[this.index + 2];
		const qualifierName = qualifier ? normalizeFixedLengthConstantName(qualifier.rawText) : undefined;
		const memberName = member ? normalizeFixedLengthConstantName(member.rawText) : undefined;
		if (!qualifierName || dot?.rawText !== '.' || !memberName) {
			return undefined;
		}
		this.index += 3;
		return `${qualifierName}.${memberName}`;
	}

	private current(): VbaToken | undefined {
		return this.tokens[this.index];
	}

	private accept(raw: string): boolean {
		if (this.current()?.rawText !== raw) {
			return false;
		}
		this.index++;
		return true;
	}
}

function parseDecimalIntegerLiteral(raw: string): number | undefined {
	if (!/^\d+$/.test(raw)) {
		return undefined;
	}
	const value = Number(raw);
	return Number.isSafeInteger(value) ? value : undefined;
}

function parseVbaIntegerLiteral(raw: string): number | undefined {
	const text = raw.trim().replace(/[%&^]$/, '');
	const hex = /^&[hH]([0-9A-Fa-f]+)$/.exec(text);
	if (hex) {
		const value = Number.parseInt(hex[1], 16);
		return Number.isSafeInteger(value) ? value : undefined;
	}
	const octal = /^&[oO]([0-7]+)$/.exec(text);
	if (octal) {
		const value = Number.parseInt(octal[1], 8);
		return Number.isSafeInteger(value) ? value : undefined;
	}
	return parseDecimalIntegerLiteral(text);
}

function safeInteger(value: number): number | undefined {
	return Number.isSafeInteger(value) ? value : undefined;
}

function normalizeDeclaredConstantName(raw: string): string | undefined {
	const text = raw.trim();
	return text.length > 0 ? text : undefined;
}

function normalizeFixedLengthConstantName(raw: string): string | undefined {
	const text = raw.trim();
	if (/^\[[^\]]+\]$/.test(text)) {
		return text.slice(1, -1);
	}
	return VBA_IDENTIFIER_NAME_RE.test(text) ? text : undefined;
}

function fixedLengthStringLengthSpan(source: string, span: Span): Span | undefined {
	const toks = statementTokens(source, span);
	const asIndex = toks.findIndex((t) => tokenText(t) === 'as');
	if (asIndex < 0) {
		return undefined;
	}
	let typeStart = asIndex + 1;
	if (tokenText(toks[typeStart]) === 'new') {
		typeStart++;
	}
	const fixed = parseFixedLengthStringType(toks, typeStart);
	const token = fixed ? toks[fixed.lengthIndex] : undefined;
	return token ? absoluteSpan(span, token) : undefined;
}

function inspectTypeField(
	field: TypeFieldNode,
	inspect: (span: Span, allowEquals: boolean) => void,
): void {
	inspect(field.span, false);
}

function inspectParameter(
	param: ParameterNode,
	inspect: (span: Span, allowEquals: boolean) => void,
): void {
	inspect(param.span, true);
}

/** Walks every VariableGroupNode in a body, descending into nested blocks. */
function forEachVariableGroup(
	body: BodyNode[],
	visit: (group: VariableGroupNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'VariableGroup') {
			visit(node);
		} else if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			forEachVariableGroup((node as { body: BodyNode[] }).body, visit, activity);
		}
	}
}

/** Walks every generic StatementNode in a procedure body, descending into nested blocks. */
function forEachBodyStatement(
	body: BodyNode[],
	visit: (statement: StatementNode) => void,
	activity?: ConditionalActivityTracker,
): void {
	for (const node of body) {
		if (isInactiveNode(activity, node)) {
			continue;
		}
		if (node.kind === 'Statement') {
			visit(node);
		} else if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			forEachBodyStatement((node as { body: BodyNode[] }).body, visit, activity);
		}
	}
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

function forEachProcedureBodyLine(
	source: string,
	procedure: ProcedureNode,
	visit: (span: Span) => void,
): void {
	const firstBreak = firstLineBreakAtOrAfter(source, procedure.span.start);
	if (firstBreak < 0 || firstBreak >= procedure.span.end) {
		return;
	}
	let lineStart = nextLineStart(source, firstBreak);
	while (lineStart < procedure.span.end) {
		let lineEnd = lineStart;
		while (lineEnd < procedure.span.end && source[lineEnd] !== '\r' && source[lineEnd] !== '\n') {
			lineEnd++;
		}
		visit({ start: lineStart, end: lineEnd });
		lineStart = nextLineStart(source, lineEnd);
	}
}

function nextLineStart(source: string, lineBreakOffset: number): number {
	if (
		source[lineBreakOffset] === '\r' &&
		lineBreakOffset + 1 < source.length &&
		source[lineBreakOffset + 1] === '\n'
	) {
		return lineBreakOffset + 2;
	}
	return lineBreakOffset + 1;
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

function unexpectedTokenAfterDeclarationType(
	source: string,
	span: Span,
	allowEquals: boolean,
): { text: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	const asIndex = toks.findIndex((t) => tokenText(t) === 'as');
	if (asIndex < 0) {
		return undefined;
	}

	let i = asIndex + 1;
	if (tokenText(toks[i]) === 'new') {
		i++;
	}

	const typeStart = i;
	i = consumeDeclarationTypeName(toks, i);
	if (i === typeStart) {
		return undefined;
	}

	const fixedLengthString = parseFixedLengthStringType(toks, typeStart);
	if (fixedLengthString && fixedLengthString.endIndex > i) {
		i = fixedLengthString.endIndex;
	}

	const next = toks[i];
	if (!next) {
		return undefined;
	}
	if (allowEquals && next.kind === 'operator' && next.rawText === '=') {
		return undefined;
	}

	return {
		text: next.rawText,
		span: absoluteSpan(span, next),
	};
}

function consumeDeclarationTypeName(toks: VbaToken[], start: number): number {
	if (!isDeclarationTypeNameToken(toks[start])) {
		return start;
	}
	let i = start + 1;
	for (;;) {
		if (toks[i]?.rawText !== '.') {
			return i;
		}
		if (!isDeclarationTypeNameToken(toks[i + 1])) {
			return start;
		}
		i += 2;
	}
}

function isDeclarationTypeNameToken(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return false;
	}
	return (
		tok.kind === 'identifier' ||
		tok.kind === 'keyword' ||
		tok.kind === 'bracketedIdentifier'
	);
}

function checkInvalidAsTypeNames(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	opts: AnalyzeModuleOptions,
	push: PushFn,
): void {
	const withEventsNewDeclarationSpans = collectWithEventsNewDeclarationSpans(mod, activity);
	for (const ref of collectTypeNameReferences(source)) {
		if (activity?.isInactive(ref.span)) {
			continue;
		}
		const lookupName = typeReferenceLookupName(ref);
		const resolved = resolveTypeName(lookupName, {
			projectTypes: opts.projectTypes,
			model: opts.hostModel,
		});
		if (resolved?.kind === 'ambiguous') {
			push(
				'invalidAsTypeName',
				`'${ref.name}' is ambiguous because multiple visible project types use that name.`,
				ref.span,
			);
			continue;
		}
		if (
			resolved &&
			isNewTypeReference(ref.kind) &&
			!isCreatableTypeCompletion(resolved)
		) {
			if (
				ref.kind === 'newDeclaration' &&
				withEventsNewDeclarationSpans.some((span) => containsSpan(span, ref.span))
			) {
				continue;
			}
			push(
				'invalidNewTypeName',
				`'${ref.name}' is ${typeKindLabelForNew(resolved.kind)} and cannot be used with New. New can create project classes and UserForms only.`,
				ref.span,
			);
			continue;
		}
		if (resolved) {
			continue;
		}
		if (isReservedIdentifier(ref.name)) {
			push(
				'invalidAsTypeName',
				`'${ref.name}' is a reserved VBA identifier, not a valid type name.`,
				ref.span,
			);
			continue;
		}
		if (resolveRuntimeFunction(ref.name)) {
			push(
				'invalidAsTypeName',
				`'${ref.name}' is a VBA runtime function, not a valid type name.`,
				ref.span,
			);
			continue;
		}
		if (opts.knownNonTypeNames?.has(ref.name.toLowerCase())) {
			push(
				'invalidAsTypeName',
				`'${ref.name}' resolves to a project declaration, but that declaration is not a type.`,
				ref.span,
			);
			continue;
		}
	}
}

function collectWithEventsNewDeclarationSpans(
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
): Span[] {
	const spans: Span[] = [];
	const inspect = (group: VariableGroupNode): void => {
		if (!group.withEvents || isInactiveNode(activity, group)) {
			return;
		}
		for (const decl of group.declarations) {
			if (decl.isNew) {
				spans.push(decl.span);
			}
		}
	};
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'VariableGroup') {
			inspect(member);
			continue;
		}
		if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, inspect, activity);
		}
	}
	return spans;
}

function containsSpan(container: Span, inner: Span): boolean {
	return inner.start >= container.start && inner.end <= container.end;
}

function declaredNameSpan(source: string, span: Span, name: string): Span {
	const lower = name.toLowerCase();
	for (const tok of statementTokens(source, span)) {
		if (tokenName(tok)?.toLowerCase() === lower) {
			return absoluteSpan(span, tok);
		}
	}
	return span;
}

function firstTokenSpan(source: string, span: Span): Span {
	const tok = statementTokens(source, span)[0];
	return tok ? absoluteSpan(span, tok) : span;
}

function keywordSpan(source: string, span: Span, ...keywords: string[]): Span {
	const expected = new Set(keywords);
	const tok = statementTokensAfterLeadingLabel(source, span)
		.find((token) => expected.has(tokenText(token)));
	return tok ? absoluteSpan(span, tok) : firstTokenSpan(source, span);
}

/**
 * Returns the absolute offset of the first top-level `=` operator in the source
 * slice for `span`, or undefined. Parenthesised regions (array bounds, default
 * sub-expressions) are skipped so only a declaration-level `=` is reported.
 */
function topLevelAssignOffset(source: string, span: Span): number | undefined {
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	let depth = 0;
	for (const t of toks) {
		const r = t.rawText;
		if (r === '(') {
			depth++;
		} else if (r === ')') {
			depth--;
		} else if (depth === 0 && t.kind === 'operator' && r === '=') {
			return span.start + t.start;
		}
	}
	return undefined;
}

/**
 * Rule: parameter-list constraints. A required parameter may not follow an
 * `Optional` one, `ParamArray` must be the final parameter, `ParamArray` cannot
 * be combined with Optional parameters in the same list, and explicitly typed
 * `ParamArray` elements must be Variant. These are read straight off the parsed
 * parameter flags, so they are deterministic.
 */
function checkParameterOrder(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const params = member.params;
		const hasOptional = params.some((p) => p.optional);
		let optionalSeen = false;
		for (let i = 0; i < params.length; i++) {
			const p = params[i];
			if (p.paramArray) {
				if (p.asType && normalizeType(p.asType) !== 'variant') {
					push(
						'paramArrayNonVariant',
						`ParamArray '${p.name}' elements must be Variant, but this parameter is declared As ${p.asType}.`,
						declaredNameSpan(source, p.span, p.name),
					);
				}
				if (hasOptional) {
					push(
						'paramArrayWithOptional',
						`ParamArray '${p.name}' cannot be used in the same parameter list as Optional arguments.`,
						declaredNameSpan(source, p.span, p.name),
					);
				}
				if (i !== params.length - 1) {
					push(
						'paramArrayNotLast',
						`ParamArray '${p.name}' must be the last parameter.`,
						declaredNameSpan(source, p.span, p.name),
					);
				}
				continue;
			}
			if (p.optional) {
				optionalSeen = true;
				continue;
			}
			if (optionalSeen) {
				push(
					'requiredParamAfterOptional',
					`Parameter '${p.name}' must be Optional because it follows an Optional parameter.`,
					declaredNameSpan(source, p.span, p.name),
				);
			}
		}
	}
}

/**
 * Rule: Optional parameter defaults must be compile-time compatible with their
 * declared type when the default expression is deterministic. VBE oracle
 * evidence rejects nonnumeric string defaults for numeric and Boolean
 * parameters as compile-time Type mismatch, while numeric strings remain valid.
 */
function checkParameterDefaultValues(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		for (const param of member.params) {
			if (!param.defaultRaw || !param.asType) {
				continue;
			}
			const defaultTokens = parameterDefaultTokens(source, param);
			if (!defaultTokens) {
				continue;
			}
			const actual = inferArgumentType(defaultTokens.tokens, param.span.start, new Map(), new Map());
			if (!actual) {
				continue;
			}
			const reason = parameterDefaultIncompatibilityReason(param.asType, actual);
			if (!reason) {
				continue;
			}
			push(
				'parameterDefaultTypeMismatch',
				`Optional parameter '${param.name}' expects ${param.asType}, but its default value is ${actual.label}. ${reason}`,
				defaultTokens.span,
			);
		}
	}
}

function parameterDefaultTokens(
	source: string,
	param: ParameterNode,
): { tokens: VbaToken[]; span: Span } | undefined {
	const toks = tokenize(source.slice(param.span.start, param.span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	const eq = topLevelOperatorIndex(toks, '=');
	if (eq < 0 || eq + 1 >= toks.length) {
		return undefined;
	}
	const tokens = toks.slice(eq + 1);
	return {
		tokens,
		span: spanForTokens(tokens, param.span.start),
	};
}

function parameterDefaultIncompatibilityReason(
	expectedRaw: string,
	actual: InferredArgumentType,
): string | undefined {
	const reason = incompatibilityReason(expectedRaw, actual);
	if (!reason || !/string literal/i.test(actual.label)) {
		return undefined;
	}
	return 'This is a VBE compile error: Type mismatch.';
}

function isNewTypeReference(kind: TypeNameReferenceKind): boolean {
	return kind === 'newExpression' || kind === 'newDeclaration';
}

function typeKindLabelForNew(kind: TypeCompletionKind): string {
	switch (kind) {
		case 'primitive':
			return 'a VBA primitive type';
		case 'external':
			return 'an external interface type';
		case 'host':
			return 'an Excel object-model type';
		case 'document':
			return 'a document module type';
		case 'enum':
			return 'an Enum type';
		case 'userType':
			return 'a user-defined Type';
		case 'ambiguous':
			return 'an ambiguous project type';
		case 'module':
			return 'a module qualifier';
		case 'class':
		case 'userform':
			return 'a creatable project type';
	}
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
	memberCtx: MemberCompletionContext,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const sourceNames = sourceNameScopeFor(symbols, member);
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
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			const incompleteMember = incompleteMemberAccess(source, stmt.span, {
				scalarTypes: env,
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
		if (receiverName && options.scalarTypes) {
			const normalized = normalizeType(options.scalarTypes.get(receiverName.toLowerCase()));
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
	projectIntegerConstants: ReadonlyMap<string, string | undefined> | undefined,
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
		forEachStatement(member.body, (stmt) => {
			for (const hit of divisionByZeroDivisors(source, stmt.span, procedureConstants)) {
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
	constants: ReadonlyMap<string, number | undefined>,
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
	constants: ReadonlyMap<string, number | undefined>,
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
	constants: ReadonlyMap<string, number | undefined>,
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
	constants: ReadonlyMap<string, number | undefined>,
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
	constants: ReadonlyMap<string, number | undefined>,
): boolean {
	if (isZeroNumericLiteral(tok)) {
		return true;
	}
	const name = tok ? tokenName(tok) : undefined;
	return name !== undefined && constants.get(name.toLowerCase()) === 0;
}

function foldIntegerExpressionTokens(
	source: string,
	span: Span,
	toks: VbaToken[],
	start: number,
	endExclusive: number,
	constants: ReadonlyMap<string, number | undefined>,
): number | undefined {
	if (start >= endExclusive) {
		return undefined;
	}
	const first = toks[start];
	const last = toks[endExclusive - 1];
	const raw = source.slice(span.start + first.start, span.start + last.end);
	return evaluateIntegerConstantExpression(raw, { get: (name) => constants.get(name.toLowerCase()) });
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
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	const functions = expressionCallableFunctionNames(symbols, projectProcedures);
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const sourceNames = sourceNameScopeFor(symbols, member);
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

function tokenName(tok: VbaToken): string | undefined {
	if (tok.kind === 'identifier' || tok.kind === 'keyword') {
		return tok.rawText;
	}
	if (tok.kind === 'bracketedIdentifier') {
		return stripHeaderBrackets(tok.rawText);
	}
	return undefined;
}

function topLevelOperatorIndex(toks: readonly VbaToken[], operator: string): number {
	let depth = 0;
	for (let i = 0; i < toks.length; i++) {
		const raw = toks[i].rawText;
		if (raw === '(' || raw === '[') {
			depth++;
		} else if (raw === ')' || raw === ']') {
			depth--;
		} else if (depth === 0 && toks[i].kind === 'operator' && raw === operator) {
			return i;
		}
	}
	return -1;
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

function callableAcceptsZeroArguments(sig: CallableTypeSignature): boolean {
	return sig.params.every((param) => param.optional || param.paramArray);
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
function matchParenFrom(toks: readonly VbaToken[], open: number): number {
	let depth = 0;
	for (let k = open; k < toks.length; k++) {
		const r = toks[k].rawText;
		if (r === '(') {
			depth++;
		} else if (r === ')') {
			depth--;
			if (depth === 0) {
				return k;
			}
		}
	}
	return -1;
}

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
			checkContextStatement(source, member, root, push);
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

function statementTokens(source: string, span: Span): VbaToken[] {
	return tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
}

function statementTokensAfterLeadingLabel(source: string, span: Span): VbaToken[] {
	const toks = statementTokens(source, span);
	const firstExecutable = firstExecutableTokenIndex(toks);
	return firstExecutable > 0 ? toks.slice(firstExecutable) : toks;
}

function firstExecutableTokenIndex(toks: readonly VbaToken[]): number {
	if (toks.length > 1 && toks[0].kind === 'integerLiteral' && /^\d+$/.test(toks[0].rawText)) {
		return 1;
	}
	if (
		toks.length > 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		return 2;
	}
	return 0;
}

function tokenText(token: VbaToken | undefined): string {
	return (token?.canonicalText ?? token?.rawText ?? '').toLowerCase();
}

function absoluteSpan(base: Span, token: VbaToken): Span {
	return { start: base.start + token.start, end: base.start + token.end };
}

function physicalLineSpanAtOffset(source: string, offset: number): Span {
	const safe = Math.max(0, Math.min(offset, source.length));
	const before = source.lastIndexOf('\n', Math.max(0, safe - 1));
	const start = before < 0 ? 0 : before + 1;
	const after = source.indexOf('\n', safe);
	let end = after < 0 ? source.length : after;
	if (end > start && source[end - 1] === '\r') {
		end--;
	}
	return { start, end };
}

function exitPhraseSpan(base: Span, first: VbaToken, target: VbaToken): Span {
	return { start: base.start + first.start, end: base.start + target.end };
}

/**
 * Rule: `Option` statements must precede every declaration and procedure (only
 * `Attribute` lines may come before them in an exported module). Once a real
 * declaration has appeared, any later `Option` is misplaced.
 */
function checkOptionPlacement(
	source: string,
	mod: ModuleNode,
	activity: ConditionalActivityTracker | undefined,
	push: PushFn,
): void {
	let declarationSeen = false;
	for (const member of activeModuleMembers(mod, activity)) {
		if (member.kind === 'Attribute') {
			continue;
		}
		if (member.kind === 'Option') {
			if (declarationSeen) {
				push(
					'optionAfterDeclaration',
					'Option statements must appear before any declaration or procedure.',
					firstTokenSpan(source, member.span),
				);
			}
			continue;
		}
		declarationSeen = true;
	}
}
