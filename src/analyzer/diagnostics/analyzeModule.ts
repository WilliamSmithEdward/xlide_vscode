// Active VBA diagnostics engine (MS-VBAL Phase 5).
//
// `analyzeModule` runs the high-confidence semantic rules from the rule
// catalogue over one module's source and returns offset-based diagnostics. It
// is pure (no `vscode`): the editor layer converts spans to ranges and severity
// names to the VS Code enum. Structural block-balance checking stays in
// `src/vbaLinter.ts` (lintVbaSource); this engine adds the semantic rules on
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
	type VbaRuntimeFunction,
} from '../runtime/vbaRuntime';
import { STATEMENT_KEYWORDS } from '../signature/signatureHelp';
import type {
	BodyNode,
	ModuleNode,
	ParameterNode,
	ProcedureNode,
	Span,
	StatementNode,
	TypeFieldNode,
	VariableGroupNode,
} from '../parser/nodes';
import { parseModule } from '../parser/parseModule';
import { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type {
	ModuleSymbolKind,
	VbaProjectClassMembers,
	VbaProcedureSignature,
	VbaSymbol,
} from '../symbols/symbolModel';
import { isProcedureKind, qualifiedProcedureKey } from '../symbols/symbolModel';
import {
	resolveMemberCompletions,
	type MemberCompletion,
	type MemberCompletionContext,
} from '../completion/memberAccess';
import { resolveTypeName, type ProjectTypeName } from '../completion/typeCompletion';
import {
	eventHandlerDocumentTypeForContext,
	eventHandlerProcedureForName,
	type EventHandlerDocumentType,
} from '../completion/eventHandlers';
import { collectTypeNameReferences } from '../semantic/typeSemanticTokens';
import {
	DIAGNOSTIC_RULES,
	DiagnosticRuleName,
	DiagnosticSeverity,
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
}

/** Per-rule severity overrides; `'off'` disables a rule. */
export type SeverityOverrides = Partial<
	Record<DiagnosticRuleName, DiagnosticSeverity | 'off'>
>;

/** Inputs for {@link analyzeModule}. */
export interface AnalyzeModuleOptions {
	/** VB component name (used only for symbol container labels). */
	moduleName?: string;
	/** Workbook-project role of the module. */
	moduleKind?: ModuleSymbolKind;
	/** Workbook/document subtype for Excel document modules when known. */
	documentType?: EventHandlerDocumentType;
	/** Optional per-rule severity overrides (e.g. Option Explicit severity). */
	severities?: SeverityOverrides;
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
	 * Exported Sub/Function signatures across the workbook project, grouped by
	 * lowercased procedure name. When omitted, type and arity validation remain
	 * single-module only.
	 */
	projectProcedures?: ReadonlyMap<string, readonly VbaProcedureSignature[]>;
	/** Source-declared workbook class/UserForm/document members visible to this module. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/** Source-declared workbook type names visible to this module. */
	projectTypes?: readonly ProjectTypeName[];
	/** Lowercased visible declaration names known not to be type names. */
	knownNonTypeNames?: ReadonlySet<string>;
	/** Host object model metadata. Defaults to Excel's curated non-exhaustive model. */
	hostModel?: HostObjectModel;
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
	overrides: SeverityOverrides | undefined,
): DiagnosticSeverity | undefined {
	const override = overrides?.[rule];
	if (override === 'off') {
		return undefined;
	}
	return override ?? DIAGNOSTIC_RULES[rule].defaultSeverity;
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

function runRules(
	source: string,
	opts: AnalyzeModuleOptions,
): VbaDiagnostic[] {
	const out: VbaDiagnostic[] = [];
	const moduleName = opts.moduleName ?? 'Module';
	const moduleKind = opts.moduleKind ?? 'standard';
	const overrides = opts.severities;

	const push = (
		rule: DiagnosticRuleName,
		message: string,
		span: Span,
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
		});
	};

	const mod = parseModule(source);
	const symbols = buildModuleSymbols(moduleName, moduleKind, source);
	const memberCtx = diagnosticMemberCompletionContext(opts);

	checkUnterminatedStrings(source, push);
	checkDuplicateProcedures(symbols.root.children ?? [], push);
	checkDuplicateDeclarations(symbols.root.children ?? [], push);
	checkDuplicateModuleMembers(symbols.root.children ?? [], push);
	checkConstAssignment(source, mod, symbols, push);
	checkOptionExplicit(source, mod, push);
	checkUndeclaredVariables(source, mod, symbols, opts.knownIdentifiers, push);
	checkOptionPlacement(mod, push);
	checkProcedureHeader(source, mod, push);
	checkReservedDeclarationNames(source, mod, push);
	checkParameterOrder(mod, push);
	checkUnbalancedParens(source, push);
	checkInvalidExpressionSyntax(source, mod, push);
	checkDimInitializer(source, mod, push);
	checkUnexpectedDeclarationTokens(source, mod, push);
	checkObjectModulePublicMembers(source, mod, moduleKind, push);
	checkEventHandlerModuleScope(source, mod, moduleName, moduleKind, opts.documentType, push);
	checkInvalidAsTypeNames(source, opts, push);
	checkCallParens(source, mod, opts.projectProcedures, push);
	checkExpressionCallParens(source, mod, push);
	checkSetAssignments(source, mod, symbols, memberCtx, push);
	checkExitStatements(source, mod, push);
	checkStatementContext(source, mod, push);
	checkScalarMemberAccess(source, mod, symbols, push);
	checkMemberNotFound(source, mod, memberCtx, push);
	checkNonCallableCallStatement(source, mod, symbols, push);
	checkArgumentCount(
		source,
		mod,
		opts.projectProcedures,
		memberCtx,
		push,
	);
	checkArgumentTypes(
		source,
		mod,
		symbols,
		opts.projectProcedures,
		memberCtx,
		push,
	);
	checkAssignmentTypes(source, mod, symbols, memberCtx, push);
	checkMissingReturnAssignments(source, mod, push);
	if (opts.knownProcedures) {
		checkUnknownCallStatement(source, mod, symbols, opts.knownProcedures, push);
	}

	return out;
}

type PushFn = (rule: DiagnosticRuleName, message: string, span: Span) => void;

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
	push: PushFn,
): void {
	const moduleConsts = new Set<string>();
	for (const sym of symbols.root.children ?? []) {
		if (sym.kind === 'constant') {
			moduleConsts.add(sym.name.toLowerCase());
		}
	}

	for (const member of mod.members) {
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
): void {
	for (const node of body) {
		if (node.kind === 'Statement') {
			visit(node);
		} else if ('body' in node && Array.isArray(node.body)) {
			forEachStatement(node.body, visit);
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
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	let i = 0;
	// Skip a leading line label: `Name:`.
	if (
		toks.length >= 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		i = 2;
	}
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
	let i = 0;
	if (
		toks.length >= 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		i = 2;
	}
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
	if (lhs.length < 3) {
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
	push: PushFn,
): void {
	for (const member of mod.members) {
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
		});
	}
}

function memberAccessReferences(
	source: string,
	span: Span,
): { member: string; memberSpan: Span; dotEndOffset: number }[] {
	const toks = statementTokens(source, span);
	const out: { member: string; memberSpan: Span; dotEndOffset: number }[] = [];
	for (let i = 1; i < toks.length - 1; i++) {
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
	const members = resolveMemberCompletions(source, dotEndOffset, memberCtx);
	if (members.length === 0) {
		return undefined;
	}
	const owner = members[0].owner;
	if (!members[0].surfaceExhaustive) {
		return undefined;
	}
	return { owner, members };
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
	knownProcedures: ReadonlySet<string>,
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

	const isKnown = (name: string, locals: ReadonlySet<string>): boolean => {
		const lower = name.toLowerCase();
		return (
			knownProcedures.has(lower) ||
			moduleNames.has(lower) ||
			locals.has(lower) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	for (const member of mod.members) {
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
				push(
					'unknownCallStatement',
					`Sub or Function not defined: '${hit.name}'.`,
					hit.span,
				);
			}
		});
	}
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
	push: PushFn,
): void {
	const moduleNonCallables = moduleNonCallableSymbols(symbols);
	for (const member of mod.members) {
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
			const target = localNonCallables.get(lower) ?? moduleNonCallables.get(lower);
			if (!target) {
				return;
			}
			push(
				'nonCallableCallStatement',
				`Cannot call '${call.name}' because it resolves to ${symbolKindLabel(target)}, not a Sub or Function.`,
				call.nameSpan,
			);
		});
	}
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

/**
 * If the statement spanning `span` is a *call statement* whose callee is a bare
 * (non-member) identifier, returns that identifier and its absolute span;
 * otherwise undefined. Three forms qualify:
 *
 *   - a lone identifier (`DoStartup`),
 *   - a parenless call with arguments (`MsgBox "hi"`, `Foo 1, 2`), and
 *   - an explicit `Call` statement (`Call DoWork`, `Call Foo(1, 2)`).
 *
 * Disqualified (returns undefined): assignments (a top-level `=`), member calls
 * (a leading `receiver.`), line labels (`done:`), a statement keyword leader
 * (`Set`, `Open`, `Print`, ...), and the parenthesized/indexed form
 * `Name(...)...` for a non-`Call` statement - the latter is excluded because a
 * bare `Cells(1, 1)` or `Range("A1")` is an implicit Application member and must
 * never be flagged. `Call Name(...)` is still inspected because `Call` makes the
 * call unambiguous.
 */
function callStatementTarget(
	source: string,
	span: Span,
): { name: string; span: Span } | undefined {
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	if (toks.length === 0) {
		return undefined;
	}

	let idx = 0;
	const explicitCall = toks[0].rawText.toLowerCase() === 'call';
	if (explicitCall) {
		idx = 1;
	}

	const callee = toks[idx];
	if (!callee || callee.kind !== 'identifier') {
		return undefined;
	}
	if (STATEMENT_KEYWORDS.has(callee.rawText.toLowerCase())) {
		return undefined;
	}

	const result = {
		name: callee.rawText,
		span: { start: span.start + callee.start, end: span.start + callee.end },
	};

	const next = toks[idx + 1];
	if (!next) {
		// A lone identifier (optionally `Call name`). For the non-Call form, an
		// identifier immediately followed by `:` is a line label, not a call; the
		// statement span excludes the colon, so peek past it in the full source.
		if (!explicitCall) {
			let j = span.start + callee.end;
			while (j < source.length && (source[j] === ' ' || source[j] === '\t')) {
				j++;
			}
			if (source[j] === ':') {
				return undefined;
			}
		}
		return result;
	}

	const r = next.rawText;
	if (r === '.' || r === ':') {
		return undefined; // member call or label/separator artifact
	}
	if (!explicitCall && r === '(') {
		return undefined; // index/group/implicit host member - defer
	}
	if (!explicitCall) {
		// The first argument must be separated from the callee by whitespace; this
		// is what makes it a parenless call rather than e.g. `a=b` glued together.
		const gap = source.slice(span.start + callee.end, span.start + next.start);
		if (!/\s/.test(gap)) {
			return undefined;
		}
	}
	// A top-level `=` (outside any parentheses/brackets) makes this an assignment.
	let depth = 0;
	for (let k = idx + 1; k < toks.length; k++) {
		const tr = toks[k].rawText;
		if (tr === '(' || tr === '[') {
			depth++;
		} else if (tr === ')' || tr === ']') {
			depth--;
		} else if (depth === 0 && tr === '=') {
			return undefined;
		}
	}
	return result;
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
	push: PushFn,
): void {
	for (const member of mod.members) {
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
			continue; // malformed in a way the structural linter already reports
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

function checkReservedDeclarationNames(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	const report = (kind: string, hit: NameTokenHit | undefined): void => {
		if (!hit || hit.bracketed || !isReservedIdentifier(hit.name)) {
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

	for (const member of mod.members) {
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
		forEachVariableGroup(member.body, inspectVariableGroup);
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
 * Rule: a call to a known Sub/Function must supply an argument count the
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
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const procsByName = new Map<string, ProcedureNode[]>();
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		if (member.procKind !== 'Sub' && member.procKind !== 'Function') {
			continue;
		}
		const key = member.name.toLowerCase();
		const arr = procsByName.get(key);
		if (arr) {
			arr.push(member);
		} else {
			procsByName.set(key, [member]);
		}
	}
	const projectSignatures = uniqueProjectTypeSignatures(projectProcedures);
	const moduleSignatures = callableTypeSignaturesFor(mod, projectProcedures);
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			const statementCall = extractCall(source, stmt.span);
			const qualifiedStatementCall = statementCall
				? undefined
				: extractQualifiedCall(source, stmt.span, moduleSignatures);
			const effectiveStatementCall = statementCall ?? qualifiedStatementCall;
			if (effectiveStatementCall) {
				validateCallableArity(effectiveStatementCall, procsByName, projectSignatures, push);
			}
			for (const call of expressionCalls(source, stmt.span, moduleSignatures)) {
				if (sameCallTarget(call, effectiveStatementCall)) {
					continue;
				}
				validateCallableArity(call, procsByName, projectSignatures, push);
			}
			for (const memberCall of memberExpressionCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				validateArity(memberCall.signature, memberCall.call, push);
			}
			for (const memberCall of memberStatementCalls(
				source,
				stmt.span,
				memberCtx,
			)) {
				validateArity(memberCall.signature, memberCall.call, push);
			}
		});
	}
}

function validateCallableArity(
	call: CallArguments,
	procsByName: ReadonlyMap<string, ProcedureNode[]>,
	projectSignatures: ReadonlyMap<string, CallableTypeSignature>,
	push: PushFn,
): void {
	const lower = call.lookupKey ?? call.name.toLowerCase();
	const candidates = call.qualifier ? undefined : procsByName.get(call.name.toLowerCase());
	if (candidates) {
		// Skip ambiguous same-module targets where the signature is not unique.
		if (candidates.length === 1) {
			validateArity(procedureTypeSignature(candidates[0]), call, push);
		}
		return;
	}
	const projectSignature = projectSignatures.get(lower);
	if (projectSignature) {
		validateArity(projectSignature, call, push);
		return;
	}
	if (!call.qualifier) {
		const runtime = resolveRuntimeFunction(call.name);
		const runtimeSignature = runtime ? runtimeAritySignature(runtime) : undefined;
		if (runtimeSignature) {
			validateArity(runtimeSignature, call, push);
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
	const relCalleeStart = hit.span.start - sliceStart;
	const calleeIdx = toks.findIndex((t) => t.start === relCalleeStart);
	if (calleeIdx < 0) {
		return undefined;
	}

	const explicitCall = toks[0].rawText.toLowerCase() === 'call';
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

	let qualifierIdx = 0;
	const explicitCall = toks[0].rawText.toLowerCase() === 'call';
	if (explicitCall) {
		qualifierIdx = 1;
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
	const finishSlot = (): void => {
		const slot = slots[slots.length - 1];
		spans.push(argumentSlotSpan(slot, emptyMarker, sliceStart));
	};
	for (const t of toks) {
		if (t.kind === 'punctuation' && t.rawText === '(') {
			depth++;
		} else if (t.kind === 'punctuation' && t.rawText === ')') {
			depth--;
		}
		if (t.kind === 'punctuation' && t.rawText === ',' && depth === 0) {
			finishSlot();
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

function argumentSlotSpan(slot: VbaToken[], emptyMarker: VbaToken | undefined, sliceStart: number): Span {
	if (slot.length > 0) {
		return {
			start: sliceStart + slot[0].start,
			end: sliceStart + slot[slot.length - 1].end,
		};
	}
	if (emptyMarker) {
		return { start: sliceStart + emptyMarker.start, end: sliceStart + emptyMarker.end };
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
	sig: CallableTypeSignature,
	call: CallArguments,
	push: PushFn,
): void {
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
					`Named argument not found: '${raw}' is not a parameter of '${sig.name}'.`,
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
			push(
				'argumentCount',
				`Argument not optional: '${name}' is required by '${sig.name}'.`,
				call.slotSpans?.[i] ?? call.nameSpan,
			);
		}
	}

	const n = call.slots.length;
	if (n < required || n > max) {
		push(
			'argumentCount',
			`Wrong number of arguments to '${sig.name}': expected ${describeArity(required, max)}, but got ${n}.`,
			call.nameSpan,
		);
	}
}

interface CallableParamType {
	name: string;
	type?: string;
	optional: boolean;
	paramArray: boolean;
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
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(mod, projectProcedures);
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
		forEachStatement(member.body, (stmt) => {
			for (const call of expressionCalls(source, stmt.span, moduleSignatures)) {
				validateArgumentTypes(call, env, moduleSignatures, push);
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
					push,
				);
			}
			const statementCall = extractCall(source, stmt.span);
			const qualifiedStatementCall = statementCall
				? undefined
				: extractQualifiedCall(source, stmt.span, moduleSignatures);
			const effectiveStatementCall = statementCall ?? qualifiedStatementCall;
			if (effectiveStatementCall) {
				validateArgumentTypes(effectiveStatementCall, env, moduleSignatures, push);
			}
		});
	}
}

function checkAssignmentTypes(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const moduleSignatures = buildModuleTypeSignatures(mod);
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
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
		});
		checkMemberAssignmentTypes(
			source,
			member,
			env,
			moduleSignatures,
			memberCtx,
			push,
		);
	}
}

/**
 * Rule: a Function/Property Get returns through its hidden return variable.
 * Falling through without assigning that variable is legal VBA, but it silently
 * returns the default value for the declared type, so XLIDE surfaces it as a
 * type-safety warning rather than a VBE compile-error equivalent.
 */
function checkMissingReturnAssignments(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	for (const member of mod.members) {
		if (
			member.kind !== 'Procedure' ||
			(member.procKind !== 'Function' && member.procKind !== 'PropertyGet')
		) {
			continue;
		}
		if (procedureHasReturnAssignment(source, member)) {
			continue;
		}
		const procLabel = member.procKind === 'PropertyGet' ? 'Property Get' : 'Function';
		const typeLabel = member.returnType ? ` As ${member.returnType}` : '';
		push(
			'missingReturnAssignment',
			`${procLabel} '${member.name}' has no return assignment; VBA will return the default value${typeLabel}. Assign to '${member.name}' before exit if a value is intended.`,
			declaredNameSpan(source, member.span, member.name),
		);
	}
}

function procedureHasReturnAssignment(source: string, proc: ProcedureNode): boolean {
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
		}
	});
	return found;
}

function checkMemberAssignmentTypes(
	source: string,
	member: ProcedureNode,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
	memberCtx: MemberCompletionContext,
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
	});
}

function checkSetAssignments(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	memberCtx: MemberCompletionContext,
	push: PushFn,
): void {
	const moduleSignatures = buildModuleTypeSignatures(mod);
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const env = typeEnvironmentFor(symbols, member);
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
		});
	}
}

function checkScalarMemberAccess(
	source: string,
	mod: ModuleNode,
	symbols: ReturnType<typeof buildModuleSymbols>,
	push: PushFn,
): void {
		for (const member of mod.members) {
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
			});
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
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	let i = 0;
	if (
		toks.length >= 2 &&
		(toks[0].kind === 'identifier' || toks[0].kind === 'keyword') &&
		toks[1].rawText === ':'
	) {
		i = 2;
	}
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

function buildModuleTypeSignatures(mod: ModuleNode): Map<string, CallableTypeSignature> {
	const out = new Map<string, CallableTypeSignature>();
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		out.set(member.name.toLowerCase(), procedureTypeSignature(member));
	}
	return out;
}

function procedureTypeSignature(member: ProcedureNode): CallableTypeSignature {
	return {
		name: member.name,
		params: member.params.map((p) => ({
			name: stripHeaderBrackets(p.name),
			type: p.asType,
			optional: p.optional,
			paramArray: p.paramArray,
		})),
		returnType: member.returnType,
	};
}

function callableTypeSignaturesFor(
	mod: ModuleNode,
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
): Map<string, CallableTypeSignature> {
	const out = buildModuleTypeSignatures(mod);
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
			})),
			returnType: candidate.returnType,
		});
	}
	return out;
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

function returnAssignmentTypeFor(proc: ProcedureNode): string | undefined {
	return (proc.procKind === 'Function' || proc.procKind === 'PropertyGet')
		? proc.returnType
		: undefined;
}

function expressionCalls(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
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
				: !callableSignatureFor(name, moduleSignatures)
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
	const out: BoundMemberCall[] = [];
	for (let i = 2; i < toks.length - 1; i++) {
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
		if (isImplicitParenthesizedMemberCallStatement(toks, i, close)) {
			continue;
		}
		const split = inner.length === 0 ? emptyArgSplit() : splitArgSlots(inner, span.start);
		out.push({
			signature: parseRuntimeDisplaySignature(member.name, member.signature),
			call: {
				name: member.name,
				nameSpan: { start: span.start + toks[i].start, end: span.start + toks[i].end },
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
	const toks = statementTokens(source, span);
	if (toks.length === 0 || topLevelOperatorIndex(toks, '=') >= 0) {
		return [];
	}
	const explicitCall = tokenText(toks[0]) === 'call';
	const chainStart = explicitCall ? 1 : 0;
	if (!tokenName(toks[chainStart])) {
		return [];
	}
	const out: BoundMemberCall[] = [];
	for (let i = chainStart + 2; i < toks.length; i++) {
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

function isMemberStatementChainThrough(
	toks: readonly VbaToken[],
	startIdx: number,
	memberIdx: number,
): boolean {
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
	push: PushFn,
): void {
	const sig = callableSignatureForCall(call, moduleSignatures);
	if (!sig || sig.params.length === 0) {
		return;
	}
	validateArgumentTypesForSignature(sig, call, env, moduleSignatures, push);
}

function validateArgumentTypesForSignature(
	sig: CallableTypeSignature,
	call: CallArguments,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
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
		const actual = inferArgumentType(valueSlot, call.sliceStart, env, moduleSignatures);
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
): CallableTypeSignature | undefined {
	if (call.lookupKey) {
		return moduleSignatures.get(call.lookupKey);
	}
	return callableSignatureFor(call.name, moduleSignatures);
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
): CallableTypeSignature | undefined {
	const user = moduleSignatures.get(name.toLowerCase());
	if (user) {
		return user;
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
		paramArray: false,
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
): InferredArgumentType | undefined {
	const toks = slot.filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return inferExpressionType(toks, sliceStart, env, moduleSignatures);
}

function inferExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): InferredArgumentType | undefined {
	const first = toks[0];
	if (!first) {
		return undefined;
	}
	const unwrapped = unwrapOuterParens(toks);
	if (unwrapped !== toks) {
		return inferExpressionType(unwrapped, sliceStart, env, moduleSignatures);
	}
	const concatenation = inferStringConcatenationExpressionType(
		toks,
		sliceStart,
		env,
		moduleSignatures,
	);
	if (concatenation) {
		return concatenation;
	}
	const arithmetic = inferArithmeticExpressionType(toks, sliceStart, env, moduleSignatures);
	if (arithmetic) {
		return arithmetic;
	}
	return inferAtomicExpressionType(toks, sliceStart, env, moduleSignatures);
}

function inferAtomicExpressionType(
	toks: VbaToken[],
	sliceStart: number,
	env: ReadonlyMap<string, string>,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
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
			case 'floatLiteral':
				return { type: 'Double', label: 'numeric literal', span };
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
		return type ? { type, label: `${name} As ${type}`, span } : undefined;
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
		const sig = callableSignatureFor(name, moduleSignatures);
		if (sig?.returnType && matchParenFrom(toks, 1) === toks.length - 1) {
			return { type: sig.returnType, label: `${name}(...) As ${sig.returnType}`, span };
		}
	}
	if (name && toks[1]?.rawText === '.') {
		const member = tokenName(toks[2]);
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
	return undefined;
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
): InferredArgumentType | undefined {
	const parts = splitTopLevelArithmeticOperands(toks);
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const inferred = inferExpressionType(part, sliceStart, env, moduleSignatures);
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
): InferredArgumentType | undefined {
	const parts = splitTopLevelOperands(toks, '&');
	if (parts.length < 2) {
		return undefined;
	}
	for (const part of parts) {
		const inferred = inferExpressionType(part, sliceStart, env, moduleSignatures);
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
		(projectType) => projectType.name.toLowerCase() === lower,
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
	push: PushFn,
): void {
	let hasExplicit = false;
	let hasCode = false;
	for (const member of mod.members) {
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
	knownIdentifiers: ReadonlySet<string> | undefined,
	push: PushFn,
): void {
	if (!hasOptionExplicit(mod) || !knownIdentifiers) {
		return;
	}

	const moduleNames = moduleLevelIdentifierNames(symbols);
	const appType = resolveHostGlobal('Application');
	const appMembers = new Set(
		(appType ? getHostMembers(appType) : []).map((member) => member.name.toLowerCase()),
	);
	const isKnown = (name: string, locals: ReadonlySet<string>): boolean => {
		const lower = name.toLowerCase();
		return (
			locals.has(lower) ||
			moduleNames.has(lower) ||
			knownIdentifiers.has(lower) ||
			appMembers.has(lower) ||
			resolveHostGlobal(name) !== undefined ||
			resolveHostConstant(name) !== undefined ||
			resolveRuntimeConstant(name) !== undefined ||
			resolveRuntimeFunction(name) !== undefined
		);
	};

	for (const member of mod.members) {
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
			for (const ref of undeclaredReadReferences(source, span, (name) => isKnown(name, locals))) {
				report(ref.name, ref.span, 'using it');
			}
		});
	}
}

function forEachUndeclaredReferenceSpan(
	source: string,
	body: BodyNode[],
	visit: (span: Span) => void,
): void {
	for (const node of body) {
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
			forEachUndeclaredReferenceSpan(source, node.body, visit);
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
): Array<{ name: string; span: Span }> {
	const toks = statementTokens(source, span);
	const out: Array<{ name: string; span: Span }> = [];
	const skip = undeclaredReferenceSkipIndexes(source, span, toks, isKnown);
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
): Set<number> {
	const skip = new Set<number>();
	if (toks.length === 0) {
		return skip;
	}
	if (toks[1]?.rawText === ':' || isLineLabelOnlyStatement(source, span, toks)) {
		skip.add(0); // line label declaration
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

function simpleAssignmentLhsIdentifierIndex(toks: readonly VbaToken[]): number {
	let start = 0;
	if (toks[1]?.rawText === ':') {
		start = 2;
	}
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

function hasOptionExplicit(mod: ModuleNode): boolean {
	return mod.members.some(
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
	for (const member of mod.members) {
		if (member.kind === 'VariableGroup') {
			inspect(member);
		} else if (member.kind === 'Procedure') {
			forEachVariableGroup(member.body, inspect);
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
 * project-wide binder, and fixed-length string forms are left alone until their
 * full grammar is verified.
 */
function checkUnexpectedDeclarationTokens(
	source: string,
	mod: ModuleNode,
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

	for (const member of mod.members) {
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
			forEachVariableGroup(member.body, inspectGroup);
		}
	}
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
): void {
	for (const node of body) {
		if (node.kind === 'VariableGroup') {
			visit(node);
		} else if ('body' in node && Array.isArray((node as { body?: unknown }).body)) {
			forEachVariableGroup((node as { body: BodyNode[] }).body, visit);
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

	for (const member of mod.members) {
		if (member.kind === 'VariableGroup' && isPublicModifier(member.modifier)) {
			for (const decl of member.declarations) {
				const span = declaredNameSpan(source, decl.span, decl.name);
				if (member.isConst) {
					report('constants', span);
				} else if (decl.isArray) {
					report('arrays', span);
				} else if (isFixedLengthStringType(decl.asType)) {
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

function checkEventHandlerModuleScope(
	source: string,
	mod: ModuleNode,
	moduleName: string,
	moduleKind: ModuleSymbolKind,
	documentType: EventHandlerDocumentType | undefined,
	push: PushFn,
): void {
	const actualDocumentType = eventHandlerDocumentTypeForContext({
		moduleName,
		moduleKind,
		documentType,
	});
	for (const member of mod.members) {
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

function isFixedLengthStringType(asType: string | undefined): boolean {
	if (!asType) {
		return false;
	}
	const toks = tokenize(asType).filter((t) => t.kind !== 'comment' && t.kind !== 'newline');
	return toks.length >= 3 && tokenText(toks[0]) === 'string' && toks[1]?.rawText === '*';
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

	// Fixed-length string syntax (`As String * 10`) has its own backlog item;
	// do not infer trailing-token errors inside that grammar until verified.
	if (isUnqualifiedStringType(toks, typeStart, i) && toks[i]?.rawText === '*') {
		return undefined;
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

function isUnqualifiedStringType(
	toks: VbaToken[],
	start: number,
	end: number,
): boolean {
	return end === start + 1 && tokenText(toks[start]) === 'string';
}

function checkInvalidAsTypeNames(
	source: string,
	opts: AnalyzeModuleOptions,
	push: PushFn,
): void {
	for (const ref of collectTypeNameReferences(source)) {
		const resolved = resolveTypeName(ref.name, {
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

function declaredNameSpan(source: string, span: Span, name: string): Span {
	const lower = name.toLowerCase();
	for (const tok of statementTokens(source, span)) {
		if (tokenName(tok)?.toLowerCase() === lower) {
			return absoluteSpan(span, tok);
		}
	}
	return span;
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
 * Rule: parameter-order constraints. A required parameter may not follow an
 * `Optional` one, and `ParamArray` must be the final parameter. Both are read
 * straight off the parsed parameter flags, so they are deterministic.
 */
function checkParameterOrder(mod: ModuleNode, push: PushFn): void {
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		const params = member.params;
		let optionalSeen = false;
		for (let i = 0; i < params.length; i++) {
			const p = params[i];
			if (p.paramArray) {
				if (i !== params.length - 1) {
					push(
						'paramArrayNotLast',
						`ParamArray '${p.name}' must be the last parameter.`,
						p.span,
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
					p.span,
				);
			}
		}
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
	projectProcedures: ReadonlyMap<string, readonly VbaProcedureSignature[]> | undefined,
	push: PushFn,
): void {
	const moduleSignatures = callableTypeSignaturesFor(mod, projectProcedures);
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			const at = unparenthesizedCallArg(source, stmt.span);
			if (at) {
				push(
					'callRequiresParens',
					'A Call statement requires parentheses around its argument list.',
					at,
				);
			}
			const bare = implicitParenthesizedBareCallableCall(source, stmt.span, moduleSignatures);
			if (bare) {
				push(
					'callStatementForbidsParens',
					'Standalone zero-argument procedure calls cannot use empty parentheses unless they are prefixed with Call or used in an expression.',
					bare.span,
				);
			}
			const implicit = implicitParenthesizedMemberCall(source, stmt.span);
			if (implicit) {
				push(
					'callStatementForbidsParens',
					'Standalone zero-argument member calls cannot use empty parentheses unless they are prefixed with Call or used in an expression.',
					implicit.span,
				);
			}
		});
	}
}

function checkInvalidExpressionSyntax(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			const hit = invalidOperatorSequence(source, stmt.span);
			if (hit) {
				push(
					'invalidExpressionSyntax',
					`Invalid operator sequence '${hit.text}'; this will fail to compile as a syntax error.`,
					hit.span,
				);
			}
		});
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

function isNonUnaryBinaryOperator(tok: VbaToken | undefined): boolean {
	if (!tok || tok.kind !== 'operator') {
		return false;
	}
	return NON_UNARY_BINARY_OPERATORS.has(tokenText(tok));
}

/**
 * Rule: when a Function is used inside an expression, its argument list must be
 * parenthesized (`x = Foo(1, 2)`). The parenless form (`Foo 1, 2`) is only a
 * call-statement form and becomes a VBE syntax error after `=`.
 */
function checkExpressionCallParens(
	source: string,
	mod: ModuleNode,
	push: PushFn,
): void {
	const moduleFunctions = new Set<string>();
	for (const member of mod.members) {
		if (
			member.kind === 'Procedure' &&
			(member.procKind === 'Function' || member.procKind === 'PropertyGet')
		) {
			moduleFunctions.add(member.name.toLowerCase());
		}
	}

	for (const member of mod.members) {
		if (member.kind !== 'Procedure') {
			continue;
		}
		forEachStatement(member.body, (stmt) => {
			const hit = parenlessExpressionCall(source, stmt.span, moduleFunctions);
			if (hit) {
				push(
					'expressionCallRequiresParens',
					`Function call arguments in an expression must be enclosed in parentheses: use '${hit.name}(...)'.`,
					hit.span,
				);
			}
		});
	}
}

function parenlessExpressionCall(
	source: string,
	span: Span,
	moduleFunctions: ReadonlySet<string>,
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
		if (!name || !isExpressionCallable(name, moduleFunctions)) {
			continue;
		}
		if (i > eq + 1 && toks[i - 1].rawText === '.') {
			continue; // member calls need receiver typing before we can be precise
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

function isExpressionCallable(
	name: string,
	moduleFunctions: ReadonlySet<string>,
): boolean {
	const lower = name.toLowerCase();
	if (moduleFunctions.has(lower)) {
		return true;
	}
	return resolveRuntimeFunction(name)?.kind === 'function';
}

function isParenlessArgumentStart(tok: VbaToken | undefined): boolean {
	if (!tok) {
		return false;
	}
	switch (tok.kind) {
		case 'identifier':
		case 'keyword':
		case 'bracketedIdentifier':
		case 'integerLiteral':
		case 'floatLiteral':
		case 'stringLiteral':
		case 'dateLiteral':
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

/** Returns the span of the first stray argument token in a `Call` statement. */
function unparenthesizedCallArg(source: string, span: Span): Span | undefined {
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
	if (toks.length === 0 || toks[0].rawText.toLowerCase() !== 'call') {
		return undefined;
	}
	let i = 1;
	const callee = toks[i];
	if (
		!callee ||
		(callee.kind !== 'identifier' && callee.kind !== 'bracketedIdentifier')
	) {
		return undefined; // malformed Call header - not this rule's concern
	}
	i++;
	for (;;) {
		const t = toks[i];
		if (!t) {
			break;
		}
		if (t.rawText === '.') {
			const name = toks[i + 1];
			if (
				name &&
				(name.kind === 'identifier' ||
					name.kind === 'bracketedIdentifier' ||
					name.kind === 'keyword')
			) {
				i += 2;
				continue;
			}
			break;
		}
		if (t.rawText === '(') {
			const close = matchParenFrom(toks, i);
			if (close < 0) {
				return undefined; // unbalanced - reported elsewhere
			}
			i = close + 1;
			continue;
		}
		break;
	}
	const stray = toks[i];
	if (stray) {
		return { start: span.start + stray.start, end: span.start + stray.end };
	}
	return undefined;
}

function implicitParenthesizedBareCallableCall(
	source: string,
	span: Span,
	moduleSignatures: ReadonlyMap<string, CallableTypeSignature>,
): { name: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	if (
		toks.length !== 3 ||
		toks[0]?.rawText.toLowerCase() === 'call' ||
		toks[1]?.rawText !== '(' ||
		toks[2]?.rawText !== ')' ||
		topLevelOperatorIndex(toks, '=') >= 0
	) {
		return undefined;
	}
	const name = tokenName(toks[0]);
	if (!name) {
		return undefined;
	}
	const signature = callableSignatureFor(name, moduleSignatures);
	if (!signature || !callableAcceptsZeroArguments(signature)) {
		return undefined;
	}
	return {
		name,
		span: { start: span.start + toks[0].start, end: span.start + toks[2].end },
	};
}

function callableAcceptsZeroArguments(sig: CallableTypeSignature): boolean {
	return sig.params.every((param) => param.optional || param.paramArray);
}

function implicitParenthesizedMemberCall(
	source: string,
	span: Span,
): { name: string; span: Span } | undefined {
	const toks = statementTokens(source, span);
	for (let i = 2; i < toks.length - 1; i++) {
		const name = tokenName(toks[i]);
		if (!name || toks[i - 1]?.rawText !== '.' || toks[i + 1]?.rawText !== '(') {
			continue;
		}
		const close = matchParenFrom(toks, i + 1);
		if (close < 0 || !isImplicitParenthesizedMemberCallStatement(toks, i, close)) {
			continue;
		}
		return {
			name,
			span: { start: span.start + toks[i].start, end: span.start + toks[close].end },
		};
	}
	return undefined;
}

function isImplicitParenthesizedMemberCallStatement(
	toks: readonly VbaToken[],
	memberIdx: number,
	closeIdx: number,
): boolean {
	if (toks[0]?.rawText.toLowerCase() === 'call') {
		return false;
	}
	if (
		closeIdx !== toks.length - 1 ||
		closeIdx !== memberIdx + 2 ||
		topLevelOperatorIndex(toks, '=') >= 0
	) {
		return false;
	}
	return isCompleteMemberChainThroughCall(toks, memberIdx, closeIdx);
}

function isCompleteMemberChainThroughCall(
	toks: readonly VbaToken[],
	memberIdx: number,
	closeIdx: number,
): boolean {
	const first = toks[0];
	if (!first || !tokenName(first)) {
		return false;
	}
	let i = 1;
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
		const nameTok = toks[nameIdx];
		if (!nameTok || !tokenName(nameTok)) {
			return false;
		}
		if (nameIdx === memberIdx) {
			return (
				toks[nameIdx + 1]?.rawText === '(' &&
				matchParenFrom(toks, nameIdx + 1) === closeIdx
			);
		}
		i = nameIdx + 1;
	}
	return false;
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
	push: PushFn,
): void {
	for (const member of mod.members) {
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
		});
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
	const toks = tokenize(source.slice(span.start, span.end)).filter(
		(t) => t.kind !== 'comment' && t.kind !== 'newline',
	);
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
	push: PushFn,
): void {
	const root: StatementContext = {
		forDepth: 0,
		doDepth: 0,
		withDepth: 0,
		selectDepth: 0,
	};

	for (const member of mod.members) {
		if (member.kind === 'Statement') {
			checkContextStatement(source, member, root, push);
		} else if (member.kind === 'Procedure') {
			checkContextBody(source, member.body, root, push);
		}
	}
}

function checkContextBody(
	source: string,
	body: BodyNode[],
	ctx: StatementContext,
	push: PushFn,
): void {
	for (const node of body) {
		switch (node.kind) {
			case 'Statement':
				checkContextStatement(source, node, ctx, push);
				break;
			case 'ForBlock':
				checkContextBody(source, node.body, { ...ctx, forDepth: ctx.forDepth + 1 }, push);
				break;
			case 'DoBlock':
				checkContextBody(source, node.body, { ...ctx, doDepth: ctx.doDepth + 1 }, push);
				break;
			case 'WithBlock':
				checkContextBody(source, node.body, { ...ctx, withDepth: ctx.withDepth + 1 }, push);
				break;
			case 'SelectBlock':
				checkContextBody(source, node.body, { ...ctx, selectDepth: ctx.selectDepth + 1 }, push);
				break;
			case 'IfBlock':
			case 'WhileBlock':
				checkContextBody(source, node.body, ctx, push);
				break;
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
	const toks = statementTokens(source, stmt.span);
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

function tokenText(token: VbaToken | undefined): string {
	return (token?.canonicalText ?? token?.rawText ?? '').toLowerCase();
}

function absoluteSpan(base: Span, token: VbaToken): Span {
	return { start: base.start + token.start, end: base.start + token.end };
}

function exitPhraseSpan(base: Span, first: VbaToken, target: VbaToken): Span {
	return { start: base.start + first.start, end: base.start + target.end };
}

/**
 * Rule: `Option` statements must precede every declaration and procedure (only
 * `Attribute` lines may come before them in an exported module). Once a real
 * declaration has appeared, any later `Option` is misplaced.
 */
function checkOptionPlacement(mod: ModuleNode, push: PushFn): void {
	let declarationSeen = false;
	for (const member of mod.members) {
		if (member.kind === 'Attribute') {
			continue;
		}
		if (member.kind === 'Option') {
			if (declarationSeen) {
				push(
					'optionAfterDeclaration',
					'Option statements must appear before any declaration or procedure.',
					member.span,
				);
			}
			continue;
		}
		declarationSeen = true;
	}
}
