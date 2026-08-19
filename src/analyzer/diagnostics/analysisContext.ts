// Per-pass shared state for the diagnostics engine.
//
// This module owns the contracts every diagnostics rule shares (the diagnostic
// shape, the analyzer options, and the `push` callback type) plus the per-pass
// memoized state that used to be recomputed rule-by-rule inside
// `analyzeModule.ts`. Caches here are value-keyed on per-pass identities (the
// `buildModuleSymbols` result, the analyzed source string), following the
// existing `tokenizeCached` precedent, so independent rules transparently
// share one computation without threading a context object through every
// helper signature.

import type { VbaToken } from '../lexer/tokenKinds';
import { getHostMembers, resolveHostGlobal } from '../host/hostModel';
import type { ModuleNode, Span } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type {
	ModuleSymbolKind,
	VbaProjectClassMembers,
	VbaProcedureSignature,
	VbaSymbol,
} from '../symbols/symbolModel';
import { isProcedureKind } from '../symbols/symbolModel';
import type { ProcedureNode } from '../parser/nodes';
import type { HostObjectModel } from '../host/excelObjectModel';
import type {
	ConditionalActivityTracker,
	ConditionalCompilationEnvironment,
} from '../conditional/conditionalCompilation';
import type { MemberCompletionContext } from '../completion/memberAccess';
import type { EventHandlerDocumentType } from '../completion/eventHandlers';
import type { ProjectTypeName } from '../completion/typeCompletion';
import type {
	DiagnosticRuleName,
	DiagnosticSeverity,
	DiagnosticSeverityOverride,
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
	/**
	 * Provenance for incremental re-analysis: 'run' = a module-wide rule pass,
	 * 'walk' = the shared per-procedure statement/expression walk. Plain data so
	 * results stay structured-cloneable across worker boundaries.
	 */
	origin?: 'run' | 'walk';
	/** For 'walk' diagnostics: span.start of the procedure being walked. */
	walkMemberStart?: number;
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

/** Inputs for `analyzeModule`. */
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
	 * Members the module has that no line of its own text declares - a UserForm's
	 * controls, declared by the designer. Referring to one is correct VBA, so
	 * without this every reference reads as an undeclared variable.
	 *
	 * Carries the type as well as the name so a member lookup can resolve it,
	 * rather than only silencing the finding.
	 */
	implicitMembers?: readonly { name: string; type: string }[];
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
	 * Which Office host the module belongs to, as a token (`excel`, `word`,
	 * `powerpoint`, `access`, ...). Resolved through the host registry when
	 * `hostModel` is not supplied directly: absent means Excel, and a named
	 * host with no model yet means no host knowledge at all rather than
	 * Excel's (issue #24).
	 */
	host?: string;
	/**
	 * Conditional-compilation constants for deterministic branch filtering. Branches
	 * that remain unknown are still analyzed; only proven-inactive code is skipped.
	 */
	conditionalCompilation?: ConditionalCompilationEnvironment;
	/**
	 * Pre-parsed AST for the analyzed source. When omitted, the analyzer parses
	 * the source itself.
	 */
	parsedModule?: ModuleNode;
	/**
	 * Incremental re-analysis support: when set, the shared statement/expression
	 * walks skip the bodies of procedures for which this returns false, and any
	 * walk-phase diagnostics for those procedures are dropped (the incremental
	 * layer splices the cached ones back in). Module-wide rule passes are never
	 * filtered. Rule factories are still invoked for every procedure so factory-
	 * level bookkeeping stays identical to a full pass.
	 */
	walkProcedureFilter?: (member: ProcedureNode) => boolean;
}

/** The diagnostics sink every rule reports through. */
export type PushFn = (
	rule: DiagnosticRuleName,
	message: string,
	span: Span,
	data?: VbaDiagnosticData,
) => void;

/**
 * Everything one diagnostics pass computes once and every rule shares: the
 * analyzed source, the resolved module identity, the caller's options, the
 * parsed AST, the module symbol table, the conditional-compilation activity
 * tracker, and the member-resolution context primed with the per-pass AST and
 * token stream (audit #0/#1).
 */
export interface RulePassContext {
	source: string;
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	opts: AnalyzeModuleOptions;
	mod: ModuleNode;
	symbols: ReturnType<typeof buildModuleSymbols>;
	activity: ConditionalActivityTracker | undefined;
	memberCtx: MemberCompletionContext;
}

export function isObjectModuleKind(moduleKind: ModuleSymbolKind | undefined): boolean {
	return moduleKind === 'class' || moduleKind === 'document' || moduleKind === 'userform';
}

const APPLICATION_MEMBER_NAMES = new WeakMap<HostObjectModel, ReadonlySet<string>>();
let DEFAULT_APPLICATION_MEMBER_NAMES: ReadonlySet<string> | undefined;

/**
 * The host's Application members, injected into the bare global scope the way
 * Office hosts inject them (Calculate, Volatile, ... under Excel). Keyed per
 * model, so a Word caller gets Word's set and a host with no model injects
 * nothing at all.
 */
export function applicationMemberNames(model?: HostObjectModel): ReadonlySet<string> {
	if (model === undefined) {
		DEFAULT_APPLICATION_MEMBER_NAMES ??= computeApplicationMemberNames(undefined);
		return DEFAULT_APPLICATION_MEMBER_NAMES;
	}
	let names = APPLICATION_MEMBER_NAMES.get(model);
	if (!names) {
		names = computeApplicationMemberNames(model);
		APPLICATION_MEMBER_NAMES.set(model, names);
	}
	return names;
}

function computeApplicationMemberNames(model: HostObjectModel | undefined): ReadonlySet<string> {
	const appType = resolveHostGlobal('Application', model);
	return new Set(
		(appType ? getHostMembers(appType, model) : []).map((member) => member.name.toLowerCase()),
	);
}

// The statement-token cache (audit #5) now lives in lexer/tokenHelpers as
// statementTokensCached, so every analyzer surface shares one
// implementation; this re-export keeps the diagnostics engine's historical
// import path working.
export { statementTokensCached as statementTokens } from '../lexer/tokenHelpers';

// Procedure symbols are looked up per procedure per rule, so index them once
// per buildModuleSymbols result instead of scanning the children every time.
const PROCEDURE_SYMBOLS_BY_START = new WeakMap<
	ReturnType<typeof buildModuleSymbols>,
	Map<number, VbaSymbol>
>();

export function procedureSymbolFor(
	symbols: ReturnType<typeof buildModuleSymbols>,
	proc: ProcedureNode,
): VbaSymbol | undefined {
	let byStart = PROCEDURE_SYMBOLS_BY_START.get(symbols);
	if (!byStart) {
		byStart = new Map<number, VbaSymbol>();
		for (const sym of symbols.root.children ?? []) {
			if (isProcedureKind(sym.kind) && !byStart.has(sym.fullSpan.start)) {
				byStart.set(sym.fullSpan.start, sym);
			}
		}
		PROCEDURE_SYMBOLS_BY_START.set(symbols, byStart);
	}
	return byStart.get(proc.span.start);
}
