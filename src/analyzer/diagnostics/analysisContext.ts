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
import { statementTokens as computeStatementTokens } from '../lexer/tokenHelpers';
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
import type { ConditionalCompilationEnvironment } from '../conditional/conditionalCompilation';
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
	/**
	 * Pre-parsed AST for the analyzed source. When omitted, the analyzer parses
	 * the source itself.
	 */
	parsedModule?: ModuleNode;
}

/** The diagnostics sink every rule reports through. */
export type PushFn = (
	rule: DiagnosticRuleName,
	message: string,
	span: Span,
	data?: VbaDiagnosticData,
) => void;

export function isObjectModuleKind(moduleKind: ModuleSymbolKind | undefined): boolean {
	return moduleKind === 'class' || moduleKind === 'document' || moduleKind === 'userform';
}

let APPLICATION_MEMBER_NAMES: ReadonlySet<string> | undefined;

export function applicationMemberNames(): ReadonlySet<string> {
	if (!APPLICATION_MEMBER_NAMES) {
		const appType = resolveHostGlobal('Application');
		APPLICATION_MEMBER_NAMES = new Set(
			(appType ? getHostMembers(appType) : []).map((member) => member.name.toLowerCase()),
		);
	}
	return APPLICATION_MEMBER_NAMES;
}

// Statement-token cache (audit #5): independent rules re-tokenized the same
// statement 25-40 times per analysis pass. Tokens are cached per source
// string (value identity, LRU of 2 like tokenizeCached) and per statement
// span, so one pass lexes each statement once. Callers must not mutate the
// returned arrays or their tokens; the diagnostics engine treats token
// streams as read-only throughout.
const STATEMENT_TOKEN_CACHE_MAX = 2;
const statementTokenCache: { src: string; bySpan: Map<number, VbaToken[]> }[] = [];

/** Significant tokens of a statement span, excluding comments and newlines (memoized per pass). */
export function statementTokens(source: string, span: Span): VbaToken[] {
	let entry: { src: string; bySpan: Map<number, VbaToken[]> } | undefined;
	for (let i = 0; i < statementTokenCache.length; i += 1) {
		if (statementTokenCache[i].src === source) {
			entry = statementTokenCache[i];
			if (i > 0) {
				statementTokenCache.splice(i, 1);
				statementTokenCache.unshift(entry);
			}
			break;
		}
	}
	if (!entry) {
		entry = { src: source, bySpan: new Map() };
		statementTokenCache.unshift(entry);
		if (statementTokenCache.length > STATEMENT_TOKEN_CACHE_MAX) {
			statementTokenCache.pop();
		}
	}
	// Statement spans never overlap, so the start offset identifies the span.
	const key = span.start * 0x100000000 + span.end;
	let toks = entry.bySpan.get(key);
	if (!toks) {
		toks = computeStatementTokens(source, span);
		entry.bySpan.set(key, toks);
	}
	return toks;
}

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
