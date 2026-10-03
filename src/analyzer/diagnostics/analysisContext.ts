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

import type { OpenedFileNumbers } from './openedFileNumbers';
import { getHostMembers, resolveHostGlobal, resolveHostGlobalMember } from '../host/hostModel';
import type { ModuleNode, Span } from '../parser/nodes';
import type { buildModuleSymbols } from '../symbols/buildModuleSymbols';
import type {
	ModuleSymbolKind,
	VbaProjectClassMembers,
	VbaProcedureSignature,
	VbaSymbol,
} from '../symbols/symbolModel';
import { isProcedureKind } from '../symbols/symbolModel';
import type { FormControlInfo } from '../symbols/projectIndex';
import type { SheetChanges, WorkbookSheetInfo } from '../symbols/sheetChanges';
import type { ProcedureNode } from '../parser/nodes';
import { getExcelObjectModel, type HostObjectModel } from '../host/excelObjectModel';
import type {
	ConditionalActivityTracker,
	ConditionalCompilationEnvironment,
} from '../conditional/conditionalCompilation';
import { msFormsControlMembers, type MemberCompletionContext } from '../completion/memberAccess';
import { VBA_USERFORM_TYPE } from '../host/userFormExtenderMembers';
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

/** A `Dim` the editor can insert for a name Option Explicit says is missing. */
export interface VbaDeclareVariableData {
	variableName: string;
	/** The type the declaration will name, e.g. `Long` or `Variant`. */
	declaredType: string;
	edit: {
		span: Span;
		newText: string;
	};
}

/** The edit that removes a declaration nothing uses. */
export interface VbaRemoveDeclarationData {
	variableName: string;
	edit: {
		span: Span;
		newText: string;
	};
}

/** The edit that removes the statements nothing can reach. */
export interface VbaRemoveUnreachableCodeData {
	edit: {
		span: Span;
		newText: string;
	};
}

/** One way to bring a doc comment in line with its declaration. */
export interface VbaDocCommentFix {
	title: string;
	isPreferred?: boolean;
	edits: {
		span: Span;
		newText: string;
	}[];
}

export interface VbaDiagnosticData {
	missingRequiredArgumentPlaceholder?: VbaMissingRequiredArgumentPlaceholderData;
	createProcedureStub?: VbaCreateProcedureStubData;
	declareVariable?: VbaDeclareVariableData;
	removeDeclaration?: VbaRemoveDeclarationData;
	removeUnreachableCode?: VbaRemoveUnreachableCodeData;
	docCommentFixes?: VbaDocCommentFix[];
	/** The Office library a missing-reference diagnostic would have the user add. */
	addLibraryReference?: VbaAddLibraryReferenceData;
}

/** Which library to add, for the quick fix on a missing reference. */
export interface VbaAddLibraryReferenceData {
	/** The host token the engine adds a reference for: excel, word, ... */
	library: string;
}

/** Per-rule severity overrides keyed by stable diagnostic code; `'off'` disables an allowed rule. */
export type DiagnosticSeverityOverrides = Partial<Record<string, DiagnosticSeverityOverride>>;

/** Inputs for `analyzeModule`. */
/**
 * Where analyzeModule recovered from a failure of its own, and so what it did
 * not check: `options`, an option it could not use and left out; `analysis`,
 * the whole pass (nothing was checked); `rule`, one rule; `statement-walk` and
 * `expression-walk`, one rule's visitor for the rest of the module, or with no
 * rule named, the walk itself for the rest of the module.
 */
export interface AnalysisFailure {
	stage: 'options' | 'analysis' | 'rule' | 'statement-walk' | 'expression-walk';
	rule?: string;
}

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
	implicitMembers?: readonly FormControlInfo[];
	/**
	 * The host type of the class the module's DESIGNER makes it, when the host
	 * can say: `VB.Form`, `VB.MDIForm`, `VB.UserControl`, `VB.PropertyPage`.
	 * Its members belong to the module the same way a control does - the text
	 * never declares `Arrange` or `PropertyChanged`, and calling one bare or
	 * through `Me` is correct code. Without it the module's own procedures are
	 * the whole of `Me`, so every inherited member reads as missing.
	 */
	designerClass?: string;
	/**
	 * Exported Sub/Function/Declare signatures across the project, grouped by
	 * lowercased procedure name. When omitted, type and arity validation remain
	 * single-module only.
	 */
	projectProcedures?: ReadonlyMap<string, readonly VbaProcedureSignature[]>;
	/**
	 * Called for each failure analyzeModule recovers from rather than throws
	 * (issue #178). The diagnostics it still returns are those of the rules
	 * that ran; this says what was not checked, so a host can tell a clean
	 * module from one the analyser could not fully check. A callback that
	 * throws is ignored.
	 */
	onInternalError?: (error: unknown, where: AnalysisFailure) => void;
	/** Source-declared project object members and UDT fields visible to this module. */
	projectClassMembers?: readonly VbaProjectClassMembers[];
	/** Source-declared project type names visible to this module. */
	projectTypes?: readonly ProjectTypeName[];
	/**
	 * Lowercased names of every module some module in the project declares with
	 * `Implements`. A module named here is an interface, so its own members are
	 * declarations for an implementer to fill in rather than unfinished code.
	 */
	implementedInterfaces?: ReadonlySet<string>;
	/**
	 * Source-backed module-level symbols visible as bare identifiers from this
	 * module. Used by call-target diagnostics to distinguish "not found" from
	 * "found, but not callable" across modules.
	 */
	projectVisibleSymbols?: readonly VbaSymbol[];
	/** Lowercased visible declaration names known not to be type names. */
	knownNonTypeNames?: ReadonlySet<string>;
	/**
	 * Lowercased Private Type and Enum names of other modules, bare and as
	 * `module.name`, which this module cannot use as a type (issue #490).
	 */
	hiddenTypeNames?: ReadonlySet<string>;
	/**
	 * Raw integer constant expressions exported from other visible project modules.
	 * Used as a conservative base for deterministic runtime-value diagnostics.
	 */
	projectIntegerConstants?: ReadonlyMap<string, string | undefined>;
	/**
	 * Lowercased identifier-shaped words inside every string literal in the
	 * project (from ProjectIndex.stringLiteralWords). A Private procedure
	 * named in one may be reached through `Application.Run`, `OnTime` or a
	 * control's `OnAction`, so the unused-procedure rule treats the name as
	 * used. When omitted, the module's own string literals are searched.
	 */
	projectStringLiteralWords?: ReadonlySet<string>;
	/**
	 * Lowercased names `Application.Run` can reach in the project (from
	 * ProjectIndex.runnableProcedureNames): every Sub and Function of a
	 * standard or document module, Private ones included, bare and as
	 * `module.name` (issue #243). When omitted, Application.Run is not judged.
	 */
	projectRunnableProcedures?: ReadonlySet<string>;
	/**
	 * Lowercased names some module of the project may write (from
	 * ProjectIndex.writtenNames). A Public variable none writes keeps its
	 * initial value, which the runtime rules use (issue #241). When omitted,
	 * a Public variable is never taken as unchanged.
	 */
	projectWrittenNames?: ReadonlySet<string>;
	/**
	 * How many modules of the project mention each lowercased name (from
	 * ProjectIndex.nameMentions). A form's list or MultiPage that only its
	 * own module names is judged from the designer's contents (issue #315).
	 * When omitted, it is not.
	 */
	projectNameMentions?: ReadonlyMap<string, number>;
	/**
	 * The sheets of the workbook the project lives in, as saved, in tab order
	 * (issue #229). `ThisWorkbook.Sheets("name")` and `(index)` are checked
	 * against them, and only when `projectSheetChanges` says no code in the
	 * project could have made the sheet. Absent for anything but an Excel file.
	 */
	workbookSheets?: readonly WorkbookSheetInfo[];
	/** What the project's code may do to the workbook's sheets (ProjectIndex.sheetChanges). */
	projectSheetChanges?: SheetChanges;
	/** The file numbers the project's Open statements name (ProjectIndex.openedFileNumbers). */
	projectOpenedFileNumbers?: OpenedFileNumbers;
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
	 * The other applications whose type libraries the project references, as
	 * host tokens. A Word document that references Excel compiles `Dim xl As
	 * Excel.Application`, so the rules run against both models, the project's
	 * own host winning any name the two share.
	 */
	referencedHosts?: readonly string[];
	/**
	 * The names of every library the project references, as its dir stream
	 * records them (`Scripting`, `MSForms`). Undefined where the project's
	 * references are not known, which keeps the rules that read it silent.
	 */
	referencedLibraries?: readonly string[];
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

/**
 * Where the host has a Global interface - Excel, Word, PowerPoint - that is
 * what VBA really calls bare, and resolveHostGlobalMember answers for it,
 * hidden members and all. Application's documented members stand in for it
 * here because they match it closely; its hidden ones do not, so they stay
 * out. `Save` is a hidden method of Excel's `_Application` and no member of
 * `_Global` (measured on EXCEL.EXE), so a bare `Save` is "Sub or Function not
 * defined", and letting the model's hidden members in would have accepted it.
 *
 * Access has no Global: its type library makes Application itself the object
 * VBA binds bare, so there every member of it, hidden or not, is in scope.
 */
function computeApplicationMemberNames(model: HostObjectModel | undefined): ReadonlySet<string> {
	const appType = resolveHostGlobal('Application', model);
	const globalAnswers = (model ?? getExcelObjectModel()).globalType !== undefined;
	// Only what Global has too: a bare `ScreenUpdating`, `Name` or `Caption`
	// is "Variable not defined" in a standard module, Application's though
	// they are (issue #318, measured in Excel 16.0; 250 of Application's 338
	// members are not Global's).
	return new Set(
		(appType ? getHostMembers(appType, model) : [])
			.filter((member) => !(globalAnswers && (member.hidden || resolveHostGlobalMember(member.name, model) === undefined)))
			.map((member) => member.name.toLowerCase()),
	);
}

const NO_NAMES: ReadonlySet<string> = new Set();
const DESIGNER_CLASS_MEMBER_NAMES = new WeakMap<HostObjectModel, Map<string, ReadonlySet<string>>>();

/**
 * The members of the class a module's designer makes it, lowercased. Inside
 * such a module they are in scope unqualified - `PropertyChanged` in a
 * UserControl, `Show` in a form - exactly as the module's own procedures are,
 * because the module IS one of these. Empty when the host cannot say which
 * class it is, or the model does not carry that type.
 */
/**
 * The members of a module's own object, which its code may name bare (issue
 * #228, measured in Excel 16.0): `UsedRange` and `Shapes` in a sheet's
 * module, `FullName` and `Saved` in ThisWorkbook, `Controls`, `Tag` and
 * `Repaint` in a UserForm's. A sheet whose document type is unknown takes a
 * Worksheet's and a Chart's, since it may be either. A module with a
 * designer class takes that class's members through designerClassMemberNames.
 */
export function ownObjectMemberNames(opts: AnalyzeModuleOptions): ReadonlySet<string> {
	if (opts.designerClass) {
		return NO_NAMES;
	}
	if (opts.moduleKind === 'userform') {
		return new Set((msFormsControlMembers(VBA_USERFORM_TYPE) ?? []).map((member) => member.name.toLowerCase()));
	}
	if (opts.moduleKind !== 'document') {
		return NO_NAMES;
	}
	const model = opts.hostModel ?? getExcelObjectModel();
	const host = (model.hostName ?? 'Excel').toLowerCase();
	const name = opts.moduleName?.toLowerCase();
	const classes = opts.documentType === 'worksheet' ? ['Excel.Worksheet']
		: opts.documentType === 'chart' ? ['Excel.Chart']
			: opts.documentType === 'workbook' ? ['Excel.Workbook']
				: opts.documentType === 'document' ? ['Word.Document']
					: host === 'excel' ? (name === 'thisworkbook' ? ['Excel.Workbook'] : ['Excel.Worksheet', 'Excel.Chart'])
						: host === 'word' && name === 'thisdocument' ? ['Word.Document'] : [];
	const names = new Set<string>();
	for (const type of classes) {
		for (const member of getHostMembers(type, model)) {
			names.add(member.name.toLowerCase());
		}
	}
	return names;
}

export function designerClassMemberNames(
	designerClass: string | undefined,
	model: HostObjectModel | undefined,
): ReadonlySet<string> {
	if (!designerClass || !model) {
		return NO_NAMES;
	}
	let byType = DESIGNER_CLASS_MEMBER_NAMES.get(model);
	if (!byType) {
		byType = new Map();
		DESIGNER_CLASS_MEMBER_NAMES.set(model, byType);
	}
	const key = designerClass.toLowerCase();
	let names = byType.get(key);
	if (!names) {
		names = new Set(getHostMembers(designerClass, model).map((member) => member.name.toLowerCase()));
		byType.set(key, names);
	}
	return names;
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
