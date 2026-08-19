// Project-wide symbol graph (Phase 4: Project-Wide Symbol Graph).
//
// Aggregates per-module symbol views into a workbook-project index and provides
// document symbols, workspace symbols, and conservative go-to-definition name
// resolution. No `vscode` dependency: the VS Code providers consume this index
// and translate spans to ranges.
//
// Name resolution order verified against MS-VBAL.pdf, v20250520, section 5.3
// (procedure scope) and 4.2 / 5.2.3.1 (module vs project visibility); see
// docs/spec/MS-VBAL.verification-map.md.

import { buildModuleSymbols, type BuildModuleSymbolsOptions } from './buildModuleSymbols';
import {
	enumMemberRawExpression,
	parseVbaIntegerLiteral,
	resolveRawIntegerConstants,
} from '../constants/integerConstantExpression';
import {
	isBareCallableKind,
	isProcedureKind,
	qualifiedProcedureKey,
	type ModuleSymbolKind,
	type ModuleSymbols,
	type VbaProjectTypeKind,
	type VbaProjectTypeName,
	type VbaProjectClassMember,
	type VbaProjectClassMemberDefinition,
	type VbaProjectClassMembers,
	type VbaSymbol,
	type VbaSymbolAttribute,
	type VbaProcedureSignature,
	formatProcedureParamLabel,
	procedureSignatureLabel,
	procedureParamsFromSymbol,
	procedureSignatureFromSymbol,
} from './symbolModel';
import {
	resolveBareIdentifierBinding,
	type BareIdentifierContext,
	type BareIdentifierResolution,
} from './nameResolution';
import type { Span } from '../parser/nodes';
import { hasAuthoritativeDesignerHeader, parseUserFormControls } from '../../vbaUserFormControls';

/** Source text + workbook role for one module fed into the index. */
export interface ModuleInput {
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	source: string;
	/** Optional per-module conditional-compilation environment. */
	conditionalCompilation?: BuildModuleSymbolsOptions['conditionalCompilation'];
	/**
	 * Members the module has that its own text never declares - a UserForm's
	 * designer-declared controls, supplied by a host that can read the
	 * designer. When absent, a form's controls come from parsing its own
	 * `.frm` header, which only standalone VB6-style exports carry.
	 */
	implicitMembers?: readonly { name: string; type: string }[];
}

/** Project-wide symbol graph options shared by every indexed module. */
export interface ProjectIndexOptions {
	conditionalCompilation?: BuildModuleSymbolsOptions['conditionalCompilation'];
}

/** How widely an identifier reference binds across the project. */
export type ReferenceScopeKind = 'local' | 'module' | 'project';

/** A procedure span (within a named module) that shadows a name with a local. */
export interface ShadowedSpan {
	moduleName: string;
	span: Span;
}

/**
 * The binding scope of an identifier, used to restrict reference/rename search
 * to exactly the modules and spans where the name binds to the same
 * declaration. Computed from the symbol graph without a full expression binder:
 *   - `local`   - a parameter/local/const of the enclosing procedure; search is
 *                 limited to {@link procedureSpan} in the single search module.
 *   - `module`  - a module-private declaration (or an unresolved name); search
 *                 is limited to the owning module.
 *   - `project` - an exported (Public/Global, default-Public procedure, or
 *                 exported enum member) declaration; search spans every module
 *                 except those that re-declare the name privately at module
 *                 level, and excludes procedures whose locals shadow the name
 *                 ({@link shadowedSpans}).
 */
export interface ReferenceScope {
	kind: ReferenceScopeKind;
	/** Resolved declaration(s); empty when the name does not resolve. */
	definitions: VbaSymbol[];
	/** Modules to scan for textual occurrences (original casing). */
	searchModules: string[];
	/** Enclosing procedure span when `kind === 'local'`. */
	procedureSpan?: Span;
	/** Procedure spans (per module) whose locals shadow the name. */
	shadowedSpans: ShadowedSpan[];
}

interface ModuleLevelBinding {
	symbol: VbaSymbol;
	exported: boolean;
}

/**
 * A top-level symbol is "exported" for cross-module lookup when it is explicitly
 * Public/Global, or when it is an unmodified procedure in a standard module
 * (procedures default to Public there - MS-VBAL 5.3.1.1). Dim/Private/Friend/
 * Static and unmodified module variables/consts stay module-private.
 */
function isExported(symbol: VbaSymbol, moduleKind?: ModuleSymbolKind): boolean {
	if (symbol.visibility === 'Public' || symbol.visibility === 'Global') {
		return true;
	}
	if (symbol.visibility) {
		return false;
	}
	return moduleKind === 'standard' && isProcedureKind(symbol.kind);
}

function addProcedureSignature(
	signatures: Map<string, VbaProcedureSignature[]>,
	key: string,
	sig: VbaProcedureSignature,
): void {
	const existing = signatures.get(key);
	if (existing) {
		existing.push(sig);
	} else {
		signatures.set(key, [sig]);
	}
}

function moduleKindAsTypeName(kind: ModuleSymbolKind): VbaProjectTypeKind | undefined {
	switch (kind) {
		case 'class':
			return 'class';
		case 'document':
			return 'document';
		case 'userform':
			return 'userform';
		default:
			return undefined;
	}
}

function projectTypeKind(symbol: VbaSymbol): VbaProjectTypeKind | undefined {
	switch (symbol.kind) {
		case 'enum':
			return 'enum';
		case 'type':
			return 'userType';
		default:
			return undefined;
	}
}

function isTypeExported(symbol: VbaSymbol): boolean {
	return symbol.visibility !== 'Private';
}

function isEnumMemberExported(
	enumSymbol: VbaSymbol,
	moduleKind?: ModuleSymbolKind,
): boolean {
	return moduleKind === 'standard' && isTypeExported(enumSymbol);
}

function moduleRawIntegerConstantExpressions(mod: ModuleSymbols): Map<string, string | undefined> {
	const out = new Map<string, string | undefined>();
	const seen = new Set<string>();
	const add = (name: string, raw: string | undefined): void => {
		const key = name.toLowerCase();
		if (seen.has(key)) {
			out.set(key, undefined);
			out.set(`${mod.moduleName.toLowerCase()}.${key}`, undefined);
			return;
		}
		seen.add(key);
		out.set(key, raw);
		out.set(`${mod.moduleName.toLowerCase()}.${key}`, raw);
	};
	for (const symbol of mod.root.children ?? []) {
		if (symbol.kind === 'constant') {
			add(symbol.name, symbol.defaultRaw);
			continue;
		}
		if (symbol.kind === 'enum') {
			let previousName: string | undefined;
			for (const member of symbol.children ?? []) {
				add(member.name, enumMemberRawExpression(member.defaultRaw, previousName));
				previousName = member.name;
			}
		}
	}
	return out;
}

function isVisibleProjectObjectMember(symbol: VbaSymbol): boolean {
	if (isProcedureKind(symbol.kind)) {
		return symbol.visibility !== 'Private';
	}
	if (symbol.kind === 'event') {
		return symbol.visibility !== 'Private';
	}
	if (symbol.kind === 'moduleVariable') {
		return symbol.visibility === 'Public' || symbol.visibility === 'Global';
	}
	return false;
}

function isVisibleStandardModuleMember(
	symbol: VbaSymbol,
	mod: ModuleSymbols,
	sameModule: boolean,
): boolean {
	if (!projectObjectMemberKind(symbol)) {
		return false;
	}
	if (sameModule) {
		return true;
	}
	if (symbol.kind === 'enum') {
		return isTypeExported(symbol);
	}
	if (symbol.kind === 'enumMember') {
		const container = enumContainerForMember(mod, symbol);
		return container ? isEnumMemberExported(container, mod.moduleKind) : false;
	}
	return isExported(symbol, mod.moduleKind);
}

function projectObjectMemberKind(symbol: VbaSymbol): VbaProjectClassMember['kind'] | undefined {
	switch (symbol.kind) {
		case 'sub':
		case 'function':
		case 'declare':
			return 'method';
		case 'event':
			return 'event';
		case 'propertyGet':
		case 'propertyLet':
		case 'propertySet':
		case 'moduleVariable':
		case 'constant':
		case 'enum':
		case 'enumMember':
			return 'property';
		default:
			return undefined;
	}
}

function projectObjectMemberWritable(symbol: VbaSymbol): boolean | undefined {
	switch (symbol.kind) {
		case 'propertyLet':
		case 'propertySet':
		case 'moduleVariable':
			return true;
		case 'propertyGet':
		case 'constant':
		case 'enum':
		case 'enumMember':
			return false;
		default:
			return undefined;
	}
}

function projectObjectMemberWriteType(symbol: VbaSymbol): string | undefined {
	switch (symbol.kind) {
		case 'propertyLet':
		case 'propertySet':
			return lastParameter(symbol)?.asType;
		case 'moduleVariable':
			return symbol.asType;
		default:
			return undefined;
	}
}

function enumContainerForMember(
	mod: ModuleSymbols,
	member: VbaSymbol,
): VbaSymbol | undefined {
	const lower = member.containerName?.toLowerCase();
	if (!lower) {
		return undefined;
	}
	return (mod.root.children ?? []).find(
		(symbol) => symbol.kind === 'enum' && symbol.name.toLowerCase() === lower,
	);
}

function projectObjectMemberReturnType(symbol: VbaSymbol): string | undefined {
	if (symbol.kind === 'enumMember') {
		return symbol.containerName;
	}
	return symbol.asType;
}

function projectObjectMemberDefinition(symbol: VbaSymbol): VbaProjectClassMemberDefinition {
	return {
		moduleName: symbol.moduleName,
		nameSpan: symbol.nameSpan,
		fullSpan: symbol.fullSpan,
	};
}

/**
 * OLE Automation DISPID for a type's default member (the `[default]` member,
 * DISPID_VALUE). VBE exports it as `Attribute X.VB_UserMemId = 0`; parse the
 * value numerically rather than by string match so non-canonical forms like
 * `&H0` resolve deterministically, while `-4` (DISPID_NEWENUM, the `_NewEnum`
 * enumerator) correctly stays non-default.
 */
const DISPID_VALUE = 0;

function isDefaultMemberAttribute(attr: VbaSymbolAttribute): boolean {
	return (
		attr.name.toLowerCase() === 'vb_usermemid' &&
		parseVbaIntegerLiteral(attr.valueRaw) === DISPID_VALUE
	);
}

function isDefaultProjectObjectMember(symbol: VbaSymbol): boolean {
	return (symbol.attributes ?? []).some(isDefaultMemberAttribute);
}

function mergeMemberAttributes(
	existing: readonly VbaSymbolAttribute[] | undefined,
	incoming: readonly VbaSymbolAttribute[] | undefined,
): VbaSymbolAttribute[] | undefined {
	if (!incoming || incoming.length === 0) {
		return existing ? [...existing] : undefined;
	}
	const out = [...(existing ?? [])];
	const seen = new Set(out.map((attr) => `${attr.fullSpan.start}:${attr.fullSpan.end}`));
	for (const attr of incoming) {
		const key = `${attr.fullSpan.start}:${attr.fullSpan.end}`;
		if (seen.has(key)) {
			continue;
		}
		seen.add(key);
		out.push(attr);
	}
	return out;
}

function moduleImplements(source: string): string[] {
	const out: string[] = [];
	const seen = new Set<string>();
	for (const line of source.split(/\r?\n/)) {
		// \p{L}: a Cyrillic or Thai interface name was invisible here, so the
		// class hierarchy came back empty and Interface_Member never resolved.
		const match = /^\s*Implements\s+([\p{L}_][\p{L}\p{M}\p{N}_]*(?:\.[\p{L}_][\p{L}\p{M}\p{N}_]*)?)/iu.exec(
			line,
		);
		if (!match) {
			continue;
		}
		const name = match[1];
		const lower = name.toLowerCase();
		if (!seen.has(lower)) {
			seen.add(lower);
			out.push(name);
		}
	}
	return out;
}

function lastParameter(symbol: VbaSymbol): VbaSymbol | undefined {
	const params = (symbol.children ?? []).filter((child) => child.kind === 'parameter');
	return params[params.length - 1];
}

function projectObjectMemberSignature(symbol: VbaSymbol): string | undefined {
	const procedure = procedureSignatureFromSymbol(symbol);
	if (procedure) {
		return procedureSignatureLabel(procedure);
	}
	if (symbol.kind === 'event') {
		const params = procedureParamsFromSymbol(symbol)
			.map((param) => formatProcedureParamLabel(param))
			.join(', ');
		return `${symbol.name}(${params})`;
	}
	if (symbol.kind !== 'propertyGet') {
		return undefined;
	}
	const params = procedureParamsFromSymbol(symbol)
		.map((param) => formatProcedureParamLabel(param))
		.join(', ');
	const returns = symbol.asType ? ` As ${symbol.asType}` : '';
	return `${symbol.name}(${params})${returns}`;
}

function projectMemberCandidateSymbols(
	mod: ModuleSymbols,
	includeEnumMembers: boolean,
): VbaSymbol[] {
	const out: VbaSymbol[] = [];
	for (const symbol of mod.root.children ?? []) {
		out.push(symbol);
		if (includeEnumMembers && symbol.kind === 'enum') {
			out.push(...(symbol.children ?? []));
		}
	}
	return out;
}

function userTypeFieldSignature(symbol: VbaSymbol): string {
	const fixedLength = symbol.fixedLength ? ` * ${symbol.fixedLength}` : '';
	const as = symbol.asType ? ` As ${symbol.asType}${fixedLength}` : '';
	return `${symbol.name}${as}`;
}

/** A project-wide symbol index built from a set of module sources. */
export class ProjectIndex {
	private readonly modules = new Map<string, ModuleSymbols>();
	private readonly moduleSources = new Map<string, string>();
	/** Host-supplied designer members (a form's controls), per module name. */
	private readonly moduleImplicitMembersByName = new Map<string, readonly { name: string; type: string }[]>();
	/** Lazily resolved per-module integer constants, dropped on module change. */
	private readonly moduleResolvedConstants = new Map<string, Map<string, number | undefined>>();
	/** Lazily scanned per-module Implements lists, dropped on module change. */
	private readonly moduleImplementsLists = new Map<string, string[]>();
	/** Whole-project query memo for the current index revision. */
	private readonly queryCache = new Map<string, unknown>();

	constructor(private readonly options: ProjectIndexOptions = {}) {}

	/** Adds or replaces a module in the index. */
	setModule(input: ModuleInput): void {
		const symbols = buildModuleSymbols(
			input.moduleName,
			input.moduleKind,
			input.source,
			{
				conditionalCompilation:
					input.conditionalCompilation ?? this.options.conditionalCompilation,
			},
		);
		const key = input.moduleName.toLowerCase();
		this.modules.set(key, symbols);
		this.moduleSources.set(key, input.source);
		if (input.implicitMembers !== undefined) {
			this.moduleImplicitMembersByName.set(key, input.implicitMembers);
		} else {
			this.moduleImplicitMembersByName.delete(key);
		}
		this.invalidate(key);
	}

	/** Removes a module from the index. */
	removeModule(moduleName: string): void {
		const key = moduleName.toLowerCase();
		this.modules.delete(key);
		this.moduleSources.delete(key);
		this.moduleImplicitMembersByName.delete(key);
		this.invalidate(key);
	}

	/** Drops module-derived artifacts and every whole-project query memo. */
	private invalidate(key: string): void {
		this.moduleResolvedConstants.delete(key);
		this.moduleImplementsLists.delete(key);
		this.queryCache.clear();
	}

	/** Memoizes a whole-project query until the indexed modules change. */
	private cached<T>(key: string, compute: () => T): T {
		if (this.queryCache.has(key)) {
			return this.queryCache.get(key) as T;
		}
		const value = compute();
		this.queryCache.set(key, value);
		return value;
	}

	/** Resolved integer constant values of one module, computed at most once. */
	private moduleIntegerConstants(mod: ModuleSymbols): ReadonlyMap<string, number | undefined> {
		const key = mod.moduleName.toLowerCase();
		let resolved = this.moduleResolvedConstants.get(key);
		if (!resolved) {
			resolved = resolveRawIntegerConstants(moduleRawIntegerConstantExpressions(mod));
			this.moduleResolvedConstants.set(key, resolved);
		}
		return resolved;
	}

	/**
	 * A module's source as the index holds it, by module name. A UserForm's
	 * controls are declared in its `.frm` header rather than its code, so
	 * reading them needs the text the index already has.
	 */
	moduleSource(moduleName: string): string | undefined {
		return this.moduleSources.get(moduleName.toLowerCase());
	}

	/**
	 * A form's designer-declared controls: what the host supplied with the
	 * module, or what the module's own `.frm` header carries. Excel stores a
	 * workbook form's control tree in a binary designer blob, so for a
	 * workbook-backed form with no host this answers nothing.
	 */
	moduleImplicitMembers(moduleName: string): readonly { name: string; type: string }[] {
		const key = moduleName.toLowerCase();
		const supplied = this.moduleImplicitMembersByName.get(key);
		if (supplied !== undefined) {
			return supplied;
		}
		// No kind gate: the header itself is the evidence. Only a form's source
		// opens with VERSION 5.00 and designer Begin blocks, so every other
		// module parses to nothing - and a form whose kind arrived mislabeled
		// still answers its controls.
		return this.cached(`implicitMembers:${key}`, () =>
			parseUserFormControls(this.moduleSources.get(key) ?? ''));
	}

	/**
	 * True when the module's control list is AUTHORITATIVE: a designer-reading
	 * host supplied it with the module (an empty array included), or the
	 * source itself carries a `.frm` designer header. Absence claims about a
	 * form's members are only sound behind this - a workbook form whose
	 * binary designer nobody has read has an unknown control list, not an
	 * empty one (issue #26).
	 */
	moduleImplicitMembersKnown(moduleName: string): boolean {
		const key = moduleName.toLowerCase();
		if (this.moduleImplicitMembersByName.get(key) !== undefined) {
			return true;
		}
		return hasAuthoritativeDesignerHeader(this.moduleSources.get(key) ?? '');
	}

	/**
	 * Interfaces a module declares with `Implements`, by module name. Renaming
	 * an interface needs this to know which classes carry its member prefix.
	 */
	moduleImplementsList(moduleName: string): string[] {
		const mod = this.modules.get(moduleName.toLowerCase());
		return mod ? this.moduleImplementsFor(mod) : [];
	}

	/** Implements declarations of one module, scanned at most once. */
	private moduleImplementsFor(mod: ModuleSymbols): string[] {
		const key = mod.moduleName.toLowerCase();
		let list = this.moduleImplementsLists.get(key);
		if (!list) {
			list = moduleImplements(this.moduleSources.get(key) ?? '');
			this.moduleImplementsLists.set(key, list);
		}
		return list;
	}

	/** All module names currently indexed (original casing). */
	moduleNames(): string[] {
		return [...this.modules.values()].map((m) => m.moduleName);
	}

	/**
	 * Lowercased procedure names callable as bare identifiers from `moduleName`.
	 * Same-module procedures are always visible to their own module. Cross-module
	 * bare calls are limited to exported procedures in standard modules; class,
	 * document, and UserForm members require object/module-qualified binding that
	 * the unknown-call rule deliberately does not guess.
	 */
	visibleProcedureNames(moduleName: string): Set<string> {
		const currentLower = moduleName.toLowerCase();
		return new Set(this.cached(`procedureNames:${currentLower}`, () => {
			const names = new Set<string>();
			for (const mod of this.modules.values()) {
				const sameModule = mod.moduleName.toLowerCase() === currentLower;
				for (const symbol of mod.root.children ?? []) {
					if (!isBareCallableKind(symbol.kind)) {
						continue;
					}
					if (
						sameModule ||
						(mod.moduleKind === 'standard' && isExported(symbol, mod.moduleKind))
					) {
						names.add(symbol.name.toLowerCase());
					}
				}
			}
			return names;
		}));
	}

	/**
	 * Visible bare-call Sub/Function/Declare signatures from `moduleName`.
	 * Same-module callables are visible to their own module. Other modules
	 * contribute only exported standard-module callables, matching the
	 * `visibleProcedureNames` rule used by diagnostics.
	 */
	visibleProcedureSignatures(moduleName: string): VbaProcedureSignature[] {
		const currentLower = moduleName.toLowerCase();
		return this.cached(`procedureSignatures:${currentLower}`, () => {
			const out: VbaProcedureSignature[] = [];
			for (const mod of this.modules.values()) {
				const sameModule = mod.moduleName.toLowerCase() === currentLower;
				for (const symbol of mod.root.children ?? []) {
					if (!isBareCallableKind(symbol.kind)) {
						continue;
					}
					if (
						!sameModule &&
						(mod.moduleKind !== 'standard' || !isExported(symbol, mod.moduleKind))
					) {
						continue;
					}
					const signature = procedureSignatureFromSymbol(symbol);
					if (signature) {
						out.push(signature);
					}
				}
			}
			return out;
		}).slice();
	}

	/**
	 * Lowercased bare identifiers visible from `moduleName`. Used by diagnostics
	 * that must know whether an identifier is declared under `Option Explicit`.
	 * Same-module declarations are visible regardless of visibility; other
	 * standard modules contribute exported declarations and exported enum members.
	 * Document/UserForm code names are also available as global object variables.
	 */
	visibleIdentifierNames(moduleName: string): Set<string> {
		const currentLower = moduleName.toLowerCase();
		return new Set(this.cached(`identifierNames:${currentLower}`, () => {
			const names = new Set<string>();
			for (const mod of this.modules.values()) {
				const sameModule = mod.moduleName.toLowerCase() === currentLower;
				if (mod.moduleKind === 'document' || mod.moduleKind === 'userform') {
					names.add(mod.moduleName.toLowerCase());
				}
				for (const symbol of this.visibleModuleLevelIdentifierSymbols(mod, sameModule)) {
					names.add(symbol.name.toLowerCase());
				}
			}
			return names;
		}));
	}

	/**
	 * Source-backed module-level symbols visible as bare identifiers from
	 * `moduleName`. Document/UserForm code names are intentionally not included
	 * here because they are object-module globals rather than source
	 * declarations; callers that need them should use the workbook module list.
	 */
	visibleIdentifierSymbols(moduleName: string): VbaSymbol[] {
		const currentLower = moduleName.toLowerCase();
		return this.cached(`identifierSymbols:${currentLower}`, () => {
			const out: VbaSymbol[] = [];
			for (const mod of this.modules.values()) {
				const sameModule = mod.moduleName.toLowerCase() === currentLower;
				out.push(...this.visibleModuleLevelIdentifierSymbols(mod, sameModule));
			}
			return out;
		}).slice();
	}

	/**
	 * Raw integer constant expressions exported from other standard modules and
	 * visible as bare identifiers from `moduleName`. Duplicate visible names are
	 * kept with an unknown value so diagnostics never guess a binding.
	 */
	visibleExternalIntegerConstantExpressions(moduleName: string): Map<string, string | undefined> {
		const currentLower = moduleName.toLowerCase();
		return new Map(this.cached(`integerConstants:${currentLower}`, () => {
			const out = new Map<string, string | undefined>();
			const seen = new Set<string>();
			const add = (name: string, raw: string | undefined): void => {
				const key = name.toLowerCase();
				if (seen.has(key)) {
					out.set(key, undefined);
					return;
				}
				seen.add(key);
				out.set(key, raw);
			};
			const addQualified = (mod: ModuleSymbols, name: string, raw: string | undefined): void => {
				out.set(`${mod.moduleName.toLowerCase()}.${name.toLowerCase()}`, raw);
			};
			const resolvedRaw = (
				resolved: ReadonlyMap<string, number | undefined>,
				key: string,
				fallback: string | undefined,
			): string | undefined => {
				const value = resolved.get(key.toLowerCase());
				return value === undefined ? fallback : String(value);
			};
			for (const mod of this.modules.values()) {
				if (mod.moduleName.toLowerCase() === currentLower || mod.moduleKind !== 'standard') {
					continue;
				}
				const moduleResolved = this.moduleIntegerConstants(mod);
				for (const symbol of mod.root.children ?? []) {
					if (symbol.kind === 'constant' && isExported(symbol, mod.moduleKind)) {
						const raw = resolvedRaw(moduleResolved, symbol.name, symbol.defaultRaw);
						add(symbol.name, raw);
						addQualified(mod, symbol.name, raw);
						continue;
					}
					if (symbol.kind === 'enum' && isEnumMemberExported(symbol, mod.moduleKind)) {
						let previousName: string | undefined;
						for (const member of symbol.children ?? []) {
							const fallback = enumMemberRawExpression(member.defaultRaw, previousName);
							const raw = resolvedRaw(moduleResolved, member.name, fallback);
							add(member.name, raw);
							addQualified(mod, member.name, raw);
							previousName = member.name;
						}
					}
				}
			}
			return out;
		}));
	}

	/**
	 * Lowercased visible declaration names that are known not to be type names.
	 * Used by type-position diagnostics after the type resolver has failed, so
	 * project/primitive/host type names still take precedence over value names.
	 */
	visibleNonTypeNames(moduleName: string): Set<string> {
		const currentLower = moduleName.toLowerCase();
		return new Set(this.cached(`nonTypeNames:${currentLower}`, () => {
			const names = new Set<string>();
			for (const mod of this.modules.values()) {
				const sameModule = mod.moduleName.toLowerCase() === currentLower;
				for (const symbol of this.visibleModuleLevelIdentifierSymbols(mod, sameModule)) {
					if (projectTypeKind(symbol)) {
						continue;
					}
					names.add(symbol.name.toLowerCase());
				}
			}
			return names;
		}));
	}

	/**
	 * Exported standard-module Sub/Function/Declare signatures grouped by
	 * lowercased procedure name, with additional `module.procedure` qualified
	 * keys. Bare duplicate exported names intentionally remain grouped together so analyzer
	 * callers can skip ambiguous unqualified calls; module-qualified calls can
	 * still resolve deterministically through their qualified key.
	 *
	 * Properties are deliberately excluded from this first callable-signature
	 * surface because their invocation syntax and Let/Set/Get pairing needs the
	 * object/member binder.
	 */
	procedureSignatures(): Map<string, VbaProcedureSignature[]> {
		return new Map(this.cached('procedureSignaturesByKey', () => {
			const signatures = new Map<string, VbaProcedureSignature[]>();
			for (const mod of this.modules.values()) {
				for (const symbol of mod.root.children ?? []) {
					if (
						!isBareCallableKind(symbol.kind) ||
						mod.moduleKind !== 'standard' ||
						!isExported(symbol, mod.moduleKind)
					) {
						continue;
					}
					const sig = procedureSignatureFromSymbol(symbol);
					if (sig) {
						addProcedureSignature(signatures, symbol.name.toLowerCase(), sig);
						addProcedureSignature(
							signatures,
							qualifiedProcedureKey(symbol.moduleName, symbol.name),
							sig,
						);
					}
				}
			}
			return signatures;
		}));
	}

	/**
	 * Project-defined type names visible from `moduleName`, excluding intrinsic
	 * VBA types and host object-model types. Current-module `Type`/`Enum`
	 * declarations are visible regardless of `Private`; other modules expose only
	 * non-Private `Type`/`Enum` declarations. Class, document, and UserForm module
	 * names are represented as type names because they are object modules.
	 *
	 * Duplicates are preserved deliberately so the shared type resolver and
	 * diagnostics can report ambiguity instead of silently picking whichever
	 * module happened to be read first.
	 */
	visibleTypeNames(moduleName: string): VbaProjectTypeName[] {
		const currentLower = moduleName.toLowerCase();
		return this.cached(`typeNames:${currentLower}`, () => {
			const out: VbaProjectTypeName[] = [];
			for (const mod of this.modules.values()) {
				const sameModule = mod.moduleName.toLowerCase() === currentLower;
				const moduleTypeKind = moduleKindAsTypeName(mod.moduleKind);
				if (moduleTypeKind) {
					out.push({
						name: mod.moduleName,
						kind: moduleTypeKind,
						moduleName: mod.moduleName,
						nameSpan: mod.root.nameSpan,
						fullSpan: mod.root.fullSpan,
						doc: mod.root.doc,
					});
				}

				for (const symbol of mod.root.children ?? []) {
					const kind = projectTypeKind(symbol);
					if (!kind) {
						continue;
					}
					if (!sameModule && !isTypeExported(symbol)) {
						continue;
					}
					out.push({
						name: symbol.name,
						kind,
						moduleName: mod.moduleName,
						nameSpan: symbol.nameSpan,
						fullSpan: symbol.fullSpan,
						visibility: symbol.visibility,
						doc: symbol.doc,
					});
				}
			}
			return out;
		}).slice();
	}

	/**
	 * Source-backed type-name definitions visible from `moduleName`. Object-module
	 * class/document/UserForm names resolve to the top of their module; Type/Enum
	 * declarations resolve to the declaration identifier.
	 */
	resolveTypeDefinitions(moduleName: string, name: string): VbaProjectTypeName[] {
		const lower = name.toLowerCase();
		return this.visibleTypeNames(moduleName).filter(
			(typeName) => typeName.name.toLowerCase() === lower,
		);
	}

	/**
	 * Public/default-public members of workbook-defined object modules. This is
	 * the source-backed surface used by member completion for variables declared
	 * `As Person` where `Person` is a class/UserForm/document module. Private
	 * members are deliberately hidden. Public fields are represented as properties;
	 * Property Get/Let/Set declarations collapse to one property item. Public
	 * constants are intentionally excluded because VBE rejects them in object
	 * modules.
	 */
	projectClassMembers(): VbaProjectClassMembers[] {
		return this.cached('projectClassMembers', () => {
			const out: VbaProjectClassMembers[] = [];
			for (const mod of this.modules.values()) {
				const kind = moduleKindAsTypeName(mod.moduleKind);
				if (kind !== 'class' && kind !== 'document' && kind !== 'userform') {
					continue;
				}
				const members = this.visibleObjectMembers(mod);
				if (kind === 'userform') {
					// A form's controls are members of the form, declared by the
					// designer rather than by code, so a qualified reference from
					// another module (`EntryForm.NameBox`) must find them on the
					// form's type - not only inside its own code-behind (#22).
					const own = new Set(members.map((member) => member.name.toLowerCase()));
					for (const control of this.moduleImplicitMembers(mod.moduleName)) {
						if (own.has(control.name.toLowerCase())) {
							continue;
						}
						members.push({
							name: control.name,
							kind: 'property',
							returns: control.type,
							moduleName: mod.moduleName,
						});
					}
				}
				out.push({
					name: mod.moduleName,
					kind,
					moduleName: mod.moduleName,
					implements: this.moduleImplementsFor(mod),
					doc: mod.root.doc,
					// Classes are source-exhaustive. A form is exhaustive when
					// its control list is authoritative (issue #26): with the
					// controls and code-behind here and the UserForm base
					// merged at resolution, the surface proves absence the
					// same way the VBE's compiler does. Document modules stay
					// non-exhaustive: their host base carries more than any
					// list here.
					exhaustive: kind === 'class'
						|| (kind === 'userform' && this.moduleImplicitMembersKnown(mod.moduleName)),
					members,
				});
			}
			return out;
		}).slice();
	}

	/**
	 * Source-backed member surfaces for standard modules. These are not type
	 * names, but they are valid module-qualified receivers such as
	 * `XlideAssert.AreEqual`.
	 */
	projectStandardModuleMembers(moduleName: string): VbaProjectClassMembers[] {
		const currentLower = moduleName.toLowerCase();
		const out: VbaProjectClassMembers[] = [];
		for (const mod of this.modules.values()) {
			if (mod.moduleKind !== 'standard') {
				continue;
			}
			const sameModule = mod.moduleName.toLowerCase() === currentLower;
			out.push({
				name: mod.moduleName,
				kind: 'standardModule',
				moduleName: mod.moduleName,
				doc: mod.root.doc,
				exhaustive: true,
				members: this.visibleStandardModuleMembers(mod, sameModule),
			});
		}
		return out;
	}

	/**
	 * Source-backed member surfaces visible from `moduleName`: workbook object
	 * modules, standard module-qualified members, plus visible `Type ... End Type`
	 * declarations. UDT fields are exhaustive, writable property-like members.
	 */
	projectMemberSurfaces(moduleName: string): VbaProjectClassMembers[] {
		const currentLower = moduleName.toLowerCase();
		return this.cached(`memberSurfaces:${currentLower}`, () => [
			...this.projectClassMembers(),
			...this.projectStandardModuleMembers(moduleName),
			...this.projectUserTypeMembers(moduleName),
			...this.projectEnumMembers(moduleName),
		]).slice();
	}

	/** The {@link ModuleSymbols} for a module, or undefined. */
	getModule(moduleName: string): ModuleSymbols | undefined {
		return this.modules.get(moduleName.toLowerCase());
	}

	/**
	 * Hierarchical document symbols for a module: the module root whose children
	 * are its members (procedures with param/local children, types with fields,
	 * enums with members).
	 */
	documentSymbols(moduleName: string): VbaSymbol | undefined {
		return this.modules.get(moduleName.toLowerCase())?.root;
	}

	/**
	 * Flat workspace symbols across every module, optionally filtered by a
	 * case-insensitive substring query. Module roots are excluded.
	 */
	workspaceSymbols(query?: string): VbaSymbol[] {
		const needle = query?.trim().toLowerCase();
		const out: VbaSymbol[] = [];
		for (const mod of this.modules.values()) {
			for (const symbol of mod.all) {
				if (!needle || symbol.name.toLowerCase().includes(needle)) {
					out.push(symbol);
				}
			}
		}
		return out;
	}

	/**
	 * Conservative go-to-definition resolution for an identifier `name` used at
	 * `offset` inside `moduleName`. Resolution order (MS-VBAL 5.3 scope rules):
	 *   1. Parameters and locals of the enclosing procedure.
	 *   2. Module-level declarations in the same module.
	 *   3. Exported (Public/Global, or default-Public procedure) declarations in
	 *      other modules.
	 * Returns every matching declaration (e.g. a Property Get and Let share a
	 * name), or an empty array when nothing resolves.
	 */
	resolveDefinition(
		moduleName: string,
		name: string,
		offset: number,
	): VbaSymbol[] {
		return this.resolveBareIdentifier(moduleName, name, offset, 'expression').definitions.slice();
	}

	/**
	 * Context-aware bare identifier resolution with the shared source precedence
	 * ladder: procedure locals/parameters, same-module declarations, then visible
	 * exported project declarations. Ambiguous project/module tiers are reported
	 * explicitly so diagnostic callers can stay silent instead of guessing.
	 */
	resolveBareIdentifier(
		moduleName: string,
		name: string,
		offset: number,
		context: BareIdentifierContext,
	): BareIdentifierResolution {
		const home = this.modules.get(moduleName.toLowerCase());
		if (!home) {
			return {
				name,
				lowerName: name.toLowerCase(),
				context,
				scope: 'unresolved',
				definitions: [],
				reason: `Module '${moduleName}' is not indexed.`,
			};
		}
		const resolution = resolveBareIdentifierBinding({
			currentModule: home,
			name,
			context,
			enclosingProcedure: this.enclosingProcedure(home, offset),
			offset,
			projectVisibleSymbols: this.visibleIdentifierSymbols(moduleName),
		});
		// Document/UserForm code names (Sheet1, UserForm1) are object-module
		// globals that visibleIdentifierNames reports as declared but
		// visibleIdentifierSymbols omits, so a bare reference would otherwise
		// resolve to nothing. Additively fall back to the module root symbol when
		// the source ladder found no match, keeping go-to-definition consistent
		// with the diagnostics that treat the name as bound.
		if (resolution.scope === 'unresolved') {
			const objectModule = this.documentOrUserFormModuleNamed(name);
			if (objectModule) {
				return {
					...resolution,
					scope: 'project',
					tier: 'project',
					definitions: [objectModule.root],
					reason: `object-module global ${context} binding for '${name}' in ${objectModule.moduleName}.`,
				};
			}
		}
		return resolution;
	}

	/**
	 * Finds the Document/UserForm module whose code name matches `name`
	 * (case-insensitive). These code names act as global object variables.
	 */
	private documentOrUserFormModuleNamed(name: string): ModuleSymbols | undefined {
		const lower = name.toLowerCase();
		const mod = this.modules.get(lower);
		if (mod && (mod.moduleKind === 'document' || mod.moduleKind === 'userform')) {
			return mod;
		}
		return undefined;
	}

	/**
	 * Go-to-definition for a qualified reference `Qualifier.name` (e.g.
	 * `Module1.DoWork`). Resolves to exported module-level declarations of `name`
	 * in the named module only.
	 */
	resolveQualifiedDefinition(
		qualifier: string,
		name: string,
	): VbaSymbol[] {
		const mod = this.modules.get(qualifier.toLowerCase());
		if (!mod) {
			return [];
		}
		return this.exportedModuleLevelMatches(mod, name.toLowerCase());
	}

	/**
	 * Determines the binding scope of the identifier `name` referenced at
	 * `offset` inside `moduleName`, so reference/rename callers can restrict
	 * their textual search to where the name binds to the same declaration.
	 * Resolution order mirrors {@link resolveDefinition} (MS-VBAL 5.3 scope).
	 */
	referenceScope(
		moduleName: string,
		name: string,
		offset: number,
	): ReferenceScope {
		const lower = name.toLowerCase();
		const home = this.modules.get(moduleName.toLowerCase());
		const resolved = home
			? this.resolveBareIdentifier(moduleName, name, offset, 'expression')
			: undefined;

		if (home && resolved) {
			if (resolved.scope === 'local') {
				// Reuse the enclosing procedure resolveBareIdentifier already
				// computed instead of re-running the O(n) scan on this hot path.
				const enclosing = resolved.enclosingProcedure;
				return {
					kind: 'local',
					definitions: resolved.definitions.slice(),
					searchModules: [home.moduleName],
					procedureSpan: enclosing?.fullSpan,
					shadowedSpans: [],
				};
			}
			if (
				resolved.scope === 'module' ||
				(resolved.scope === 'ambiguous' && resolved.tier === 'module')
			) {
				const exportedHomeHits = this.exportedModuleLevelMatches(home, lower);
				if (exportedHomeHits.length > 0) {
					return this.projectScope(lower, exportedHomeHits);
				}
				return {
					kind: 'module',
					definitions: resolved.definitions.slice(),
					searchModules: [home.moduleName],
					shadowedSpans: this.localShadowSpans(home, lower),
				};
			}
			if (resolved.scope === 'project') {
				return this.projectScope(lower, resolved.definitions.slice());
			}
			if (resolved.scope === 'ambiguous' && resolved.tier === 'project') {
				return this.projectScope(lower, resolved.definitions.slice());
			}
		}

		// Unresolved (host member, undeclared, etc.): stay inside the home module.
		return {
			kind: 'module',
			definitions: [],
			searchModules: home ? [home.moduleName] : [moduleName],
			shadowedSpans: home ? this.localShadowSpans(home, lower) : [],
		};
	}


	/** Procedure names declared more than once in a module (duplicate diagnostic). */
	duplicateProcedures(moduleName: string): VbaSymbol[] {
		const mod = this.modules.get(moduleName.toLowerCase());
		if (!mod) {
			return [];
		}
		const seen = new Map<string, VbaSymbol[]>();
		for (const symbol of mod.root.children ?? []) {
			if (isProcedureKind(symbol.kind)) {
				const key = symbol.name.toLowerCase();
				const list = seen.get(key) ?? [];
				list.push(symbol);
				seen.set(key, list);
			}
		}
		const dupes: VbaSymbol[] = [];
		for (const list of seen.values()) {
			if (list.length > 1) {
				dupes.push(...list);
			}
		}
		return dupes;
	}

	/** Finds the procedure symbol whose full span contains `offset`. */
	private enclosingProcedure(
		mod: ModuleSymbols,
		offset: number,
	): VbaSymbol | undefined {
		return (mod.root.children ?? []).find(
			(c) =>
				isProcedureKind(c.kind) &&
				offset >= c.fullSpan.start &&
				offset <= c.fullSpan.end,
		);
	}

	/**
	 * Builds a project-wide reference scope for an exported `lower` name:
	 * searches every module that does not re-declare the name privately at module
	 * level, and collects the procedure spans whose locals shadow it.
	 */
	private projectScope(lower: string, definitions: VbaSymbol[]): ReferenceScope {
		const searchModules: string[] = [];
		const shadowedSpans: ShadowedSpan[] = [];
		for (const mod of this.modules.values()) {
			const moduleHits = this.moduleLevelBindings(mod, lower);
			const privatelyShadowed =
				moduleHits.length > 0 &&
				!moduleHits.some((binding) => binding.exported);
			if (privatelyShadowed) {
				continue;
			}
			searchModules.push(mod.moduleName);
			shadowedSpans.push(...this.localShadowSpans(mod, lower));
		}
		return { kind: 'project', definitions, searchModules, shadowedSpans };
	}

	/** Procedure spans in `mod` whose params/locals/consts shadow `lower`. */
	private localShadowSpans(mod: ModuleSymbols, lower: string): ShadowedSpan[] {
		const spans: ShadowedSpan[] = [];
		for (const symbol of mod.root.children ?? []) {
			if (!isProcedureKind(symbol.kind)) {
				continue;
			}
			const shadows = (symbol.children ?? []).some(
				(c) =>
					(c.kind === 'parameter' ||
						c.kind === 'localVariable' ||
						c.kind === 'constant') &&
					c.name.toLowerCase() === lower,
			);
			if (shadows) {
				spans.push({ moduleName: mod.moduleName, span: symbol.fullSpan });
			}
		}
		return spans;
	}

	private isBareIdentifierVisible(
		symbol: VbaSymbol,
		mod: ModuleSymbols,
		sameModule: boolean,
	): boolean {
		if (sameModule) {
			return true;
		}
		if (mod.moduleKind !== 'standard') {
			return false;
		}
		if (symbol.kind === 'enum' || symbol.kind === 'type') {
			return isTypeExported(symbol);
		}
		return isExported(symbol, mod.moduleKind);
	}

	private visibleModuleLevelIdentifierSymbols(
		mod: ModuleSymbols,
		sameModule: boolean,
	): VbaSymbol[] {
		const out: VbaSymbol[] = [];
		for (const symbol of mod.root.children ?? []) {
			if (!this.isBareIdentifierVisible(symbol, mod, sameModule)) {
				continue;
			}
			out.push(symbol);
			if (symbol.kind === 'enum') {
				out.push(...(symbol.children ?? []));
			}
		}
		return out;
	}

	/** Exported module-level declarations (incl. exported enum members). */
	private exportedModuleLevelMatches(mod: ModuleSymbols, lower: string): VbaSymbol[] {
		return this.moduleLevelBindings(mod, lower)
			.filter((binding) => binding.exported)
			.map((binding) => binding.symbol);
	}

	/**
	 * Module-level declaration bindings with their project visibility. Enum
	 * members inherit project visibility from the containing Enum declaration.
	 */
	private moduleLevelBindings(mod: ModuleSymbols, lower: string): ModuleLevelBinding[] {
		const hits: ModuleLevelBinding[] = [];
		for (const symbol of mod.root.children ?? []) {
			if (symbol.name.toLowerCase() === lower) {
				hits.push({ symbol, exported: isExported(symbol, mod.moduleKind) });
			}
			// Enum members are referenceable at module scope by their bare name.
			if (symbol.kind === 'enum') {
				const exported = isEnumMemberExported(symbol, mod.moduleKind);
				for (const member of symbol.children ?? []) {
					if (member.name.toLowerCase() === lower) {
						hits.push({ symbol: member, exported });
					}
				}
			}
		}
		return hits;
	}

	private visibleObjectMembers(mod: ModuleSymbols): VbaProjectClassMember[] {
		return this.visibleProjectMembers(mod, (symbol) => isVisibleProjectObjectMember(symbol));
	}

	private visibleStandardModuleMembers(
		mod: ModuleSymbols,
		sameModule: boolean,
	): VbaProjectClassMember[] {
		return this.visibleProjectMembers(
			mod,
			(symbol) => isVisibleStandardModuleMember(symbol, mod, sameModule),
			{ includeEnumMembers: true },
		);
	}

	private visibleProjectMembers(
		mod: ModuleSymbols,
		isVisible: (symbol: VbaSymbol) => boolean,
		options: { includeEnumMembers?: boolean } = {},
	): VbaProjectClassMember[] {
		const byName = new Map<string, VbaProjectClassMember>();
		for (const symbol of projectMemberCandidateSymbols(mod, options.includeEnumMembers === true)) {
			if (!isVisible(symbol)) {
				continue;
			}
			const kind = projectObjectMemberKind(symbol);
			if (!kind) {
				continue;
			}
			const key = symbol.name.toLowerCase();
			const existing = byName.get(key);
			if (existing) {
				const returns = projectObjectMemberReturnType(symbol);
				if (!existing.returns && returns) {
					existing.returns = returns;
				}
				const writable = projectObjectMemberWritable(symbol);
				if (writable === true) {
					existing.writable = true;
				} else if (existing.writable === undefined && writable === false) {
					existing.writable = false;
				}
				if (!existing.writeType) {
					existing.writeType = projectObjectMemberWriteType(symbol);
				}
				if (!existing.signature) {
					existing.signature = projectObjectMemberSignature(symbol);
				}
				if (!existing.doc && symbol.doc) {
					existing.doc = symbol.doc;
				}
				if (isDefaultProjectObjectMember(symbol)) {
					existing.defaultMember = true;
				}
				existing.attributes = mergeMemberAttributes(existing.attributes, symbol.attributes);
				existing.definitions = [
					...(existing.definitions ?? []),
					projectObjectMemberDefinition(symbol),
				];
				continue;
			}
			byName.set(key, {
				name: symbol.name,
				kind,
				returns: projectObjectMemberReturnType(symbol),
				signature: projectObjectMemberSignature(symbol),
				writable: projectObjectMemberWritable(symbol),
				writeType: projectObjectMemberWriteType(symbol),
				moduleName: mod.moduleName,
				visibility: symbol.visibility,
				doc: symbol.doc,
				definitions: [projectObjectMemberDefinition(symbol)],
				defaultMember: isDefaultProjectObjectMember(symbol) || undefined,
				attributes: mergeMemberAttributes(undefined, symbol.attributes),
			});
		}
		return [...byName.values()];
	}

	private projectUserTypeMembers(moduleName: string): VbaProjectClassMembers[] {
		const currentLower = moduleName.toLowerCase();
		const out: VbaProjectClassMembers[] = [];
		for (const mod of this.modules.values()) {
			const sameModule = mod.moduleName.toLowerCase() === currentLower;
			for (const symbol of mod.root.children ?? []) {
				if (symbol.kind !== 'type') {
					continue;
				}
				if (!sameModule && !isTypeExported(symbol)) {
					continue;
				}
				out.push({
					name: symbol.name,
					kind: 'userType',
					moduleName: mod.moduleName,
					doc: symbol.doc,
					exhaustive: true,
					members: this.userTypeFieldMembers(symbol),
				});
			}
		}
		return out;
	}

	/**
	 * Enum names are member surfaces too: `Corner.TopLeft` is ordinary VBA and is
	 * how a reader tells one enum's TopLeft from another's. Enums were never
	 * emitted, so qualifying by the enum name offered nothing at all.
	 */
	private projectEnumMembers(moduleName: string): VbaProjectClassMembers[] {
		const currentLower = moduleName.toLowerCase();
		const out: VbaProjectClassMembers[] = [];
		for (const mod of this.modules.values()) {
			const sameModule = mod.moduleName.toLowerCase() === currentLower;
			for (const symbol of mod.root.children ?? []) {
				if (symbol.kind !== 'enum') {
					continue;
				}
				if (!sameModule && !isTypeExported(symbol)) {
					continue;
				}
				out.push({
					name: symbol.name,
					kind: 'enum',
					moduleName: mod.moduleName,
					doc: symbol.doc,
					exhaustive: true,
					members: this.enumConstantMembers(symbol),
				});
			}
		}
		return out;
	}

	private enumConstantMembers(symbol: VbaSymbol): VbaProjectClassMember[] {
		return (symbol.children ?? [])
			.filter((member) => member.kind === 'enumMember')
			.map((member) => ({
				name: member.name,
				kind: 'property' as const,
				returns: symbol.name,
				signature: `${symbol.name}.${member.name} As ${symbol.name}`,
				writable: false,
				moduleName: member.moduleName,
				doc: member.doc,
				definitions: [projectObjectMemberDefinition(member)],
			}));
	}

	private userTypeFieldMembers(symbol: VbaSymbol): VbaProjectClassMember[] {
		return (symbol.children ?? [])
			.filter((field) => field.kind === 'typeField')
			.map((field) => ({
				name: field.name,
				kind: 'property' as const,
				returns: field.asType,
				signature: userTypeFieldSignature(field),
				writable: true,
				writeType: field.asType,
				moduleName: field.moduleName,
				doc: field.doc,
				definitions: [projectObjectMemberDefinition(field)],
			}));
	}
}
