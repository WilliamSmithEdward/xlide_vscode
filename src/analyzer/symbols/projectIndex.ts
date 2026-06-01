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
	procedureParamsFromSymbol,
	procedureSignatureFromSymbol,
} from './symbolModel';
import type { Span } from '../parser/nodes';

/** Source text + workbook role for one module fed into the index. */
export interface ModuleInput {
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	source: string;
	/** Optional per-module conditional-compilation environment. */
	conditionalCompilation?: BuildModuleSymbolsOptions['conditionalCompilation'];
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
 *   - `project` - an exported (Public/Global, or default-Public procedure)
 *                 declaration; search spans every module except those that
 *                 re-declare the name privately at module level, and excludes
 *                 procedures whose locals shadow the name ({@link shadowedSpans}).
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


/**
 * A symbol declared at module level is "exported" for cross-module lookup when
 * it is explicitly Public/Global, or when it is an unmodified procedure in a
 * standard module (procedures default to Public there - MS-VBAL 5.3.1.1).
 * Dim/Private/Friend/Static and unmodified module variables/consts stay
 * module-private.
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

function isVisibleProjectObjectMember(symbol: VbaSymbol): boolean {
	if (isProcedureKind(symbol.kind)) {
		return symbol.visibility !== 'Private';
	}
	if (symbol.kind === 'moduleVariable') {
		return symbol.visibility === 'Public' || symbol.visibility === 'Global';
	}
	return false;
}

function projectObjectMemberKind(symbol: VbaSymbol): VbaProjectClassMember['kind'] | undefined {
	switch (symbol.kind) {
		case 'sub':
		case 'function':
			return 'method';
		case 'propertyGet':
		case 'propertyLet':
		case 'propertySet':
		case 'moduleVariable':
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

function projectObjectMemberDefinition(symbol: VbaSymbol): VbaProjectClassMemberDefinition {
	return {
		moduleName: symbol.moduleName,
		nameSpan: symbol.nameSpan,
		fullSpan: symbol.fullSpan,
	};
}

function isDefaultMemberAttribute(attr: VbaSymbolAttribute): boolean {
	return attr.name.toLowerCase() === 'vb_usermemid' && attr.valueRaw.trim() === '0';
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
		const match = /^\s*Implements\s+([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)\b/i.exec(
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
	if (
		symbol.kind !== 'sub' &&
		symbol.kind !== 'function' &&
		symbol.kind !== 'propertyGet'
	) {
		return undefined;
	}
	const params = procedureParamsFromSymbol(symbol)
		.map((param) => formatProcedureParamLabel(param))
		.join(', ');
	const returns = symbol.asType ? ` As ${symbol.asType}` : '';
	return `${symbol.name}(${params})${returns}`;
}

/** A project-wide symbol index built from a set of module sources. */
export class ProjectIndex {
	private readonly modules = new Map<string, ModuleSymbols>();
	private readonly moduleSources = new Map<string, string>();

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
	}

	/** Removes a module from the index. */
	removeModule(moduleName: string): void {
		const key = moduleName.toLowerCase();
		this.modules.delete(key);
		this.moduleSources.delete(key);
	}

	/** All module names currently indexed (original casing). */
	moduleNames(): string[] {
		return [...this.modules.values()].map((m) => m.moduleName);
	}

	/**
	 * Lowercased names of every procedure/Declare callable across all
	 * indexed modules. Used by the unknown-call diagnostic to decide whether a
	 * bare call statement names a procedure that exists anywhere in the project.
	 */
	procedureNames(): Set<string> {
		const names = new Set<string>();
		for (const mod of this.modules.values()) {
			for (const symbol of mod.all) {
				if (isBareCallableKind(symbol.kind) || isProcedureKind(symbol.kind)) {
					names.add(symbol.name.toLowerCase());
				}
			}
		}
		return names;
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
	}

	/**
	 * Visible bare-call Sub/Function/Declare signatures from `moduleName`.
	 * Same-module callables are visible to their own module. Other modules
	 * contribute only exported standard-module callables, matching the
	 * `visibleProcedureNames` rule used by diagnostics.
	 */
	visibleProcedureSignatures(moduleName: string): VbaProcedureSignature[] {
		const currentLower = moduleName.toLowerCase();
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
		const names = new Set<string>();
		for (const mod of this.modules.values()) {
			const sameModule = mod.moduleName.toLowerCase() === currentLower;
			if (mod.moduleKind === 'document' || mod.moduleKind === 'userform') {
				names.add(mod.moduleName.toLowerCase());
			}
			for (const symbol of mod.root.children ?? []) {
				if (!this.isBareIdentifierVisible(symbol, mod, sameModule)) {
					continue;
				}
				names.add(symbol.name.toLowerCase());
				if (symbol.kind === 'enum') {
					for (const member of symbol.children ?? []) {
						names.add(member.name.toLowerCase());
					}
				}
			}
		}
		return names;
	}

	/**
	 * Lowercased visible declaration names that are known not to be type names.
	 * Used by type-position diagnostics after the type resolver has failed, so
	 * project/primitive/host type names still take precedence over value names.
	 */
	visibleNonTypeNames(moduleName: string): Set<string> {
		const currentLower = moduleName.toLowerCase();
		const names = new Set<string>();
		for (const mod of this.modules.values()) {
			const sameModule = mod.moduleName.toLowerCase() === currentLower;
			for (const symbol of mod.root.children ?? []) {
				if (!this.isBareIdentifierVisible(symbol, mod, sameModule)) {
					continue;
				}
				if (symbol.kind === 'enum') {
					for (const member of symbol.children ?? []) {
						names.add(member.name.toLowerCase());
					}
					continue;
				}
				if (projectTypeKind(symbol)) {
					continue;
				}
				names.add(symbol.name.toLowerCase());
			}
		}
		return names;
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
		const out: VbaProjectClassMembers[] = [];
		for (const mod of this.modules.values()) {
			const kind = moduleKindAsTypeName(mod.moduleKind);
			if (kind !== 'class' && kind !== 'document' && kind !== 'userform') {
				continue;
			}
			const members = this.visibleObjectMembers(mod);
			out.push({
				name: mod.moduleName,
				kind,
				moduleName: mod.moduleName,
				implements: moduleImplements(this.moduleSources.get(mod.moduleName.toLowerCase()) ?? ''),
				doc: mod.root.doc,
				exhaustive: kind === 'class',
				members,
			});
		}
		return out;
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
		const lower = name.toLowerCase();
		const home = this.modules.get(moduleName.toLowerCase());

		if (home) {
			const enclosing = this.enclosingProcedure(home, offset);
			if (enclosing) {
				const localHits = (enclosing.children ?? []).filter(
					(c) =>
						(c.kind === 'parameter' ||
							c.kind === 'localVariable' ||
							c.kind === 'constant') &&
						c.name.toLowerCase() === lower,
				);
				if (localHits.length > 0) {
					return localHits;
				}
			}

			const moduleHits = this.moduleLevelMatches(home, lower);
			if (moduleHits.length > 0) {
				return moduleHits;
			}
		}

		const exported: VbaSymbol[] = [];
		for (const mod of this.modules.values()) {
			if (mod.moduleName.toLowerCase() === moduleName.toLowerCase()) {
				continue;
			}
			for (const symbol of this.moduleLevelMatches(mod, lower)) {
				if (isExported(symbol, mod.moduleKind)) {
					exported.push(symbol);
				}
			}
		}
		return exported;
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
		return this.moduleLevelMatches(mod, name.toLowerCase()).filter((symbol) =>
			isExported(symbol, mod.moduleKind),
		);
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

		if (home) {
			const enclosing = this.enclosingProcedure(home, offset);
			if (enclosing) {
				const localHits = (enclosing.children ?? []).filter(
					(c) =>
						(c.kind === 'parameter' ||
							c.kind === 'localVariable' ||
							c.kind === 'constant') &&
						c.name.toLowerCase() === lower,
				);
				if (localHits.length > 0) {
					return {
						kind: 'local',
						definitions: localHits,
						searchModules: [home.moduleName],
						procedureSpan: enclosing.fullSpan,
						shadowedSpans: [],
					};
				}
			}

			const moduleHits = this.moduleLevelMatches(home, lower);
			if (moduleHits.length > 0) {
				if (moduleHits.some((symbol) => isExported(symbol, home.moduleKind))) {
					return this.projectScope(
						lower,
						moduleHits.filter((symbol) => isExported(symbol, home.moduleKind)),
					);
				}
				return {
					kind: 'module',
					definitions: moduleHits,
					searchModules: [home.moduleName],
					shadowedSpans: this.localShadowSpans(home, lower),
				};
			}
		}

		const exported: VbaSymbol[] = [];
		for (const mod of this.modules.values()) {
			if (home && mod.moduleName.toLowerCase() === moduleName.toLowerCase()) {
				continue;
			}
			for (const symbol of this.moduleLevelMatches(mod, lower)) {
				if (isExported(symbol, mod.moduleKind)) {
					exported.push(symbol);
				}
			}
		}
		if (exported.length > 0) {
			return this.projectScope(lower, exported);
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
			const moduleHits = this.moduleLevelMatches(mod, lower);
			const privatelyShadowed =
				moduleHits.length > 0 &&
				!moduleHits.some((symbol) => isExported(symbol, mod.moduleKind));
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

	/** Module-level declarations (incl. enum members) matching a lowercased name. */
	private moduleLevelMatches(mod: ModuleSymbols, lower: string): VbaSymbol[] {
		const hits: VbaSymbol[] = [];
		for (const symbol of mod.root.children ?? []) {
			if (symbol.name.toLowerCase() === lower) {
				hits.push(symbol);
			}
			// Enum members are referenceable at module scope by their bare name.
			if (symbol.kind === 'enum') {
				for (const member of symbol.children ?? []) {
					if (member.name.toLowerCase() === lower) {
						hits.push(member);
					}
				}
			}
		}
		return hits;
	}

	private visibleObjectMembers(mod: ModuleSymbols): VbaProjectClassMember[] {
		const byName = new Map<string, VbaProjectClassMember>();
		for (const symbol of mod.root.children ?? []) {
			if (!isVisibleProjectObjectMember(symbol)) {
				continue;
			}
			const kind = projectObjectMemberKind(symbol);
			if (!kind) {
				continue;
			}
			const key = symbol.name.toLowerCase();
			const existing = byName.get(key);
			if (existing) {
				if (!existing.returns && symbol.asType) {
					existing.returns = symbol.asType;
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
				returns: symbol.asType,
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
}
