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

import { buildModuleSymbols } from './buildModuleSymbols';
import {
	isProcedureKind,
	type ModuleSymbolKind,
	type ModuleSymbols,
	type VbaSymbol,
	type VbaProcedureSignature,
} from './symbolModel';
import type { Span } from '../parser/nodes';

/** Source text + workbook role for one module fed into the index. */
export interface ModuleInput {
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	source: string;
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
 * A symbol declared at module level (procedure/variable/const/type/enum/declare)
 * is "exported" (visible to other modules) when it is explicitly Public/Global,
 * or when it is a procedure with no visibility keyword (procedures default to
 * Public in standard modules - MS-VBAL 5.3.1.1). Dim/Private/Friend/Static and
 * unmodified module variables/consts stay module-private.
 */
function isExported(symbol: VbaSymbol): boolean {
	if (symbol.visibility === 'Public' || symbol.visibility === 'Global') {
		return true;
	}
	if (symbol.visibility) {
		return false;
	}
	return isProcedureKind(symbol.kind);
}

/** A project-wide symbol index built from a set of module sources. */
export class ProjectIndex {
	private readonly modules = new Map<string, ModuleSymbols>();

	/** Adds or replaces a module in the index. */
	setModule(input: ModuleInput): void {
		const symbols = buildModuleSymbols(
			input.moduleName,
			input.moduleKind,
			input.source,
		);
		this.modules.set(input.moduleName.toLowerCase(), symbols);
	}

	/** Removes a module from the index. */
	removeModule(moduleName: string): void {
		this.modules.delete(moduleName.toLowerCase());
	}

	/** All module names currently indexed (original casing). */
	moduleNames(): string[] {
		return [...this.modules.values()].map((m) => m.moduleName);
	}

	/**
	 * Lowercased names of every procedure (Sub/Function/Property) across all
	 * indexed modules. Used by the unknown-call diagnostic to decide whether a
	 * bare call statement names a procedure that exists anywhere in the project.
	 */
	procedureNames(): Set<string> {
		const names = new Set<string>();
		for (const mod of this.modules.values()) {
			for (const symbol of mod.all) {
				if (isProcedureKind(symbol.kind)) {
					names.add(symbol.name.toLowerCase());
				}
			}
		}
		return names;
	}

	/**
	 * Exported Sub/Function signatures grouped by lowercased procedure name.
	 * Properties are deliberately excluded from this first callable-signature
	 * surface because their invocation syntax and Let/Set/Get pairing needs the
	 * object/member binder.
	 */
	procedureSignatures(): Map<string, VbaProcedureSignature[]> {
		const signatures = new Map<string, VbaProcedureSignature[]>();
		for (const mod of this.modules.values()) {
			for (const symbol of mod.root.children ?? []) {
				if (
					(symbol.kind !== 'sub' && symbol.kind !== 'function') ||
					!isExported(symbol)
				) {
					continue;
				}
				const params = (symbol.children ?? [])
					.filter((child) => child.kind === 'parameter')
					.map((child) => ({
						name: child.name,
						type: child.asType,
						optional: child.optional ?? false,
						paramArray: child.paramArray ?? false,
						isArray: child.isArray ?? false,
					}));
				const sig: VbaProcedureSignature = {
					name: symbol.name,
					moduleName: symbol.moduleName,
					kind: symbol.kind,
					params,
					returnType: symbol.asType,
				};
				const key = symbol.name.toLowerCase();
				const existing = signatures.get(key);
				if (existing) {
					existing.push(sig);
				} else {
					signatures.set(key, [sig]);
				}
			}
		}
		return signatures;
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
				if (isExported(symbol)) {
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
		return this.moduleLevelMatches(mod, name.toLowerCase()).filter(isExported);
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
				if (moduleHits.some(isExported)) {
					return this.projectScope(lower, moduleHits.filter(isExported));
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
				if (isExported(symbol)) {
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
				moduleHits.length > 0 && !moduleHits.some(isExported);
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
}
