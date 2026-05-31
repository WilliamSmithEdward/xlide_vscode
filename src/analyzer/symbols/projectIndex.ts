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
} from './symbolModel';

/** Source text + workbook role for one module fed into the index. */
export interface ModuleInput {
	moduleName: string;
	moduleKind: ModuleSymbolKind;
	source: string;
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
