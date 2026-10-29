// What another VBA project offers the project that references it (issue
// #1503). Measured through the VBE oracle in Excel, Word, PowerPoint and
// Access 16.0 with a referenced project ProbeLibrary whose standard module
// Service has a Public Function Value:
//
// - `ProbeLibrary.Service.Value()`, `ProbeLibrary.Value()`, `Service.Value()`
//   and `Value()` all compile, in any case and bracketed.
// - With `Option Private Module` in Service, none does: the qualified forms
//   are "Method or data member not found", `Service.Value()` "Variable not
//   defined" and `Value()` "Sub or Function not defined". A Private Function
//   is the same as a private module for it.
// - A class of the referenced project is no type the referring one can name
//   unless the class is exposed (VB_Exposed, Instancing PublicNotCreatable):
//   `Dim t As ProbeLibrary.Thing` is "User-defined type not defined".
// - The referring project's own names come first: its own Value() wins.

import { ProjectIndex } from './projectIndex';
import type { ModuleSymbolKind } from './symbolModel';

/** One module of a referenced project, as the referring project sees it. */
export interface ReferencedProjectModule {
	name: string;
	kind: ModuleSymbolKind;
	/** Whether the referring project can name the module at all. */
	visible: boolean;
	/** The names a visible standard module offers, its procedures, variables, constants and enum members. */
	members: string[];
	/** The types it offers by name: its Public Types and Enums, or the class itself when exposed. */
	types: string[];
}

/** A referenced VBA project that was read, under the name the reference gives it. */
export interface ReferencedProjectSurface {
	name: string;
	modules: ReferencedProjectModule[];
}

/** A module of the referenced project as its file holds it, attribute header included. */
export interface ReferencedProjectModuleSource {
	name: string;
	kind: ModuleSymbolKind;
	source: string;
}

const OPTION_PRIVATE_MODULE = /^[ \t]*Option[ \t]+Private[ \t]+Module\b/im;
const EXPOSED = /^[ \t]*Attribute[ \t]+VB_Exposed[ \t]*=[ \t]*True\b/im;
/** No module of the referenced project: < is no identifier character. */
const OUTSIDE = '<outside>';

/** The surface of a referenced project, from its modules' sources. */
export function referencedProjectSurface(name: string, modules: readonly ReferencedProjectModuleSource[]): ReferencedProjectSurface {
	const index = new ProjectIndex();
	for (const mod of modules) {
		try {
			index.setModule({ moduleName: mod.name, moduleKind: mod.kind, source: mod.source });
		} catch {
			// A module that does not parse offers nothing it can be held to.
		}
	}
	// What a module of another project could see, were it in this one.
	const visibleSymbols = index.visibleIdentifierSymbols(OUTSIDE);
	return {
		name,
		modules: modules.map((mod): ReferencedProjectModule => {
			if (mod.kind === 'standard') {
				const visible = !OPTION_PRIVATE_MODULE.test(mod.source);
				const own = visible ? visibleSymbols.filter((symbol) => symbol.moduleName.toLowerCase() === mod.name.toLowerCase()) : [];
				return {
					name: mod.name,
					kind: mod.kind,
					visible,
					members: own.filter((symbol) => symbol.kind !== 'type' && symbol.kind !== 'enum').map((symbol) => symbol.name),
					types: own.filter((symbol) => symbol.kind === 'type' || symbol.kind === 'enum').map((symbol) => symbol.name),
				};
			}
			const exposed = EXPOSED.test(mod.source);
			return { name: mod.name, kind: mod.kind, visible: exposed, members: [], types: exposed ? [mod.name] : [] };
		}),
	};
}
