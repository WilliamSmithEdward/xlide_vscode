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
// - An exposed class is a type another project can name but not create:
//   `New ProbeLibrary.Widget` is "Invalid use of New keyword". Its name is a
//   value only as a default instance (VB_PredeclaredId): otherwise
//   `v = ProbeLibrary.Widget` is "Method or data member not found" and
//   `v = Widget` "Variable not defined" (issue #1602).

import { libraryTypeNames } from '../host/libraryTypeNames';
import { tokenName } from '../lexer/tokenHelpers';
import { tokenize } from '../lexer/tokenize';
import { isDefaultProjectObjectMember, ProjectIndex } from './projectIndex';
import { procedureSignatureFromSymbol, type ModuleSymbolKind, type VbaSymbol } from './symbolModel';

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
	/** Its Public Enums with their members: `WidgetColor.wcRed`, and in a class `ProbeLibrary.wcRed`. */
	enums?: ReferencedEnum[];
	/** Whether an exposed class's name is its default instance. */
	predeclared?: boolean;
	/** What an object of an exposed class offers another project. */
	objectMembers?: ReferencedObjectMember[];
	/** Whether the class has a default member, which an index into an object reaches. */
	defaultMember?: boolean;
}

/** A Public member of an exposed class of a referenced project. */
export interface ReferencedObjectMember {
	name: string;
	kind: 'method' | 'property';
	/** The exposed class of the same project it returns, `Public Child As Widget`. */
	returnsClass?: string;
	/** Whether it declares parameters, so arguments in a call are its own. */
	takesArguments: boolean;
}

/** A Public Enum of a referenced project's module. */
export interface ReferencedEnum {
	name: string;
	members: string[];
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
const PREDECLARED = /^[ \t]*Attribute[ \t]+VB_PredeclaredId[ \t]*=[ \t]*True\b/im;
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
	const exposedClasses = new Map(modules.filter((mod) => mod.kind === 'class' && EXPOSED.test(mod.source)).map((mod) => [mod.name.toLowerCase(), mod.name]));
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
					enums: publicEnums(own),
					types: own.filter((symbol) => symbol.kind === 'type' || symbol.kind === 'enum').map((symbol) => symbol.name),
				};
			}
			const exposed = EXPOSED.test(mod.source);
			if (!exposed) {
				return { name: mod.name, kind: mod.kind, visible: false, members: [], types: [] };
			}
			// A document module or form is a class module too: its name is a
			// type and its Public Enums are offered as an exposed class's are.
			const symbols = index.getModule(mod.name)?.root.children ?? [];
			const enums = publicEnums(symbols);
			return {
				name: mod.name,
				kind: mod.kind,
				visible: true,
				members: [],
				types: [mod.name, ...enums.map((decl) => decl.name)],
				enums,
				// A document module or form's members are its host object's too: unjudged.
				...(mod.kind === 'class' ? {
					predeclared: PREDECLARED.test(mod.source),
					objectMembers: objectMembersOf(symbols, name, exposedClasses),
					defaultMember: symbols.some(isDefaultProjectObjectMember),
				} : {}),
			};
		}),
	};
}

/** Existing name lookups shared by all rules analyzing the same surface. */
export interface ReferencedProjectNames {
	/** First module of each name; callers still check its visibility. */
	modules: ReadonlyMap<string, ReferencedProjectModule>;
	/** Types offered by the surface, by lowercased name. */
	types: ReadonlyMap<string, ReferencedProjectModule>;
	/** Members of visible standard modules and exposed classes' Enum members, available after the project name. */
	values: ReadonlySet<string>;
	/** The Public Enums of visible modules, by lowercased name. */
	enums: ReadonlyMap<string, ReferencedEnum>;
	/** Visible module names and their members, available unqualified. */
	bareNames: ReadonlySet<string>;
}

const PROJECT_NAMES = new WeakMap<ReferencedProjectSurface, ReferencedProjectNames>();

/**
 * Built once per surface object. Surfaces are snapshots: replace the object
 * when a referenced project changes. Worker requests receive fresh copies,
 * so this shares work within an analysis, not across worker requests.
 */
export function referencedProjectNames(surface: ReferencedProjectSurface): ReferencedProjectNames {
	let names = PROJECT_NAMES.get(surface);
	if (!names) {
		const modules = new Map<string, ReferencedProjectModule>();
		const types = new Map<string, ReferencedProjectModule>();
		const values = new Set<string>();
		const bareNames = new Set<string>();
		const enums = new Map<string, ReferencedEnum>();
		for (const module of surface.modules) {
			const key = module.name.toLowerCase();
			if (!modules.has(key)) {
				modules.set(key, module);
			}
			for (const type of module.types) {
				if (!types.has(type.toLowerCase())) {
					types.set(type.toLowerCase(), module);
				}
			}
			if (!module.visible) {
				continue;
			}
			// A class with no default instance is a type, never a value by name.
			if (module.kind !== 'class' || module.predeclared) {
				bareNames.add(key);
			}
			for (const member of module.members) {
				bareNames.add(member.toLowerCase());
				if (module.kind === 'standard') {
					values.add(member.toLowerCase());
				}
			}
			for (const decl of module.enums ?? []) {
				bareNames.add(decl.name.toLowerCase());
				if (!enums.has(decl.name.toLowerCase())) {
					enums.set(decl.name.toLowerCase(), decl);
				}
				// A standard module's Enum members are among its members already.
				for (const member of module.kind === 'standard' ? [] : decl.members) {
					bareNames.add(member.toLowerCase());
					values.add(member.toLowerCase());
				}
			}
		}
		names = { modules, types, values, enums, bareNames };
		PROJECT_NAMES.set(surface, names);
	}
	return names;
}

/**
 * The read project a bare name reaches. A name is looked up in VBA, the host's
 * library and then the references in the project's order, and the first that
 * has it wins. Undefined where a library before the project might have it: one
 * whose type names are known and hold it, or one XLIDE did not read, whose
 * names are unknown. `offers` says whether a read project has the name.
 */
export function referencedProjectReached(
	lower: string,
	surfaces: readonly ReferencedProjectSurface[] | undefined,
	order: ReferencedLookupOrder,
	offers: (names: ReferencedProjectNames) => boolean,
): ReferencedProjectSurface | undefined {
	if (!surfaces?.length || !order.libraries) {
		return undefined;
	}
	for (const library of ['VBA', order.hostName, ...order.libraries]) {
		const key = library.toLowerCase();
		const surface = surfaces.find((candidate) => candidate.name.toLowerCase() === key);
		if (surface) {
			if (offers(referencedProjectNames(surface))) {
				return surface;
			}
			continue;
		}
		const known = libraryTypeNames(library);
		if (!known || known.has(lower)) {
			return undefined;
		}
	}
	return undefined;
}

/**
 * Whether a declared type names a read referenced project's Public Enum:
 * `ProbeLibrary.WidgetColor`, or a bare `WidgetColor` the lookup reaches. A
 * variable of it is a Long, so a member on it is "Invalid qualifier" (issue
 * #1623). A name the project itself declares is the caller's to rule out.
 */
export function declaredTypeIsReferencedEnum(
	asType: string,
	surfaces: readonly ReferencedProjectSurface[] | undefined,
	order: ReferencedLookupOrder | undefined,
): boolean {
	const parts = asType.split('.').map((part) => part.trim().toLowerCase());
	if (!surfaces?.length || parts.length > 2 || parts.some((part) => !/^[\p{L}_][\p{L}\p{M}0-9_]*$/u.test(part))) {
		return false;
	}
	const name = parts[parts.length - 1];
	const project = parts.length === 2
		? surfaces.find((surface) => surface.name.toLowerCase() === parts[0])
		: order && referencedProjectReached(name, surfaces, order, (offered) => offered.types.has(name));
	return !!project && referencedProjectNames(project).enums.has(name);
}

/** What decides which library a bare name reaches. */
export interface ReferencedLookupOrder {
	/** The project's references, in order; undefined where they are not known. */
	libraries: readonly string[] | undefined;
	hostName: string;
}

/** What an object of a class offers another project: Public procedures and variables. Friend, Private and events are none. */
function objectMembersOf(symbols: readonly VbaSymbol[], projectName: string, exposedClasses: ReadonlyMap<string, string>): ReferencedObjectMember[] {
	const out = new Map<string, ReferencedObjectMember>();
	for (const symbol of symbols) {
		const kind = objectMemberKind(symbol);
		if (!kind) {
			continue;
		}
		const key = symbol.name.toLowerCase();
		const member = out.get(key) ?? { name: symbol.name, kind, takesArguments: false };
		out.set(key, member);
		// A Property Let or Set takes the value; what the member gives is its Get's.
		if (symbol.kind !== 'propertyLet' && symbol.kind !== 'propertySet') {
			member.returnsClass = exposedClassNamed(symbol.asType, projectName, exposedClasses);
			member.takesArguments = (procedureSignatureFromSymbol(symbol)?.params.length ?? 0) > 0;
		}
	}
	return [...out.values()];
}

function objectMemberKind(symbol: VbaSymbol): ReferencedObjectMember['kind'] | undefined {
	switch (symbol.kind) {
		case 'sub':
		case 'function':
			return symbol.visibility === 'Private' || symbol.visibility === 'Friend' ? undefined : 'method';
		case 'propertyGet':
		case 'propertyLet':
		case 'propertySet':
			return symbol.visibility === 'Private' || symbol.visibility === 'Friend' ? undefined : 'property';
		case 'moduleVariable':
			return symbol.visibility === 'Public' || symbol.visibility === 'Global' ? 'property' : undefined;
		default:
			return undefined;
	}
}

/** The same project's exposed class a declared type names, bare or after the project's name. */
function exposedClassNamed(asType: string | undefined, projectName: string, exposedClasses: ReadonlyMap<string, string>): string | undefined {
	const names = asType === undefined ? undefined : declaredTypeNames(asType);
	if (!names || names.length > 2 || (names.length === 2 && names[0].toLowerCase() !== projectName.toLowerCase())) {
		return undefined;
	}
	return exposedClasses.get(names[names.length - 1].toLowerCase());
}

/**
 * The names of a declared type as the lexer reads them, `ProbeLibrary.Widget`:
 * across any line continuation it accepts, and with brackets off. Undefined
 * unless the text is names joined by dots.
 */
export function declaredTypeNames(text: string): string[] | undefined {
	const tokens = tokenize(text).filter((token) => token.kind !== 'newline');
	const names: string[] = [];
	for (let i = 0; i < tokens.length; i += 2) {
		const name = tokenName(tokens[i]);
		if (name === undefined || (i > 0 && tokens[i - 1].rawText !== '.')) {
			return undefined;
		}
		names.push(name);
	}
	return names.length > 0 && tokens.length % 2 === 1 ? names : undefined;
}

/** The Public Enums among a module's symbols, each with its members. */
function publicEnums(symbols: readonly VbaSymbol[]): ReferencedEnum[] {
	return symbols.filter((symbol) => symbol.kind === 'enum' && symbol.visibility !== 'Private').map((symbol) => ({
		name: symbol.name,
		members: (symbol.children ?? []).filter((member) => member.kind === 'enumMember').map((member) => member.name),
	}));
}
