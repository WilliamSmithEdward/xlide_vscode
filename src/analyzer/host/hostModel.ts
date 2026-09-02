// Host object-model resolver. Pure functions over a HostObjectModel; no vscode
// or I/O dependencies so the resolver is unit-testable. Defaults to the Excel
// object model but accepts any HostObjectModel for testing/extensibility.

import {
	getExcelObjectModel,
	HostConstant,
	HostEnum,
	HostMember,
	HostObjectModel,
	HostType,
} from './excelObjectModel';

function isObjectAccessMember(member: HostMember): boolean {
	return member.kind !== 'event';
}

const HOST_CONSTANT_INDEX = new WeakMap<HostObjectModel, Map<string, HostConstant>>();

function hostConstantIndex(model: HostObjectModel): Map<string, HostConstant> {
	const cached = HOST_CONSTANT_INDEX.get(model);
	if (cached) {
		return cached;
	}
	const index = new Map<string, HostConstant>();
	for (const [key, constant] of Object.entries(model.constants ?? {})) {
		index.set(key.toLowerCase(), constant);
	}
	HOST_CONSTANT_INDEX.set(model, index);
	return index;
}

interface HostTypeIndex {
	members: HostMember[];
	byLowerName: Map<string, HostMember>;
	rawByLowerName: Map<string, HostMember>;
}

interface HostModelIndex {
	membersByType: Map<string, HostTypeIndex>;
	typeKeysByLower: Map<string, string>;
	globalsByLower: Map<string, string>;
	/** Every member name in the model, lowercased, across every type. */
	memberNames: Set<string>;
	/** Lowercased enum name -> the enumeration. */
	enumsByLower: Map<string, HostEnum>;
	/** Lowercased enum name -> its constants, in declaration order. */
	constantsByEnum: Map<string, HostConstant[]>;
}

const HOST_MODEL_INDEX = new WeakMap<HostObjectModel, HostModelIndex>();

function hostModelIndex(model: HostObjectModel): HostModelIndex {
	const cached = HOST_MODEL_INDEX.get(model);
	if (cached) {
		return cached;
	}
	const membersByType = new Map<string, HostTypeIndex>();
	const typeKeysByLower = new Map<string, string>();
	for (const [key, type] of Object.entries(model.types)) {
		const keyLower = key.toLowerCase();
		if (!typeKeysByLower.has(keyLower)) {
			typeKeysByLower.set(keyLower, key);
		}
		const members: HostMember[] = [];
		const byLowerName = new Map<string, HostMember>();
		const rawByLowerName = new Map<string, HostMember>();
		for (const member of type.members ?? []) {
			const lower = member.name.toLowerCase();
			if (!rawByLowerName.has(lower)) {
				rawByLowerName.set(lower, member);
			}
			if (!isObjectAccessMember(member)) {
				continue;
			}
			members.push(member);
			if (!byLowerName.has(lower)) {
				byLowerName.set(lower, member);
			}
		}
		membersByType.set(key, { members, byLowerName, rawByLowerName });
	}
	const globalsByLower = new Map<string, string>();
	for (const [key, type] of Object.entries(model.globals)) {
		const keyLower = key.toLowerCase();
		if (!globalsByLower.has(keyLower)) {
			globalsByLower.set(keyLower, type);
		}
	}
	const memberNames = new Set<string>();
	for (const type of membersByType.values()) {
		for (const lower of type.rawByLowerName.keys()) {
			memberNames.add(lower);
		}
	}
	const enumsByLower = new Map<string, HostEnum>();
	for (const entry of Object.values(model.enums ?? {})) {
		const lower = entry.displayName.toLowerCase();
		if (!enumsByLower.has(lower)) {
			enumsByLower.set(lower, entry);
		}
	}
	// An enum's members are the constants that name it, so the two can never
	// disagree and the generated tables stay a single list.
	const constantsByEnum = new Map<string, HostConstant[]>();
	for (const constant of Object.values(model.constants ?? {})) {
		if (!constant.type) {
			continue;
		}
		const lower = constant.type.toLowerCase();
		const bucket = constantsByEnum.get(lower);
		if (bucket) {
			bucket.push(constant);
		} else {
			constantsByEnum.set(lower, [constant]);
		}
	}
	const index = {
		membersByType,
		typeKeysByLower,
		globalsByLower,
		memberNames,
		enumsByLower,
		constantsByEnum,
	};
	HOST_MODEL_INDEX.set(model, index);
	return index;
}

/** The enumerations the host library declares, for type positions and hover. */
export function getHostEnums(
	model: HostObjectModel = getExcelObjectModel(),
): HostEnum[] {
	return Object.values(model.enums ?? {});
}

/**
 * The enumeration named `name`, case-insensitively. VBA accepts an enum name as
 * a declared type (`Dim k As XlAxisType`) and as a qualifier
 * (`XlAxisType.xlCategory`); neither resolved to anything before.
 */
export function resolveHostEnum(
	name: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostEnum | undefined {
	if (!name) {
		return undefined;
	}
	return hostModelIndex(model).enumsByLower.get(name.trim().toLowerCase());
}

/** The constants belonging to one enumeration, in declaration order. */
export function getHostEnumMembers(
	enumName: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostConstant[] {
	return hostModelIndex(model).constantsByEnum.get(enumName.toLowerCase()) ?? [];
}

/**
 * True when any type in the model carries a member of this name. Cheap enough
 * to run per member-access dot in a module, which is what semantic-token
 * painting does before paying for a receiver walk that could only fail.
 */
export function isHostMemberNameAnywhere(
	name: string,
	model: HostObjectModel = getExcelObjectModel(),
): boolean {
	return hostModelIndex(model).memberNames.has(name.toLowerCase());
}

/** Returns the type metadata for a qualified type name (e.g. "Excel.Range"). */
export function getHostType(
	qualified: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostType | undefined {
	return model.types[qualified];
}

/**
 * The application name origin labels use ("Excel host method", "Word type").
 * An absent model answers Excel - the default model when no host is named -
 * and so does a model without a hostName, keeping the historical wording for
 * models that predate the field (issue #28).
 */
export function hostDisplayName(model?: HostObjectModel): string {
	return model?.hostName ?? 'Excel';
}

/**
 * Resolves an object-access member (a property or method, never an event) of
 * a qualified host type by name. Case-insensitive. Undefined when the type or
 * the member is unknown to the model.
 */
export function resolveHostMember(
	qualified: string,
	memberName: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostMember | undefined {
	return hostModelIndex(model).membersByType
		.get(qualified)?.byLowerName.get(memberName.toLowerCase());
}

/** Returns the members of a qualified type, or an empty array if unknown. */
export function getHostMembers(
	qualified: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostMember[] {
	return hostModelIndex(model).membersByType.get(qualified)?.members ?? [];
}

/**
 * The events a qualified type raises, in declaration order. Events are not
 * object-access members (`Form.Load()` is not a call), so `getHostMembers`
 * leaves them out; a form's handler stubs (Form_Load, Command1_Click) are
 * what reads them. Empty for a type the model does not carry or whose
 * events it does not model.
 */
export function getHostEvents(
	qualified: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostMember[] {
	const key = hostModelIndex(model).typeKeysByLower.get(qualified.toLowerCase());
	const type = key === undefined ? undefined : model.types[key];
	return (type?.members ?? []).filter((member) => member.kind === 'event');
}

const HOST_MEMBER_NAMES = new WeakMap<HostObjectModel, Set<string>>();

/**
 * True when `name` is a member of ANY type in the host object model.
 *
 * Callers reasoning about a late-bound receiver cannot know its runtime type,
 * so they need the weaker question "could this name legally dispatch somewhere
 * in the host model at all?". Answering yes keeps them quiet; the set is
 * deliberately broad for that reason. Case-insensitive.
 */
export function isHostMemberName(
	name: string,
	model: HostObjectModel = getExcelObjectModel(),
): boolean {
	let names = HOST_MEMBER_NAMES.get(model);
	if (!names) {
		names = new Set<string>();
		for (const type of hostModelIndex(model).membersByType.values()) {
			for (const lower of type.byLowerName.keys()) { names.add(lower); }
			for (const lower of type.rawByLowerName.keys()) { names.add(lower); }
		}
		HOST_MEMBER_NAMES.set(model, names);
	}
	return names.has(name.toLowerCase());
}

/**
 * Resolves a bare identifier as a member of the host's hidden Global
 * interface - the surface VBA calls unqualified: Word's InchesToPoints,
 * Excel's Union (issue #34). Object-access members only (never an event),
 * case-insensitive; undefined when the model carries no Global type or the
 * name is not among its members.
 */
export function resolveHostGlobalMember(
	name: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostMember | undefined {
	if (!model.globalType || isHiddenDispatchName(name)) {
		return undefined;
	}
	return hostModelIndex(model).membersByType
		.get(model.globalType)?.byLowerName.get(name.toLowerCase());
}

/** All object-access members of the host's hidden Global interface. */
export function getHostGlobalMembers(
	model: HostObjectModel = getExcelObjectModel(),
): HostMember[] {
	return model.globalType
		? getHostMembers(model.globalType, model).filter((member) => !isHiddenDispatchName(member.name))
		: [];
}

/**
 * `_Run2`, `_Default` and friends: the type library's dispatch internals, which
 * the VBE's own completion hides and no one writes by hand. They are members
 * like any other for a qualified lookup, but they must not become bare-callable
 * names just because the Global interface carries them (issue #41).
 */
function isHiddenDispatchName(name: string): boolean {
	return name.startsWith('_');
}

/**
 * Resolves a host-injected global identifier (ThisWorkbook, Application, ...)
 * to its qualified host type. Case-insensitive.
 */
export function resolveHostGlobal(
	name: string,
	model: HostObjectModel = getExcelObjectModel(),
): string | undefined {
	// O(1) via the WeakMap-cached, lowercase-keyed index (first-wins), instead of
	// an Object.entries allocation + linear scan on every identifier-token lookup.
	return hostModelIndex(model).globalsByLower.get(name.toLowerCase());
}

/** Resolves a host enum constant such as `xlUp` or `xlCalculationAutomatic`. */
export function resolveHostConstant(
	name: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostConstant | undefined {
	return hostConstantIndex(model).get(name.toLowerCase());
}

/**
 * Returns the verified call signature for a callable member of a host type,
 * first from generated member metadata, then from the curated fallback table.
 * Returns undefined when no signature is known. Case-insensitive on the member
 * name.
 */
export function resolveHostMemberSignature(
	qualified: string,
	member: string,
	model: HostObjectModel = getExcelObjectModel(),
): string | undefined {
	const lower = member.toLowerCase();
	const typeIndex = hostModelIndex(model).membersByType.get(qualified);
	const rawMember = typeIndex?.rawByLowerName.get(lower);
	if (rawMember?.kind === 'event') {
		return undefined;
	}
	return (
		typeIndex?.byLowerName.get(lower)?.signature ??
			model.memberSignatures?.[qualified]?.[lower]
	);
}

/** A host-injected global identifier and the qualified type it denotes. */
export interface HostGlobal {
	name: string;
	type: string;
}

/** Returns all host-injected globals (canonical casing) of the model. */
export function getHostGlobals(
	model: HostObjectModel = getExcelObjectModel(),
): HostGlobal[] {
	return Object.entries(model.globals).map(([name, type]) => ({ name, type }));
}

/** Returns all host enum constants (canonical casing) of the model. */
export function getHostConstants(
	model: HostObjectModel = getExcelObjectModel(),
): HostConstant[] {
	return Object.values(model.constants ?? {});
}

/**
 * Resolves a declared type name (as written after `As`) to a qualified host
 * type. Accepts both bare names (Worksheet) and already-qualified names
 * (Excel.Worksheet). Case-insensitive. Returns undefined for non-host types.
 */
export function resolveHostAlias(
	typeName: string,
	model: HostObjectModel = getExcelObjectModel(),
): string | undefined {
	if (!typeName) {
		return undefined;
	}
	const trimmed = typeName.trim();
	const lower = trimmed.toLowerCase();
	if (model.types[trimmed]) {
		return trimmed;
	}
	// Match qualified form case-insensitively.
	const key = hostModelIndex(model).typeKeysByLower.get(lower);
	if (key) {
		return key;
	}
	return model.aliases[lower];
}

/**
 * Returns the qualified type produced by accessing `memberName` on `qualified`,
 * or undefined when the member is unknown or has no chainable return type.
 * Case-insensitive on the member name.
 */
export function resolveMemberReturnType(
	qualified: string,
	memberName: string,
	model: HostObjectModel = getExcelObjectModel(),
): string | undefined {
	return resolveHostMember(qualified, memberName, model)?.returns;
}
