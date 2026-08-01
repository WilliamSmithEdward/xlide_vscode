// Host object-model resolver. Pure functions over a HostObjectModel; no vscode
// or I/O dependencies so the resolver is unit-testable. Defaults to the Excel
// object model but accepts any HostObjectModel for testing/extensibility.

import {
	getExcelObjectModel,
	HostConstant,
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
	const index = { membersByType, typeKeysByLower, globalsByLower };
	HOST_MODEL_INDEX.set(model, index);
	return index;
}

/** Returns the type metadata for a qualified type name (e.g. "Excel.Range"). */
export function getHostType(
	qualified: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostType | undefined {
	return model.types[qualified];
}

/** Returns the members of a qualified type, or an empty array if unknown. */
export function getHostMembers(
	qualified: string,
	model: HostObjectModel = getExcelObjectModel(),
): HostMember[] {
	return hostModelIndex(model).membersByType.get(qualified)?.members ?? [];
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
	const member = hostModelIndex(model).membersByType
		.get(qualified)?.byLowerName.get(memberName.toLowerCase());
	return member?.returns;
}
