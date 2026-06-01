// Host object-model resolver. Pure functions over a HostObjectModel; no vscode
// or I/O dependencies so the resolver is unit-testable. Defaults to the Excel
// object model but accepts any HostObjectModel for testing/extensibility.

import {
	EXCEL_OBJECT_MODEL,
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

/** Returns the type metadata for a qualified type name (e.g. "Excel.Range"). */
export function getHostType(
	qualified: string,
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
): HostType | undefined {
	return model.types[qualified];
}

/** Returns the members of a qualified type, or an empty array if unknown. */
export function getHostMembers(
	qualified: string,
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
): HostMember[] {
	return (model.types[qualified]?.members ?? []).filter(isObjectAccessMember);
}

/**
 * Resolves a host-injected global identifier (ThisWorkbook, Application, ...)
 * to its qualified host type. Case-insensitive.
 */
export function resolveHostGlobal(
	name: string,
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
): string | undefined {
	const lower = name.toLowerCase();
	for (const [key, type] of Object.entries(model.globals)) {
		if (key.toLowerCase() === lower) {
			return type;
		}
	}
	return undefined;
}

/** Resolves a host enum constant such as `xlUp` or `xlCalculationAutomatic`. */
export function resolveHostConstant(
	name: string,
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
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
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
): string | undefined {
	const lower = member.toLowerCase();
	const rawMember = model.types[qualified]?.members.find(
		(candidate) => candidate.name.toLowerCase() === lower,
	);
	if (rawMember?.kind === 'event') {
		return undefined;
	}
	return (
		getHostMembers(qualified, model).find((candidate) =>
			candidate.name.toLowerCase() === lower
		)?.signature ?? model.memberSignatures?.[qualified]?.[lower]
	);
}

/** A host-injected global identifier and the qualified type it denotes. */
export interface HostGlobal {
	name: string;
	type: string;
}

/** Returns all host-injected globals (canonical casing) of the model. */
export function getHostGlobals(
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
): HostGlobal[] {
	return Object.entries(model.globals).map(([name, type]) => ({ name, type }));
}

/** Returns all host enum constants (canonical casing) of the model. */
export function getHostConstants(
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
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
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
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
	for (const key of Object.keys(model.types)) {
		if (key.toLowerCase() === lower) {
			return key;
		}
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
	model: HostObjectModel = EXCEL_OBJECT_MODEL,
): string | undefined {
	const lower = memberName.toLowerCase();
	const member = getHostMembers(qualified, model).find(
		(mem) => mem.name.toLowerCase() === lower,
	);
	return member?.returns;
}
