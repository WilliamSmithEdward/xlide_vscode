// Descriptions for host members the Office VBA reference never published.
//
// The reference covers most of each library but not all of it: Excel publishes
// prose for 4,938 of its 13,187 members, Access for 3,136 of 6,132. The rest
// arrive from the type library with a name, a kind, a declared type and a
// read/write contract, and nothing to show in a tooltip.
//
// This composes a description from exactly those four facts and marks it
// `derived`, which hover renders as a provenance note, so a reader can tell a
// transcribed sentence from a composed one. Nothing is inferred about what the
// member MEANS - only what the type library declares about it.

import type { HostConstant, HostEnum, HostMember } from './excelObjectModel';
import type { VbaDoc } from '../docs/docModel';

/** Bare type name from a qualified one: "Excel.Range" -> "Range". */
function bareTypeName(qualified: string): string {
	const dot = qualified.lastIndexOf('.');
	return dot >= 0 ? qualified.slice(dot + 1) : qualified;
}

/**
 * The type a member declares: stated outright for a property, read from a
 * method's signature tail, and falling back to the chaining return. Undefined
 * when none of the three states one, which is the honest answer for a Sub.
 */
function declaredType(member: HostMember): string | undefined {
	if (member.declaredType) {
		return bareTypeName(member.declaredType);
	}
	const tail = member.signature?.match(/\bAs ([A-Za-z_][A-Za-z0-9_.]*)$/)?.[1];
	if (tail) {
		return bareTypeName(tail);
	}
	return member.returns ? bareTypeName(member.returns) : undefined;
}

const ACCESS_PREFIX: Record<string, string> = {
	'read-only': 'Read-only',
	'read/write': 'Read/write',
	'write-only': 'Write-only',
};

/**
 * A description composed from the type library's own declaration, for a member
 * the reference does not document. Undefined when the member has too little to
 * say - an unqualified method on an unnamed owner adds nothing a reader cannot
 * already see in the signature.
 */
export function derivedMemberDoc(
	member: HostMember,
	ownerDisplayName: string,
): VbaDoc | undefined {
	if (member.doc?.summary) {
		return undefined;
	}
	const type = declaredType(member);
	const owner = bareTypeName(ownerDisplayName);
	if (!owner) {
		return undefined;
	}
	const summary = member.kind === 'property'
		? propertySummary(member, type, owner)
		: methodSummary(type, owner);
	return {
		summary,
		params: member.doc?.params ?? [],
		source: 'derived',
	};
}

function propertySummary(
	member: HostMember,
	type: string | undefined,
	owner: string,
): string {
	const access = member.access ? ACCESS_PREFIX[member.access] : undefined;
	const lead = access ? `${access} ` : '';
	const typed = type ? `${type} ` : '';
	return `${lead}${typed}property of the ${owner} object.`;
}

function methodSummary(type: string | undefined, owner: string): string {
	const returns = type && type !== 'void' ? ` Returns ${type}.` : '';
	return `Method of the ${owner} object.${returns}`;
}

/**
 * A description for an enum constant the reference does not document, built from
 * the enumeration it belongs to. The reference describes 15,570 of the 19,816
 * constants these libraries declare; for the rest, the enumeration's own
 * description is the nearest true thing there is to say - `xlPrimary` gets
 * "Member of XlAxisGroup: Specifies the axis group."
 *
 * Undefined when the constant is already documented, so a transcribed sentence
 * is never replaced.
 */
export function derivedConstantDoc(
	constant: HostConstant,
	owner: HostEnum | undefined,
): VbaDoc | undefined {
	if (constant.doc?.summary) {
		return undefined;
	}
	const enumName = owner?.displayName ?? constant.type;
	if (!enumName) {
		return undefined;
	}
	const about = owner?.doc?.summary;
	return {
		summary: about
			? `Member of ${enumName}: ${about}`
			: `Member of the ${enumName} enumeration.`,
		params: [],
		source: 'derived',
	};
}
