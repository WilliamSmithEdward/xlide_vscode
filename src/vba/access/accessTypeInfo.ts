import { decodeCodePage, encodeCodePage } from '../codePages';
import { AccessFormatError } from './accessFormat';
import {
	accessDesignObjectName,
	accessVbaIdentifier,
	isAccessDesignSection,
	type AccessDesign,
	type AccessDesignObject,
} from './accessDesign';
import {
	ACTIVEX_CONTROL,
	CONTROL_TYPES,
	TYPE_INFO_ACTIVEX_TAIL_LENGTH,
	TYPE_INFO_CLSID_AT,
	TYPE_INFO_ENTRIES_AT,
	TYPE_INFO_HELD_IDS,
	TYPE_INFO_IDS,
	TYPE_INFO_KIND,
	TYPE_INFO_MAGIC,
	TYPE_INFO_REPORT_IDS,
	TYPE_INFO_SECTIONS,
} from './accessDesignTable';
import { designObjectHolders, type AccessDesignKind } from './accessDesignEdit';

/**
 * The `TypeInfo` stream beside a design: the members VBA sees on the form's
 * class. `Me.Qty` compiles, and `Qty_Click` binds, only for a name listed
 * here - so a control added to the design and not to this stream is one the
 * code behind the form cannot reach.
 *
 * Access keeps the stream rather than rebuilding it, so no rule of the design
 * alone reproduces one Access wrote. The edits are reproduced instead: a new
 * member is appended with the ordinal above the highest present, a removed one
 * drops out while the others keep their ordinals, and a renamed one moves to
 * the end with its ordinal. A freed ordinal is never reused.
 *
 * A member has two names, because a control is rarely named as VBA would name
 * it: the form wizard names one after its field, so `Order Date` is ordinary,
 * and VBA reaches it as `Me.Order_Date`. An entry stores the identifier, a
 * NUL, the design's name, a NUL - and leaves the design's name empty where
 * the two are the same, which is why `Plain` reads as `Plain\0\0`:
 *
 *     6d 12 00 00  01 00 00 00  "Order_Date" 00 "Order Date" 00
 *     6d 12 00 00  05 00 00 00  "Plain" 00 00
 *
 * Measured on Access 16.0, for controls, sections and an ActiveX control,
 * whose tail follows both names.
 *
 * Ported from pyOpenVBA's `_designs.py`.
 */

/** One entry of the stream, as it is stored. */
export interface AccessTypeInfoEntry {
	/** The member's type id: its class index over its type code. */
	ident: number;
	ordinal: number;
	/** The name the design gives the object, which the designer shows. */
	name: string;
	/** The name VBA compiles against: `name` itself where that is an identifier. */
	identifier: string;
	/** An ActiveX control carries 36 more bytes after its names. */
	tail: Buffer;
}

const PAGE = 'Page';

/** The members a `TypeInfo` stream lists, in the order it lists them. */
export function readTypeInfo(stream: Buffer, codePage: number): AccessTypeInfoEntry[] {
	if (!stream.subarray(0, 4).equals(TYPE_INFO_MAGIC) || stream.length < TYPE_INFO_ENTRIES_AT) {
		throw new AccessFormatError('This is not a TypeInfo stream.');
	}
	const count = stream.readUInt32LE(12);
	const out: AccessTypeInfoEntry[] = [];
	let at = TYPE_INFO_ENTRIES_AT;
	for (let i = 0; i < count; i += 1) {
		if (at + 8 > stream.length) {
			throw new AccessFormatError('This TypeInfo stream ends inside an entry.');
		}
		const ident = stream.readUInt32LE(at);
		const ordinal = stream.readUInt32LE(at + 4);
		const identifierEnd = stream.indexOf(0, at + 8);
		const nameEnd = identifierEnd < 0 ? -1 : stream.indexOf(0, identifierEnd + 1);
		if (nameEnd < 0) {
			throw new AccessFormatError('This TypeInfo stream ends inside a name.');
		}
		const identifier = decodeCodePage(stream.subarray(at + 8, identifierEnd), codePage);
		const shown = decodeCodePage(stream.subarray(identifierEnd + 1, nameEnd), codePage);
		at = nameEnd + 1;
		let tail = Buffer.alloc(0);
		if ((ident & 0xff) === CODE_OF_ACTIVEX) {
			tail = Buffer.from(stream.subarray(at, at + TYPE_INFO_ACTIVEX_TAIL_LENGTH));
			at += tail.length;
		}
		out.push({ ident, ordinal, name: shown || identifier, identifier, tail });
	}
	return out;
}

const CODE_OF_ACTIVEX = [...CONTROL_TYPES]
	.find(([, name]) => name === ACTIVEX_CONTROL)?.[0] ?? -1;

/** A stream listing the entries in that order. */
export function buildTypeInfo(
	kind: AccessDesignKind,
	clsid: Buffer,
	entries: readonly AccessTypeInfoEntry[],
	codePage: number,
): Buffer {
	const kindWord = TYPE_INFO_KIND.get(kind);
	if (kindWord === undefined) {
		throw new AccessFormatError(`Kind must be form or report, not ${kind}.`);
	}
	if (clsid.length !== 16) {
		throw new AccessFormatError('A TypeInfo CLSID is 16 bytes.');
	}
	const head = Buffer.alloc(TYPE_INFO_ENTRIES_AT);
	TYPE_INFO_MAGIC.copy(head, 0);
	head.writeUInt32LE(kindWord, 4);
	head.writeInt32LE(-1, 8);
	head.writeUInt32LE(entries.length, 12);
	clsid.copy(head, TYPE_INFO_CLSID_AT);
	const parts: Buffer[] = [head];
	const nul = Buffer.alloc(1);
	for (const entry of entries) {
		const framed = Buffer.alloc(8);
		framed.writeUInt32LE(entry.ident, 0);
		framed.writeUInt32LE(entry.ordinal, 4);
		parts.push(
			framed,
			encodeCodePage(entry.identifier, codePage), nul,
			// The design's name is stored only where VBA knows the member by another.
			entry.name === entry.identifier ? Buffer.alloc(0) : encodeCodePage(entry.name, codePage), nul,
			entry.tail,
		);
	}
	return Buffer.concat(parts);
}

/** One member the design has, before its ordinal is decided. */
interface Member {
	name: string;
	identifier: string;
	ident: number;
	tail: Buffer;
}

/**
 * The members a design has - its sections, then its named controls, each in
 * design order - with the type id Access gives each. A type whose id has not
 * been measured on that kind of design is refused rather than guessed.
 */
function membersOf(kind: AccessDesignKind, objects: AccessDesignObject[]): Member[] {
	const ids = TYPE_INFO_SECTIONS.get(kind);
	if (!ids) {
		throw new AccessFormatError(`Kind must be form or report, not ${kind}.`);
	}
	const holders = designObjectHolders(objects);
	const sections: Member[] = [];
	const controls: Member[] = [];
	objects.forEach((object, index) => {
		const name = accessDesignObjectName(object);
		if (index === 0 || !name) {
			return;
		}
		const typeName = object.type === undefined ? '' : (CONTROL_TYPES.get(object.type) ?? '');
		let ident: number | undefined;
		if (isAccessDesignSection(object)) {
			ident = ids.get(typeName);
		} else if (kind === 'report') {
			ident = TYPE_INFO_REPORT_IDS.get(typeName);
		} else {
			const holder = holders.get(index);
			const held = holder !== undefined && !isAccessDesignSection(objects[holder])
				&& (objects[holder].type === undefined
					|| CONTROL_TYPES.get(objects[holder].type!) !== PAGE);
			ident = (held ? TYPE_INFO_HELD_IDS.get(typeName) : undefined)
				?? TYPE_INFO_IDS.get(typeName);
		}
		if (ident === undefined) {
			throw new AccessFormatError(
				`No TypeInfo id has been measured for a ${typeName || object.type} on a ${kind}.`,
			);
		}
		const tail = typeName === ACTIVEX_CONTROL
			? Buffer.alloc(TYPE_INFO_ACTIVEX_TAIL_LENGTH)
			: Buffer.alloc(0);
		(isAccessDesignSection(object) ? sections : controls)
			.push({ name, identifier: accessVbaIdentifier(name), ident, tail });
	});
	return [...sections, ...controls];
}

/**
 * An entry as it is carried forward. The identifier Access stored is kept as
 * it is. One that is not an identifier at all is not Access's: this module
 * once read an entry as a single name, and rewrote a member named `Order
 * Date` as exactly that, which `Me.` cannot reach. That one is put right from
 * the design's name.
 */
function carried(entry: AccessTypeInfoEntry): AccessTypeInfoEntry {
	return accessVbaIdentifier(entry.identifier) === entry.identifier
		? entry
		: { ...entry, identifier: accessVbaIdentifier(entry.name) };
}

/**
 * VBA knows a member by its identifier alone, so two members sharing one are
 * a single name to the code behind the design. Access refuses the second
 * control name ("already in use"), and so does this.
 */
function refuseSharedIdentifiers(entries: readonly AccessTypeInfoEntry[]): void {
	const seen = new Map<string, string>();
	for (const entry of entries) {
		const key = entry.identifier.toLowerCase();
		const other = seen.get(key);
		if (other !== undefined) {
			throw new AccessFormatError(
				`${entry.name} and ${other} would both be ${entry.identifier} to VBA, which knows a `
				+ 'member by that name alone. Access refuses the second name as already in use.',
			);
		}
		seen.set(key, entry.name);
	}
}

/**
 * The stream after the design changed, carried forward the way Access carries
 * it. `renamed` maps a member's old name in the design to its new one.
 */
export function updateTypeInfo(
	kind: AccessDesignKind,
	design: AccessDesign,
	existing: Buffer,
	codePage: number,
	renamed: ReadonlyMap<string, string> = new Map(),
): Buffer {
	const entries = readTypeInfo(existing, codePage);
	const members = membersOf(kind, design.objects);
	const wanted = new Set(members.map((member) => member.name));
	const kept = entries
		.filter((entry) => wanted.has(entry.name) && !renamed.has(entry.name))
		.map(carried);
	const moved = entries
		.filter((entry) => renamed.has(entry.name) && wanted.has(renamed.get(entry.name)!))
		.map((entry) => {
			const name = renamed.get(entry.name)!;
			return { ...entry, name, identifier: accessVbaIdentifier(name) };
		});
	const present = new Set([...kept, ...moved].map((entry) => entry.name));
	let ordinal = [...kept, ...moved]
		.reduce((most, entry) => Math.max(most, entry.ordinal), -1) + 1;
	const added: AccessTypeInfoEntry[] = [];
	for (const member of members) {
		if (present.has(member.name)) {
			continue;
		}
		added.push({
			ident: member.ident, ordinal, name: member.name, identifier: member.identifier, tail: member.tail,
		});
		ordinal += 1;
	}
	const next = [...kept, ...moved, ...added];
	refuseSharedIdentifiers(next);
	const clsid = existing.subarray(TYPE_INFO_CLSID_AT, TYPE_INFO_CLSID_AT + 16);
	return buildTypeInfo(kind, clsid, next, codePage);
}
