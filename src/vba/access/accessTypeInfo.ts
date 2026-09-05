import { decodeCodePage, encodeCodePage } from '../codePages';
import { AccessFormatError } from './accessFormat';
import {
	accessDesignObjectName,
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
 * Ported from pyOpenVBA's `_designs.py`.
 */

/** One entry of the stream, as it is stored. */
export interface AccessTypeInfoEntry {
	/** The member's type id: its class index over its type code. */
	ident: number;
	ordinal: number;
	name: string;
	/** An ActiveX control carries 36 more bytes after the name. */
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
		const end = stream.indexOf(Buffer.alloc(2), at + 8);
		if (end < 0) {
			throw new AccessFormatError('This TypeInfo stream ends inside a name.');
		}
		const name = decodeCodePage(stream.subarray(at + 8, end), codePage);
		at = end + 2;
		let tail = Buffer.alloc(0);
		if ((ident & 0xff) === CODE_OF_ACTIVEX) {
			tail = Buffer.from(stream.subarray(at, at + TYPE_INFO_ACTIVEX_TAIL_LENGTH));
			at += tail.length;
		}
		out.push({ ident, ordinal, name, tail });
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
	for (const entry of entries) {
		const framed = Buffer.alloc(8);
		framed.writeUInt32LE(entry.ident, 0);
		framed.writeUInt32LE(entry.ordinal, 4);
		parts.push(framed, encodeCodePage(entry.name, codePage), Buffer.alloc(2), entry.tail);
	}
	return Buffer.concat(parts);
}

/** One member the design has, before its ordinal is decided. */
interface Member {
	name: string;
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
		(isAccessDesignSection(object) ? sections : controls).push({ name, ident, tail });
	});
	return [...sections, ...controls];
}

/**
 * The stream after the design changed, carried forward the way Access carries
 * it. `renamed` maps an old member name to its new one.
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
	const kept = entries.filter(
		(entry) => wanted.has(entry.name) && !renamed.has(entry.name),
	);
	const moved = entries
		.filter((entry) => renamed.has(entry.name) && wanted.has(renamed.get(entry.name)!))
		.map((entry) => ({ ...entry, name: renamed.get(entry.name)! }));
	const present = new Set([...kept, ...moved].map((entry) => entry.name));
	let ordinal = [...kept, ...moved]
		.reduce((most, entry) => Math.max(most, entry.ordinal), -1) + 1;
	const added: AccessTypeInfoEntry[] = [];
	for (const member of members) {
		if (present.has(member.name)) {
			continue;
		}
		added.push({ ident: member.ident, ordinal, name: member.name, tail: member.tail });
		ordinal += 1;
	}
	const clsid = existing.subarray(TYPE_INFO_CLSID_AT, TYPE_INFO_CLSID_AT + 16);
	return buildTypeInfo(kind, clsid, [...kept, ...moved, ...added], codePage);
}
