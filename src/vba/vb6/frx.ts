// The `.frx` sidecar: the bytes a form's header points into with
// `"Form1.frx":0000`. There is no specification; these are the record shapes
// measured on files Visual Basic 6 wrote (tests/fixtures/vb6 and the wider
// corpus the census in the roadmap ran over):
//
// - Records are contiguous and referenced in header order, so the span of a
//   record is the distance to the next referenced offset (or the end of the
//   file). Nothing here needs a record's length prefix to find its end.
// - A `$"..."` reference (a long string: a Caption, a Text) is a 32-bit
//   little-endian length followed by that many bytes in the form's code page.
// - A short string (a TextBox's `Text` without the `$`) is an 8-bit length
//   followed by the bytes.
// - A picture-valued property (Picture, Icon, Image, MouseIcon, ToolboxBitmap,
//   ...) is a 32-bit length followed by an OLE-persisted picture; an empty
//   picture is eight bytes, `lt` and zeros.
// - `List` and `ItemData` (a ListBox's or ComboBox's rows) are a 16-bit count,
//   then, when the count is not zero, one more 16-bit word, then per row a
//   16-bit length and the bytes. An empty list is the two zero bytes alone.
//
// Anything else is carried as opaque bytes, which is all a byte-identical
// round trip needs.

import type { FrmHeader } from './frmHeader';
import { frmFrxRefs } from './frmHeader';

export type FrxValue =
	| { kind: 'longString'; text: string }
	| { kind: 'shortString'; text: string }
	| { kind: 'list'; items: string[] }
	| { kind: 'picture'; bytes: Buffer }
	| { kind: 'opaque'; bytes: Buffer };

export interface FrxRecord {
	/** The control that owns the property. */
	control: string;
	/** The property, with its group path when it sits in one: `Font.Name`. */
	property: string;
	offset: number;
	/** The record's bytes, exactly. */
	bytes: Buffer;
	value: FrxValue;
}

const LIST_PROPERTIES = new Set(['list', 'itemdata']);

/** The ranges the referenced offsets cut the sidecar into: offset -> end. */
export function frxSpans(blob: Buffer, offsets: readonly number[]): Map<number, number> {
	const sorted = [...new Set(offsets)].sort((a, b) => a - b);
	const spans = new Map<number, number>();
	for (let i = 0; i < sorted.length; i++) {
		const end = i + 1 < sorted.length ? sorted[i + 1] : blob.length;
		spans.set(sorted[i], Math.min(end, blob.length));
	}
	return spans;
}

/**
 * A 32-bit-length string, when the bytes are exactly one - or, with
 * `allowTrailing`, when they start with one and what follows is not the
 * record's (a record nothing references any more, left by an editor that
 * appends and never rewrites).
 */
export function decodeFrxLongString(bytes: Buffer, decode: (b: Buffer) => string, allowTrailing = false): string | undefined {
	if (bytes.length < 4) {
		return undefined;
	}
	const length = bytes.readUInt32LE(0);
	const fits = allowTrailing ? length + 4 <= bytes.length : length + 4 === bytes.length;
	return fits ? decode(bytes.subarray(4, 4 + length)) : undefined;
}

/** An 8-bit-length string, when the bytes are exactly one (or start with one, with `allowTrailing`). */
export function decodeFrxShortString(bytes: Buffer, decode: (b: Buffer) => string, allowTrailing = false): string | undefined {
	if (bytes.length < 1) {
		return undefined;
	}
	const fits = allowTrailing ? bytes[0] + 1 <= bytes.length : bytes[0] + 1 === bytes.length;
	return fits ? decode(bytes.subarray(1, 1 + bytes[0])) : undefined;
}

/** A `List`/`ItemData` record's rows, when the bytes parse as one exactly (or start with one, with `allowTrailing`). */
export function decodeFrxList(bytes: Buffer, decode: (b: Buffer) => string, allowTrailing = false): string[] | undefined {
	if (bytes.length < 2) {
		return undefined;
	}
	const count = bytes.readUInt16LE(0);
	if (count === 0) {
		return allowTrailing || bytes.length === 2 ? [] : undefined;
	}
	let at = 4;
	const items: string[] = [];
	for (let i = 0; i < count; i++) {
		if (at + 2 > bytes.length) {
			return undefined;
		}
		const length = bytes.readUInt16LE(at);
		at += 2;
		if (at + length > bytes.length) {
			return undefined;
		}
		items.push(decode(bytes.subarray(at, at + length)));
		at += length;
	}
	return allowTrailing || at === bytes.length ? items : undefined;
}

/** A 32-bit-length picture record, when the bytes are exactly one (or start with one, with `allowTrailing`). */
export function decodeFrxPicture(bytes: Buffer, allowTrailing = false): Buffer | undefined {
	if (bytes.length < 4) {
		return undefined;
	}
	const length = bytes.readUInt32LE(0);
	const fits = allowTrailing ? length + 4 <= bytes.length : length + 4 === bytes.length;
	return fits ? bytes.subarray(4, 4 + length) : undefined;
}

/** The intrinsic properties that keep a picture in the sidecar. */
const PICTURE_PROPERTIES = new Set([
	'picture', 'icon', 'mouseicon', 'dragicon', 'downpicture', 'disabledpicture', 'maskpicture', 'toolboxbitmap', 'tabpicture',
]);
/** The intrinsic properties that keep a string in the sidecar when it is long or has line breaks. */
const STRING_PROPERTIES = new Set([
	'caption', 'text', 'tooltiptext', 'tag', 'datafield', 'datamember', 'linktopic', 'linkitem', 'helpfile', 'title', 'formatstring',
]);

/**
 * Every record the header references, decoded where its shape is one of the
 * measured ones and carried opaque otherwise. `decode` turns the form's code
 * page bytes into text.
 */
export function readFrxRecords(header: FrmHeader, blob: Buffer, decode: (b: Buffer) => string): FrxRecord[] {
	const refs = frmFrxRefs(header);
	const spans = frxSpans(blob, refs.map((r) => r.property.frx!.offset));
	return refs.map(({ control, property, group }) => {
		const offset = property.frx!.offset;
		const bytes = blob.subarray(offset, spans.get(offset) ?? blob.length);
		const name = group ? `${group}.${property.key}` : property.key;
		return {
			control: control.name,
			property: name,
			offset,
			bytes,
			value: decodeFrxValue(property.key, property.frx!.long, bytes, decode),
		};
	});
}

function decodeFrxValue(key: string, long: boolean, bytes: Buffer, decode: (b: Buffer) => string): FrxValue {
	if (long) {
		const text = decodeFrxLongString(bytes, decode);
		if (text !== undefined) {
			return { kind: 'longString', text };
		}
	}
	if (LIST_PROPERTIES.has(key.replace(/\(.*$/, '').toLowerCase())) {
		const items = decodeFrxList(bytes, decode);
		if (items !== undefined) {
			return { kind: 'list', items };
		}
	}
	const picture = decodeFrxPicture(bytes);
	if (picture !== undefined) {
		return { kind: 'picture', bytes: picture };
	}
	const short = decodeFrxShortString(bytes, decode);
	if (short !== undefined) {
		return { kind: 'shortString', text: short };
	}
	// The bytes run to the next referenced offset, so a record that nothing
	// references any more (an editor that appends and never rewrites leaves
	// them) sits inside the span of the record before it. A second pass
	// reads the record by the length it declares, for the kind the
	// reference and the property imply; an OCX's own record stays opaque,
	// because nothing says what it is.
	const base = key.replace(/\(.*$/, '').toLowerCase();
	if (long) {
		const text = decodeFrxLongString(bytes, decode, true);
		if (text !== undefined) {
			return { kind: 'longString', text };
		}
	} else if (LIST_PROPERTIES.has(base)) {
		const items = decodeFrxList(bytes, decode, true);
		if (items !== undefined) {
			return { kind: 'list', items };
		}
	} else if (PICTURE_PROPERTIES.has(base)) {
		const prefix = decodeFrxPicture(bytes, true);
		if (prefix !== undefined) {
			return { kind: 'picture', bytes: prefix };
		}
	} else if (STRING_PROPERTIES.has(base)) {
		const prefix = decodeFrxShortString(bytes, decode, true);
		if (prefix !== undefined) {
			return { kind: 'shortString', text: prefix };
		}
	}
	return { kind: 'opaque', bytes };
}
