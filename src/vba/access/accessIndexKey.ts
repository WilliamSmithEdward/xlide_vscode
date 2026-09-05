import {
	AccessColumnType,
	AccessFormatError,
	decodeAccessText,
	type AccessColumn,
	type AccessRealIndex,
	type AccessTableDefinition,
} from './accessFormat';
import { encodeTextKey } from './accessCollation';
import type { AccessRawRow } from './accessRow';

/**
 * The bytes the engine stores for a row's key columns in an index entry.
 *
 * Each key column starts with a flag byte: 0x7F for a value in an ascending
 * column, 0x80 descending, 0x00 for null ascending, 0xFF null descending. A
 * descending value is the ascending encoding with every byte inverted.
 * Ascending encodings are all big-endian:
 *
 * - Boolean: one byte, 0x00 for True and 0xFF for False, so True sorts first.
 * - Byte: the byte.
 * - Integer, Long, BigInt, Currency: two's complement with the sign bit flipped.
 * - Single, Double, DateTime: the IEEE bits, sign bit flipped when positive and
 *   every bit inverted when negative.
 * - Decimal: 0xFF then the 16-byte magnitude for a positive value.
 * - GUID and Binary: eight-byte chunks, each followed by 0x09 when another
 *   follows and by the count of real bytes in the last.
 * - Text: through the collation table.
 *
 * The key is built from the bytes the row stores rather than from decoded
 * values, so a Currency, a Decimal or a date never passes through a JavaScript
 * number on the way. Ported from pyOpenVBA's `_index.py`.
 */

const FLAG_ASCENDING = 0x7f;
const FLAG_DESCENDING = 0x80;
const FLAG_NULL_ASCENDING = 0x00;
const FLAG_NULL_DESCENDING = 0xff;
const BINARY_CHUNK = 8;
const BINARY_MORE = 0x09;

/** Width of the fixed-size key encodings, by column type. */
const FIXED_KEY_SIZES: Partial<Record<AccessColumnType, number>> = {
	[AccessColumnType.Boolean]: 1,
	[AccessColumnType.Byte]: 1,
	[AccessColumnType.Integer]: 2,
	[AccessColumnType.Long]: 4,
	[AccessColumnType.Complex]: 4,
	[AccessColumnType.Money]: 8,
	[AccessColumnType.Single]: 4,
	[AccessColumnType.Double]: 8,
	[AccessColumnType.DateTime]: 8,
	[AccessColumnType.BigInt]: 8,
	[AccessColumnType.Numeric]: 17,
};

function invert(raw: Buffer): Buffer {
	const out = Buffer.alloc(raw.length);
	for (let i = 0; i < raw.length; i += 1) {
		out[i] = raw[i] ^ 0xff;
	}
	return out;
}

/** Little-endian stored bytes, zero-extended and turned big-endian. */
function bigEndian(raw: Buffer, size: number): Buffer {
	const wide = raw.length >= size
		? raw.subarray(0, size)
		: Buffer.concat([raw, Buffer.alloc(size - raw.length)]);
	return Buffer.from(wide).reverse();
}

/** Sign bit set means positive with the bit flipped; clear means all inverted. */
function floatKey(bigEndianBits: Buffer): Buffer {
	if ((bigEndianBits[0] & 0x80) !== 0) {
		return invert(bigEndianBits);
	}
	const out = Buffer.from(bigEndianBits);
	out[0] |= 0x80;
	return out;
}

function binaryKey(data: Buffer): Buffer {
	if (data.length === 0) {
		throw new AccessFormatError('An empty binary value has no key.');
	}
	const parts: Buffer[] = [];
	for (let start = 0; start < data.length; start += BINARY_CHUNK) {
		const chunk = data.subarray(start, start + BINARY_CHUNK);
		const padded = Buffer.alloc(BINARY_CHUNK);
		chunk.copy(padded);
		const marker = start + BINARY_CHUNK < data.length ? BINARY_MORE : chunk.length;
		parts.push(padded, Buffer.from([marker]));
	}
	return Buffer.concat(parts);
}

/** One column's key bytes, from the bytes the row stores for it. */
function columnKey(column: AccessColumn, stored: Buffer): Buffer {
	const size = FIXED_KEY_SIZES[column.type];
	switch (column.type) {
		case AccessColumnType.Boolean:
			// A Boolean lives only in the null mask, and a set bit means True.
			return Buffer.from([stored.length > 0 ? 0x00 : 0xff]);
		case AccessColumnType.Byte:
			return Buffer.from([stored[0] ?? 0]);
		case AccessColumnType.Integer:
		case AccessColumnType.Long:
		case AccessColumnType.Complex:
		case AccessColumnType.Money:
		case AccessColumnType.BigInt: {
			const out = bigEndian(stored, size!);
			out[0] ^= 0x80;
			return out;
		}
		case AccessColumnType.Single:
		case AccessColumnType.Double:
		case AccessColumnType.DateTime:
			return floatKey(bigEndian(stored, size!));
		case AccessColumnType.Numeric: {
			// Stored as a sign byte then four little-endian words, most
			// significant first; the key is 0xFF and the magnitude big-endian.
			const magnitude = Buffer.alloc(16);
			for (let word = 0; word < 4; word += 1) {
				for (let byte = 0; byte < 4; byte += 1) {
					magnitude[word * 4 + byte] = stored[1 + word * 4 + (3 - byte)] ?? 0;
				}
			}
			return stored[0] === 0
				? Buffer.concat([Buffer.from([0xff]), magnitude])
				: Buffer.concat([Buffer.from([0x00]), invert(magnitude)]);
		}
		case AccessColumnType.Guid:
			return binaryKey(stored.subarray(0, 16));
		case AccessColumnType.Binary:
			return binaryKey(stored);
		case AccessColumnType.Text:
			return encodeTextKey(decodeAccessText(stored));
		default:
			throw new AccessFormatError(
				`Column ${column.name}: type ${column.type} cannot be part of an index key.`,
			);
	}
}

/**
 * The key an index holds for a row, or undefined when the index ignores nulls
 * and every key column of the row is null - which is a row it holds no entry
 * for at all.
 */
export function encodeIndexKey(
	definition: AccessTableDefinition,
	index: AccessRealIndex,
	row: AccessRawRow,
): Buffer | undefined {
	const parts: Buffer[] = [];
	let anyPresent = false;
	for (const key of index.columns) {
		const column = definition.columns.find((c) => c.number === key.number);
		if (!column) {
			throw new AccessFormatError(`The index names column ${key.number}, which the table has not.`);
		}
		const present = row.present.get(column.number) === true;
		if (column.type === AccessColumnType.Boolean) {
			// A Boolean is never null: the mask bit is the value.
			anyPresent = true;
			parts.push(Buffer.from([key.ascending ? FLAG_ASCENDING : FLAG_DESCENDING]));
			const encoded = columnKey(column, present ? Buffer.alloc(1) : Buffer.alloc(0));
			parts.push(key.ascending ? encoded : invert(encoded));
			continue;
		}
		if (!present) {
			parts.push(Buffer.from([key.ascending ? FLAG_NULL_ASCENDING : FLAG_NULL_DESCENDING]));
			continue;
		}
		anyPresent = true;
		parts.push(Buffer.from([key.ascending ? FLAG_ASCENDING : FLAG_DESCENDING]));
		const encoded = columnKey(column, row.values.get(column.number) ?? Buffer.alloc(0));
		parts.push(key.ascending ? encoded : invert(encoded));
	}
	if (!anyPresent && (index.flags & INDEX_IGNORE_NULLS) !== 0) {
		return undefined;
	}
	return Buffer.concat(parts);
}

/** An index with this flag holds no entry for a row whose key is all null. */
export const INDEX_IGNORE_NULLS = 0x02;
export const INDEX_UNIQUE = 0x01;
