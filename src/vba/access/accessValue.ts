import { AccessColumnType, AccessFormatError, type AccessColumn } from './accessFormat';

/**
 * Encoding a value into the bytes a row stores for its column.
 *
 * The reader in `accessFormat.ts` decodes; this is its inverse, for the column
 * types the catalog and storage tables use. A type this does not cover belongs
 * to a table XLIDE does not write, and is refused rather than guessed.
 *
 * Ported from pyOpenVBA's `_rows.encode_scalar`.
 */

/** Text is UTF-16LE, or the compressed form that opens `FF FE`. */
const TEXT_COMPRESSION_MARK = Buffer.from([0xff, 0xfe]);
/** Byte 16 bit 0: the column asks for Unicode compression. */
const COLUMN_COMPRESSED_UNICODE = 0x01;
/** Days between the OLE automation epoch and the Unix one. */
const OLE_EPOCH_OFFSET = 25569;
const MS_PER_DAY = 86400000;

export type AccessScalar = number | string | boolean | Buffer | Date | null;

/** Compress when every character fits one byte, as Access does. */
export function encodeAccessText(text: string): Buffer {
	if (text.length > 0 && [...text].every((ch) => {
		const code = ch.codePointAt(0)!;
		return code >= 1 && code <= 0xff;
	})) {
		return Buffer.concat([TEXT_COMPRESSION_MARK, Buffer.from(text, 'latin1')]);
	}
	return Buffer.from(text, 'utf16le');
}

/** The OLE automation serial a Date stores as. */
export function accessDateSerial(when: Date): number {
	return when.getTime() / MS_PER_DAY + OLE_EPOCH_OFFSET;
}

/**
 * One column's stored bytes. A Boolean yields none: it lives in the null mask,
 * so the caller passes its column number in the Boolean set instead.
 */
export function encodeAccessScalar(column: AccessColumn, value: AccessScalar): Buffer {
	const name = column.name;
	switch (column.type) {
		case AccessColumnType.Boolean:
			return Buffer.alloc(0);
		case AccessColumnType.Byte: {
			assertInteger(name, value, 0, 255);
			return Buffer.from([value as number]);
		}
		case AccessColumnType.Integer: {
			assertInteger(name, value, -32768, 32767);
			const out = Buffer.alloc(2);
			out.writeInt16LE(value as number, 0);
			return out;
		}
		case AccessColumnType.Long:
		case AccessColumnType.Complex: {
			assertInteger(name, value, -2147483648, 2147483647);
			const out = Buffer.alloc(4);
			out.writeInt32LE(value as number, 0);
			return out;
		}
		case AccessColumnType.BigInt: {
			if (typeof value !== 'number' && typeof value !== 'bigint') {
				throw new AccessFormatError(`Column ${name}: ${String(value)} is not a BigInt.`);
			}
			const out = Buffer.alloc(8);
			out.writeBigInt64LE(BigInt(value), 0);
			return out;
		}
		case AccessColumnType.Single: {
			assertNumber(name, value, 'a Single');
			const out = Buffer.alloc(4);
			out.writeFloatLE(value as number, 0);
			return out;
		}
		case AccessColumnType.Double: {
			assertNumber(name, value, 'a Double');
			const out = Buffer.alloc(8);
			out.writeDoubleLE(value as number, 0);
			return out;
		}
		case AccessColumnType.DateTime: {
			// A number is taken as the stored serial itself, which is how a
			// stamp read from another database is reproduced bit for bit.
			const serial = value instanceof Date ? accessDateSerial(value) : value;
			assertNumber(name, serial, 'a date');
			const out = Buffer.alloc(8);
			out.writeDoubleLE(serial as number, 0);
			return out;
		}
		case AccessColumnType.Binary: {
			if (!Buffer.isBuffer(value)) {
				throw new AccessFormatError(`Column ${name}: the value is not bytes.`);
			}
			if (value.length > column.length) {
				throw new AccessFormatError(
					`Column ${name}: ${value.length} bytes exceed its size ${column.length}.`,
				);
			}
			// A fixed-size Binary column always holds its full width; the
			// engine pads with zeros and hands the padded value back.
			if (!column.fixed) {
				return Buffer.from(value);
			}
			const out = Buffer.alloc(column.length);
			value.copy(out);
			return out;
		}
		case AccessColumnType.Text: {
			if (typeof value !== 'string') {
				throw new AccessFormatError(`Column ${name}: the value is not text.`);
			}
			if (2 * value.length > column.length) {
				throw new AccessFormatError(
					`Column ${name}: ${value.length} characters exceed its size ${column.length / 2}.`,
				);
			}
			const plain = Buffer.from(value, 'utf16le');
			if ((column.miscFlags & COLUMN_COMPRESSED_UNICODE) === 0) {
				return plain;
			}
			// The engine compresses a Text value only when that makes it
			// shorter: one and two Latin-1 characters go in as plain UTF-16,
			// since the FF FE mark would cost what the compression saves.
			const compressed = encodeAccessText(value);
			return compressed.length < plain.length ? compressed : plain;
		}
		case AccessColumnType.Ole:
		case AccessColumnType.Memo: {
			if (!Buffer.isBuffer(value)) {
				throw new AccessFormatError(`Column ${name}: a long value must be passed encoded.`);
			}
			return Buffer.from(value);
		}
		default:
			throw new AccessFormatError(`Column ${name}: type ${column.type} cannot be encoded.`);
	}
}

function assertInteger(name: string, value: unknown, low: number, high: number): void {
	if (typeof value !== 'number' || !Number.isInteger(value) || value < low || value > high) {
		throw new AccessFormatError(
			`Column ${name}: ${String(value)} is not an integer between ${low} and ${high}.`,
		);
	}
}

function assertNumber(name: string, value: unknown, what: string): void {
	// Infinities and NaN store as their bit patterns, which is what the engine
	// does with them, so only a non-number is refused here.
	if (typeof value !== 'number') {
		throw new AccessFormatError(`Column ${name}: ${String(value)} is not ${what}.`);
	}
}
