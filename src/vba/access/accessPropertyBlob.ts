/**
 * A Jet property blob (`MR2`), the `LvProp` a catalog row carries: MSysDb's
 * holds the database's own properties, `Theme Resource Name` among them.
 *
 *     "MR2\0"
 *     blocks: <u32 length, itself included> <u16 type> <body>
 *       type 0x80, the names:   (<u16 byte length> <name UTF-16>)...
 *       type 0x00 or 0x01, values: <u32 map-name length, itself included>
 *         <map name when longer than 6>, then per value
 *         <u16 length, itself included> <u8 flag> <u8 data type>
 *         <u16 name index> <u16 data size> <data>
 */

const SIGNATURE = 'MR2\0';
const NAME_LIST = 0x80;
const TYPE_TEXT = 0x0a;
const TYPE_MEMO = 0x0c;
const TYPE_BYTE = 0x02;
const TYPE_INTEGER = 0x03;
const TYPE_LONG = 0x04;
const TYPE_BOOLEAN = 0x01;

export type AccessPropertyValue = string | number | boolean | Buffer;

/**
 * Every value the blob holds, by property name. A name that several maps
 * carry keeps the first map's value, which for MSysDb is the database's own.
 * Undefined when the bytes are not a property blob.
 */
export function readAccessPropertyBlob(blob: Buffer): Map<string, AccessPropertyValue> | undefined {
	if (blob.length < 4 || blob.toString('latin1', 0, 4) !== SIGNATURE) {
		return undefined;
	}
	const names: string[] = [];
	const out = new Map<string, AccessPropertyValue>();
	let at = 4;
	while (at + 6 <= blob.length) {
		const length = blob.readUInt32LE(at);
		const type = blob.readUInt16LE(at + 4);
		const end = at + length;
		if (length < 6 || end > blob.length) {
			return out;
		}
		let cursor = at + 6;
		if (type === NAME_LIST) {
			while (cursor + 2 <= end) {
				const size = blob.readUInt16LE(cursor);
				names.push(blob.toString('utf16le', cursor + 2, cursor + 2 + size));
				cursor += 2 + size;
			}
		} else if (cursor + 4 <= end) {
			cursor += blob.readUInt32LE(cursor);
			while (cursor + 8 <= end) {
				const size = blob.readUInt16LE(cursor);
				if (size < 8) {
					break;
				}
				const dataType = blob[cursor + 3];
				const name = names[blob.readUInt16LE(cursor + 4)];
				const dataSize = blob.readUInt16LE(cursor + 6);
				const data = blob.subarray(cursor + 8, Math.min(cursor + 8 + dataSize, end));
				if (name !== undefined && !out.has(name)) {
					out.set(name, decodeValue(dataType, data));
				}
				cursor += size;
			}
		}
		at = end;
	}
	return out;
}

function decodeValue(dataType: number, data: Buffer): AccessPropertyValue {
	switch (dataType) {
		case TYPE_TEXT:
		case TYPE_MEMO:
			return data.toString('utf16le');
		case TYPE_BOOLEAN:
			return data.length > 0 && data[0] !== 0;
		case TYPE_BYTE:
			return data.length > 0 ? data[0] : 0;
		case TYPE_INTEGER:
			return data.length >= 2 ? data.readInt16LE(0) : 0;
		case TYPE_LONG:
			return data.length >= 4 ? data.readInt32LE(0) : 0;
		default:
			return Buffer.from(data);
	}
}
