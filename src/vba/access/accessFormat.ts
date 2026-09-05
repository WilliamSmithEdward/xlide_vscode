// The Jet 4 / ACE page format, as far as reading a table's rows.
//
// Every offset here was derived from the fixtures in tests/fixtures/binaries
// and checked by what came out: the table-definition layout is confirmed by
// MSysObjects' seventeen column names and types landing exactly, and the row
// layout by the object names the catalog then reads back. Where a published
// offset table and the bytes disagreed, the bytes won - the table-definition
// header is 0x28 for the table type here, not the 0x2C some references give.
//
// This is the structural half of github.com/WilliamSmithEdward/xlide_vscode/issues/65.
// `accessDatabase.ts` finds the VBA project by scanning and decompressing
// candidate rows; that works for reading and cannot support writing, which has
// to know which row on which page holds what.

export class AccessFormatError extends Error {}

export const ACCESS_PAGE_SIZE = 4096;

/** Page types, by the first byte of a page. */
export const enum AccessPageType {
	Header = 0x00,
	Data = 0x01,
	TableDefinition = 0x02,
	IndexNode = 0x03,
	IndexLeaf = 0x04,
	PageUsageBitmap = 0x05,
}

/** The catalog table, which names every other table and where its definition is. */
export const MSYS_OBJECTS_PAGE = 2;

/** Jet column type codes. */
export const enum AccessColumnType {
	Boolean = 1,
	Byte = 2,
	Integer = 3,
	Long = 4,
	Money = 5,
	Single = 6,
	Double = 7,
	DateTime = 8,
	Binary = 9,
	Text = 10,
	Ole = 11,
	Memo = 12,
	Guid = 15,
	Numeric = 16,
	Complex = 18,
}

export interface AccessColumn {
	name: string;
	type: AccessColumnType;
	/** Position in the table's own column order. */
	number: number;
	/** Index among the variable-length columns, or -1 when fixed. */
	variableIndex: number;
	/** Byte offset within a row's fixed-length block, or -1 when variable. */
	fixedOffset: number;
	/** Declared length in bytes. */
	length: number;
	fixed: boolean;
}

export interface AccessTableDefinition {
	/** The page the definition starts on. */
	page: number;
	numRows: number;
	/** 0x53 for a system table, 0x4E for a user one. */
	tableType: number;
	numColumns: number;
	numVariableColumns: number;
	columns: AccessColumn[];
}

/** Whether the buffer is a Jet 4 or ACE database. */
export function isAccessFile(data: Buffer): boolean {
	if (data.length < ACCESS_PAGE_SIZE || data.length % ACCESS_PAGE_SIZE !== 0) {
		return false;
	}
	const signature = data.subarray(4, 20).toString('latin1');
	return signature === 'Standard ACE DB\0' || signature === 'Standard Jet DB\0';
}

export function pageCount(data: Buffer): number {
	return Math.floor(data.length / ACCESS_PAGE_SIZE);
}

export function pageAt(data: Buffer, page: number): Buffer {
	if (page < 0 || page >= pageCount(data)) {
		throw new AccessFormatError(`Page ${page} is outside the database.`);
	}
	return data.subarray(page * ACCESS_PAGE_SIZE, (page + 1) * ACCESS_PAGE_SIZE);
}

export function pageType(data: Buffer, page: number): number {
	return pageAt(data, page)[0];
}

// The table-definition header, all confirmed by the column names coming out
// right on both an .accdb and an .mdb fixture.
const TDEF_TABLE_TYPE = 0x28;
const TDEF_NUM_VAR_COLUMNS = 0x2b;
const TDEF_NUM_COLUMNS = 0x2d;
const TDEF_NUM_REAL_INDEXES = 0x33;
const TDEF_HEADER_SIZE = 0x3f;
const REAL_INDEX_SIZE = 12;
const COLUMN_DEFINITION_SIZE = 25;

/**
 * A table's definition. A definition longer than one page continues on the
 * page named at offset 4, so the pages are joined before anything is read out
 * of them - a column list that spans the boundary is otherwise cut in half.
 */
export function readTableDefinition(data: Buffer, page: number): AccessTableDefinition {
	if (pageType(data, page) !== AccessPageType.TableDefinition) {
		throw new AccessFormatError(`Page ${page} is not a table definition.`);
	}
	const first = pageAt(data, page);
	const chunks: Buffer[] = [first];
	const seen = new Set([page]);
	let next = first.readUInt32LE(4);
	while (next !== 0) {
		if (seen.has(next)) {
			throw new AccessFormatError(`Table definition at page ${page} loops back to page ${next}.`);
		}
		seen.add(next);
		const continuation = pageAt(data, next);
		// A continuation page repeats the 8-byte page header before its data.
		chunks.push(continuation.subarray(8));
		next = continuation.readUInt32LE(4);
	}
	const d = chunks.length === 1 ? first : Buffer.concat(chunks);

	const numColumns = d.readUInt16LE(TDEF_NUM_COLUMNS);
	const numVariableColumns = d.readUInt16LE(TDEF_NUM_VAR_COLUMNS);
	const numRealIndexes = d.readUInt32LE(TDEF_NUM_REAL_INDEXES);
	const definitionsAt = TDEF_HEADER_SIZE + numRealIndexes * REAL_INDEX_SIZE;

	const columns: AccessColumn[] = [];
	for (let i = 0; i < numColumns; i += 1) {
		const at = definitionsAt + i * COLUMN_DEFINITION_SIZE;
		if (at + COLUMN_DEFINITION_SIZE > d.length) {
			throw new AccessFormatError(`Table definition at page ${page} is shorter than its column count.`);
		}
		const flags = d[at + 15];
		const fixed = (flags & 0x01) !== 0;
		columns.push({
			name: '',
			type: d[at] as AccessColumnType,
			number: d.readUInt16LE(at + 5),
			variableIndex: fixed ? -1 : d.readUInt16LE(at + 7),
			fixedOffset: fixed ? d.readUInt16LE(at + 21) : -1,
			length: d.readUInt16LE(at + 23),
			fixed,
		});
	}

	// The names follow the definitions, in the same order: a u16 byte count
	// then UTF-16LE.
	let at = definitionsAt + numColumns * COLUMN_DEFINITION_SIZE;
	for (const column of columns) {
		const length = d.readUInt16LE(at);
		at += 2;
		column.name = d.subarray(at, at + length).toString('utf16le');
		at += length;
	}

	return {
		page,
		numRows: first.readUInt32LE(0x10),
		tableType: d[TDEF_TABLE_TYPE],
		numColumns,
		numVariableColumns,
		columns,
	};
}

/** A value read out of a row. A long value is returned as its reference. */
export type AccessValue = number | string | Buffer | AccessLongValueRef | null;

/**
 * An OLE or Memo column's value. Under about 64 bytes it is stored inline;
 * otherwise the row holds a reference to a Long Value page and the bytes are
 * fetched separately.
 */
export interface AccessLongValueRef {
	kind: 'longValue';
	inline?: Buffer;
	/** Total length of the value in bytes. */
	length: number;
	/** The LVAL page holding it, when not inline. */
	page?: number;
	/** The row on that page. */
	row?: number;
	/** The value spans several LVAL rows, chained. */
	chained?: boolean;
}

export interface AccessRow {
	page: number;
	/** Slot on the page, which with the page identifies the row. */
	slot: number;
	values: Map<string, AccessValue>;
}

const DATA_NUM_RECORDS = 0x0c;
const DATA_RECORD_OFFSETS = 0x0e;
/**
 * A record offset carries flags in its top bits. A 4096-byte page needs twelve
 * bits for an offset, so the mask is 0x0FFF and everything above it is a flag -
 * masking wider lets a flag bit through as part of the offset.
 */
const RECORD_OFFSET_MASK = 0x0fff;
const RECORD_DELETED = 0x8000;
const RECORD_POINTER = 0x4000;

/** Every live row on one data page of the given table. */
export function readDataPageRows(
	data: Buffer,
	page: number,
	definition: AccessTableDefinition,
): AccessRow[] {
	const d = pageAt(data, page);
	if (d[0] !== AccessPageType.Data) {
		throw new AccessFormatError(`Page ${page} is not a data page.`);
	}
	const count = d.readUInt16LE(DATA_NUM_RECORDS);
	const out: AccessRow[] = [];
	for (let slot = 0; slot < count; slot += 1) {
		const raw = d.readUInt16LE(DATA_RECORD_OFFSETS + slot * 2);
		if ((raw & RECORD_DELETED) !== 0 || (raw & RECORD_POINTER) !== 0) {
			// A deleted row leaves its bytes in place, and a pointer row lives
			// somewhere else; neither is a row of this table here.
			continue;
		}
		const start = raw & RECORD_OFFSET_MASK;
		const end = slot === 0
			? ACCESS_PAGE_SIZE
			: (d.readUInt16LE(DATA_RECORD_OFFSETS + (slot - 1) * 2) & RECORD_OFFSET_MASK);
		if (start >= end || end > ACCESS_PAGE_SIZE) {
			continue;
		}
		const values = readRow(d.subarray(start, end), definition);
		if (values) {
			out.push({ page, slot, values });
		}
	}
	return out;
}

/**
 * One row.
 *
 * The layout runs from both ends: a column count and the fixed-length block
 * from the front, and from the back a null bitmap, the variable-column count,
 * and that many offsets plus one for the end of the data - stored in reverse,
 * so the first variable column's offset is the last one written.
 */
function readRow(row: Buffer, definition: AccessTableDefinition): Map<string, AccessValue> | undefined {
	if (row.length < 2) {
		return undefined;
	}
	const columnCount = row.readUInt16LE(0);
	const bitmapBytes = Math.ceil(columnCount / 8);
	if (bitmapBytes + 2 > row.length) {
		return undefined;
	}
	const bitmapAt = row.length - bitmapBytes;
	const isNull = (columnNumber: number): boolean =>
		columnNumber >= columnCount
		|| (row[bitmapAt + (columnNumber >> 3)] & (1 << (columnNumber & 7))) === 0;

	const variableCount = bitmapAt >= 2 ? row.readUInt16LE(bitmapAt - 2) : 0;
	const offsetsAt = bitmapAt - 2 - (variableCount + 1) * 2;
	if (variableCount > definition.numVariableColumns || offsetsAt < 2) {
		return undefined;
	}
	// Reversed: entry i counts back from the end of the offset block.
	const variableOffset = (index: number): number =>
		row.readUInt16LE(offsetsAt + (variableCount - index) * 2);

	const values = new Map<string, AccessValue>();
	for (const column of definition.columns) {
		if (isNull(column.number)) {
			values.set(column.name, null);
			continue;
		}
		if (column.fixed) {
			values.set(column.name, readFixed(row, 2 + column.fixedOffset, column));
			continue;
		}
		if (column.variableIndex < 0 || column.variableIndex >= variableCount) {
			values.set(column.name, null);
			continue;
		}
		const start = variableOffset(column.variableIndex);
		const end = variableOffset(column.variableIndex + 1);
		if (start > end || end > row.length) {
			values.set(column.name, null);
			continue;
		}
		values.set(column.name, readVariable(row.subarray(start, end), column));
	}
	return values;
}

function readFixed(row: Buffer, at: number, column: AccessColumn): AccessValue {
	if (at + column.length > row.length) {
		return null;
	}
	return readScalar(row.subarray(at, at + column.length), column);
}

/**
 * A scalar from exactly the bytes that hold it.
 *
 * The width is not always the declared one. A column may be declared
 * variable-length and still hold a number - `MSysAccessStorage` declares all
 * seven of its columns that way, `Id`, `ParentId` and `Type` included - and
 * Access then writes only the bytes the value needs, so a `Type` of 1 arrives
 * as a single byte where the column says four. Short values are zero-extended
 * to the declared width before being read.
 */
function readScalar(bytes: Buffer, column: AccessColumn): AccessValue {
	const width = numericWidth(column.type);
	if (width === undefined) {
		return Buffer.from(bytes);
	}
	if (bytes.length === 0) {
		return null;
	}
	const b = bytes.length >= width
		? bytes
		: Buffer.concat([bytes, Buffer.alloc(width - bytes.length)]);
	switch (column.type) {
		case AccessColumnType.Boolean:
			return b[0] !== 0 ? 1 : 0;
		case AccessColumnType.Byte:
			return b[0];
		case AccessColumnType.Integer:
			return b.readInt16LE(0);
		case AccessColumnType.Long:
			return b.readInt32LE(0);
		case AccessColumnType.Single:
			return b.readFloatLE(0);
		case AccessColumnType.Double:
		case AccessColumnType.DateTime:
			// A date is an OLE automation double; a caller wanting a Date
			// converts it, rather than this deciding what a date means.
			return b.readDoubleLE(0);
		case AccessColumnType.Money:
			return Number(b.readBigInt64LE(0)) / 10000;
		default:
			return Buffer.from(bytes);
	}
}

/** The declared width of a numeric type, or undefined when it is not one. */
function numericWidth(type: AccessColumnType): number | undefined {
	switch (type) {
		case AccessColumnType.Boolean:
		case AccessColumnType.Byte:
			return 1;
		case AccessColumnType.Integer:
			return 2;
		case AccessColumnType.Long:
		case AccessColumnType.Single:
			return 4;
		case AccessColumnType.Double:
		case AccessColumnType.DateTime:
		case AccessColumnType.Money:
			return 8;
		default:
			return undefined;
	}
}

/** How a long value's reference is encoded, in the four bits above its length. */
const LVAL_INLINE = 0x80;
const LVAL_SINGLE_PAGE = 0x40;

function readVariable(bytes: Buffer, column: AccessColumn): AccessValue {
	switch (column.type) {
		case AccessColumnType.Text:
			return decodeText(bytes);
		case AccessColumnType.Ole:
		case AccessColumnType.Memo:
			return readLongValueRef(bytes);
		default:
			// Not necessarily bytes: a number can be declared variable-length.
			return readScalar(bytes, column);
	}
}

/**
 * Jet 4 text is UTF-16LE, except where it is compressed: a `0xFF 0xFE` marker
 * TOGGLES between one byte per character and two, and a string may carry
 * several. Treating the first marker as "single-byte from here on" reads a
 * string that switches back as mojibake from that point.
 */
export const decodeTextForTests = (bytes: Buffer): string => decodeText(bytes);

function decodeText(bytes: Buffer): string {
	if (bytes.length < 2 || bytes[0] !== 0xff || bytes[1] !== 0xfe) {
		return bytes.toString('utf16le');
	}
	let out = '';
	let compressed = false;
	let at = 0;
	while (at + 1 < bytes.length || (compressed && at < bytes.length)) {
		if (at + 1 < bytes.length && bytes[at] === 0xff && bytes[at + 1] === 0xfe) {
			compressed = !compressed;
			at += 2;
			continue;
		}
		if (compressed) {
			out += String.fromCharCode(bytes[at]);
			at += 1;
		} else {
			out += String.fromCharCode(bytes.readUInt16LE(at));
			at += 2;
		}
	}
	return out;
}

/** Guards a corrupt chain from spinning forever. */
const MAX_LONG_VALUE_CHUNKS = 4096;

/**
 * The bytes a long value holds, fetched from its LVAL page when they are not
 * inline.
 *
 * A single-page value is the row itself. A chained one opens each chunk with
 * `u8 next slot, u24 next page`, (0, 0) ending it, and the value is the chunks
 * with those prefixes stripped.
 */
export function readLongValue(data: Buffer, ref: AccessLongValueRef): Buffer {
	if (ref.inline) {
		return ref.inline;
	}
	if (ref.page === undefined || ref.row === undefined) {
		throw new AccessFormatError('Long value reference names neither inline bytes nor a page.');
	}
	if (!ref.chained) {
		return longValueRowBytes(data, ref.page, ref.row).subarray(0, ref.length);
	}

	const parts: Buffer[] = [];
	const seen = new Set<number>();
	let page = ref.page;
	let row = ref.row;
	for (let i = 0; i < MAX_LONG_VALUE_CHUNKS; i += 1) {
		const key = page * 65536 + row;
		if (seen.has(key)) {
			throw new AccessFormatError(`Long value chain loops at page ${page} row ${row}.`);
		}
		seen.add(key);
		const chunk = longValueRowBytes(data, page, row);
		if (chunk.length < 4) {
			throw new AccessFormatError(`Long value chunk at page ${page} row ${row} has no chain prefix.`);
		}
		parts.push(chunk.subarray(4));
		const nextRow = chunk[0];
		const nextPage = chunk.readUIntLE(1, 3);
		if (nextPage === 0 && nextRow === 0) {
			return Buffer.concat(parts).subarray(0, ref.length);
		}
		page = nextPage;
		row = nextRow;
	}
	throw new AccessFormatError(`Long value chain is longer than ${MAX_LONG_VALUE_CHUNKS} chunks.`);
}

/**
 * One row of an LVAL page. The slot table works as it does on any data page:
 * rows grow down from the end, so a row runs from its own offset to the next
 * live offset above it.
 */
function longValueRowBytes(data: Buffer, page: number, row: number): Buffer {
	const d = pageAt(data, page);
	const count = d.readUInt16LE(DATA_NUM_RECORDS);
	if (row >= count) {
		throw new AccessFormatError(`Page ${page} has no row ${row}.`);
	}
	const offsets: number[] = [];
	for (let i = 0; i < count; i += 1) {
		offsets.push(d.readUInt16LE(DATA_RECORD_OFFSETS + i * 2));
	}
	if ((offsets[row] & RECORD_DELETED) !== 0) {
		throw new AccessFormatError(`Page ${page} row ${row} is deleted.`);
	}
	const start = offsets[row] & RECORD_OFFSET_MASK;
	let end = ACCESS_PAGE_SIZE;
	for (const other of offsets) {
		if ((other & RECORD_DELETED) !== 0) {
			continue;
		}
		const offset = other & RECORD_OFFSET_MASK;
		if (offset > start && offset < end) {
			end = offset;
		}
	}
	return d.subarray(start, end);
}

/**
 * A Memo or OLE reference: twelve bytes, of which the first four carry the
 * length and a flag saying where the bytes are.
 */
function readLongValueRef(bytes: Buffer): AccessLongValueRef | null {
	if (bytes.length < 12) {
		return bytes.length > 0 ? { kind: 'longValue', inline: Buffer.from(bytes), length: bytes.length } : null;
	}
	const header = bytes.readUInt32LE(0);
	const length = header & 0x00ffffff;
	const flags = (header >>> 24) & 0xff;
	if ((flags & LVAL_INLINE) !== 0) {
		return { kind: 'longValue', inline: Buffer.from(bytes.subarray(12, 12 + length)), length };
	}
	const pointer = bytes.readUInt32LE(4);
	return {
		kind: 'longValue',
		length,
		row: pointer & 0xff,
		page: pointer >>> 8,
		chained: (flags & LVAL_SINGLE_PAGE) === 0,
	};
}
