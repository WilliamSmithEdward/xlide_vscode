import { ACCESS_PAGE_SIZE, AccessFormatError } from './accessFormat';
import {
	AccessDataPage,
	LVAL_OWNER_TAG,
	PAGE_DATA,
	decodeRowPointer,
	encodeRowPointer,
	isLongValuePage,
} from './accessDataPage';
import { AccessPageStore, readUsageMapRef } from './accessPageStore';
import { addToMap, allocatePage, releasePage, removeFromMap } from './accessAlloc';

/**
 * Long values: the Memo and OLE data that does not live in the row.
 *
 * The row holds a 12-byte definition whose kind byte says where the bytes
 * are: 0x80 right after the definition, 0x40 one whole row on a long-value
 * page, 0x00 a chain of such rows each opening with a pointer to the next.
 *
 * Ported from pyOpenVBA's `_lval.py`. Where the engine puts a value was
 * measured on values Access wrote, and the sizes below are those
 * measurements, not a specification.
 */

/** A value this size or smaller goes in the row itself. */
const LVAL_INLINE_MAX = 64;
/** Above this a value is chained rather than given one row. */
const LVAL_SINGLE_MAX = 3816;
/** A chained chunk's payload, after its four-byte pointer. */
const LVAL_CHUNK_PAYLOAD = 4072;
/** A page is listed in the free-space map while more than 256 bytes are free. */
const LVAL_PAGE_MIN_FREE = 257;
/** So any listed page has room for a value of this size or smaller. */
const LVAL_SHARED_MAX = 256;
/** The word a chained value's first page carries, matching its definition. */
const OFFSET_LVAL_STAMP = 0x08;
const OFFSET_FREE_SPACE = 0x02;
const OFFSET_OWNER = 0x04;

const KIND_INLINE = 0x80;
const KIND_SINGLE_PAGE = 0x40;
const KIND_CHAINED = 0x00;

/** The 12-byte definition a row holds for a long value. */
export interface AccessLongValue {
	length: number;
	kind: number;
	/** The bytes themselves, for an inline value. */
	inline: Buffer;
	page: number;
	row: number;
}

/** A long-value column's `(owned pages, free-space pages)` map references. */
export interface AccessLongValueMaps {
	owned: number;
	/** Zero on a column the engine gave no free-space map. */
	free: number;
}

export function decodeLongValue(raw: Buffer): AccessLongValue {
	if (raw.length < 12) {
		throw new AccessFormatError(`A long-value definition is 12 bytes, got ${raw.length}.`);
	}
	const length = raw.readUIntLE(0, 3);
	const kind = raw[3];
	const inline = kind === KIND_INLINE ? raw.subarray(12) : Buffer.alloc(0);
	if (kind === KIND_INLINE && inline.length !== length) {
		throw new AccessFormatError(
			`An inline long value declares ${length} bytes but carries ${inline.length}.`,
		);
	}
	return { length, kind, inline, row: raw[4], page: raw.readUIntLE(5, 3) };
}

function definition(length: number, kind: number, row: number, page: number, stamp: number): Buffer {
	if (length >= 1 << 24) {
		throw new AccessFormatError(`A long value of ${length} bytes exceeds the format's limit.`);
	}
	const out = Buffer.alloc(12);
	out.writeUIntLE(length, 0, 3);
	out[3] = kind;
	out[4] = row;
	out.writeUIntLE(page, 5, 3);
	out.writeUInt32LE(stamp, 8);
	return out;
}

/** A fresh long-value page, taken from the free map. */
export function newLongValuePage(store: AccessPageStore, stamp = 0): number {
	const page = allocatePage(store);
	const raw = Buffer.alloc(ACCESS_PAGE_SIZE);
	raw[0] = PAGE_DATA;
	raw[1] = 0x01;
	raw.writeUInt16LE(ACCESS_PAGE_SIZE - 14, OFFSET_FREE_SPACE);
	raw.writeUInt32LE(LVAL_OWNER_TAG, OFFSET_OWNER);
	raw.writeUInt32LE(stamp, OFFSET_LVAL_STAMP);
	store.write(page, raw);
	return page;
}

/**
 * Store `data` for a long-value column and return the bytes its row should
 * hold: the 12-byte definition, followed by the data itself when inline.
 */
export function writeLongValue(
	store: AccessPageStore,
	maps: AccessLongValueMaps,
	data: Buffer,
	stamp: number,
): Buffer {
	if (data.length <= LVAL_INLINE_MAX) {
		return Buffer.concat([definition(data.length, KIND_INLINE, 0, 0, 0), data]);
	}
	if (data.length <= LVAL_SINGLE_MAX) {
		const page = singleRowPage(store, maps, data.length);
		const lval = new AccessDataPage(store.read(page));
		const slot = lval.addRow(data);
		store.write(page, lval.toBuffer());
		if (lval.freeSpace < LVAL_PAGE_MIN_FREE && maps.free) {
			removeFromMap(store, maps.free, page);
		}
		return definition(data.length, KIND_SINGLE_PAGE, slot, page, 0);
	}
	// A chain: one fresh page per chunk, linked front to back, so the pages
	// are made first and the pointers filled in from the last one back.
	const chunks: Buffer[] = [];
	for (let at = 0; at < data.length; at += LVAL_CHUNK_PAYLOAD) {
		chunks.push(data.subarray(at, at + LVAL_CHUNK_PAYLOAD));
	}
	const pages = chunks.map((_chunk, index) => newLongValuePage(store, index === 0 ? stamp : 0));
	for (const page of pages) {
		addToMap(store, maps.owned, page);
	}
	let next = encodeRowPointer(0, 0);
	for (let index = chunks.length - 1; index >= 0; index -= 1) {
		const page = pages[index];
		const lval = new AccessDataPage(store.read(page));
		const slot = lval.addRow(Buffer.concat([next, chunks[index]]));
		store.write(page, lval.toBuffer());
		next = encodeRowPointer(page, slot);
	}
	return definition(data.length, KIND_CHAINED, 0, pages[0], stamp);
}

/**
 * A long-value page of this column with room for a row of `length` bytes,
 * chosen as the engine chooses.
 *
 * The free-space map lists a page while more than 256 bytes are free, so any
 * listed page has room for a value of 256 bytes or fewer and the engine takes
 * the first of them. A larger value needs the page checked, and the engine
 * checks one: the highest-numbered page the map lists. When that page cannot
 * take the value the engine starts another rather than looking further back.
 */
function singleRowPage(
	store: AccessPageStore,
	maps: AccessLongValueMaps,
	length: number,
): number {
	// A column the engine gave no free-space map has no listed pages to look
	// through, and every value it stores starts a page of its own.
	const listed = (maps.free ? readUsageMapRef(store, maps.free).pages() : [])
		.filter((page) => page < store.pageCount && isLongValuePage(store.read(page)));
	for (const candidate of length <= LVAL_SHARED_MAX ? listed : listed.slice(-1)) {
		if (new AccessDataPage(store.read(candidate)).fits(length)) {
			return candidate;
		}
	}
	const page = newLongValuePage(store);
	addToMap(store, maps.owned, page);
	if (maps.free) {
		addToMap(store, maps.free, page);
	}
	return page;
}

/**
 * Give back what a long value occupied: nothing for an inline one, the row
 * for a single-page value, every page for a chain. Content is left in place,
 * as the engine leaves it.
 */
export function freeLongValue(
	store: AccessPageStore,
	maps: AccessLongValueMaps,
	value: AccessLongValue,
): void {
	if (value.kind === KIND_INLINE) {
		return;
	}
	if (value.kind === KIND_SINGLE_PAGE) {
		const lval = new AccessDataPage(store.read(value.page));
		lval.removeRow(value.row);
		if (lval.liveRows > 0 && lval.freeSpace >= LVAL_PAGE_MIN_FREE) {
			// Room came back, so the page is listed again.
			store.write(value.page, lval.toBuffer());
			if (maps.free) {
				addToMap(store, maps.free, value.page);
			}
			return;
		}
		if (lval.liveRows === 0) {
			// A long-value page that lost its last value is retired: type
			// 0x09, released, and out of both of the column's maps.
			lval.retire();
			store.write(value.page, lval.toBuffer());
			releasePage(store, value.page);
			removeFromMap(store, maps.owned, value.page);
			if (maps.free) {
				removeFromMap(store, maps.free, value.page);
			}
			return;
		}
		store.write(value.page, lval.toBuffer());
		return;
	}
	if (value.kind === KIND_CHAINED) {
		for (const { page } of longValueChain(store, value)) {
			releasePage(store, page, 'value');
			removeFromMap(store, maps.owned, page);
		}
		return;
	}
	throw new AccessFormatError(`Unknown long-value kind 0x${value.kind.toString(16)}.`);
}

/** The bytes a long value holds, gathered from wherever they live. */
export function readLongValueFrom(store: AccessPageStore, value: AccessLongValue): Buffer {
	if (value.kind === KIND_INLINE) {
		return value.inline;
	}
	if (value.kind === KIND_SINGLE_PAGE) {
		const row = longValueRow(store, value.page, value.row);
		if (row.length !== value.length) {
			throw new AccessFormatError(
				`The long value at (${value.page}, ${value.row}) holds ${row.length} bytes, `
				+ `its definition says ${value.length}.`,
			);
		}
		return row;
	}
	if (value.kind === KIND_CHAINED) {
		const parts = longValueChain(store, value).map(({ chunk }) => chunk.subarray(4));
		const out = Buffer.concat(parts);
		if (out.length !== value.length) {
			throw new AccessFormatError(
				`A long value chain holds ${out.length} bytes, its definition says ${value.length}.`,
			);
		}
		return out;
	}
	throw new AccessFormatError(`Unknown long-value kind 0x${value.kind.toString(16)}.`);
}

function longValueChain(
	store: AccessPageStore,
	value: AccessLongValue,
): Array<{ page: number; row: number; chunk: Buffer }> {
	const out: Array<{ page: number; row: number; chunk: Buffer }> = [];
	const seen = new Set<number>();
	let { page, row } = value;
	while (page !== 0) {
		const key = page * 256 + row;
		if (seen.has(key)) {
			throw new AccessFormatError(`A long value chain loops at (${page}, ${row}).`);
		}
		seen.add(key);
		const chunk = longValueRow(store, page, row);
		if (chunk.length < 4) {
			throw new AccessFormatError(
				`The long value chunk at (${page}, ${row}) is too short for its pointer.`,
			);
		}
		out.push({ page, row, chunk });
		({ page, row } = decodeRowPointer(chunk));
	}
	return out;
}

function longValueRow(store: AccessPageStore, page: number, row: number): Buffer {
	const raw = store.read(page);
	if (!isLongValuePage(raw)) {
		throw new AccessFormatError(`Page ${page} is not a long-value page.`);
	}
	const bytes = new AccessDataPage(raw).row(row);
	if (bytes === undefined) {
		throw new AccessFormatError(`The long value row (${page}, ${row}) is deleted.`);
	}
	return bytes;
}
