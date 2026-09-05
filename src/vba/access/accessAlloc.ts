import { ACCESS_PAGE_SIZE, AccessFormatError } from './accessFormat';
import { AccessDataPage } from './accessDataPage';
import {
	AccessPageStore,
	AccessUsageMap,
	GLOBAL_USAGE_MAP_PAGE,
	GLOBAL_USAGE_MAP_ROW,
	PAGES_PER_BITMAP_PAGE,
	PAGE_USAGE_BITMAP,
	USAGE_BITMAP_PAGE_DATA,
	USAGE_MAP_INLINE,
	USAGE_MAP_REFERENCE,
	readUsageMap,
	readUsageMapRef,
} from './accessPageStore';

/**
 * Page allocation and usage-map maintenance.
 *
 * The global map on page 1 row 0 marks FREE pages, pages past the end of the
 * file included; the engine takes the lowest free page and grows the file to
 * reach it. A table's owned-pages and free-space maps are the same structure
 * and are edited in place.
 *
 * Ported from pyOpenVBA's `_alloc.py`. The growth rules in it were measured on
 * maps the engine grew, and the comments say what was measured, because none of
 * it is derivable: a map covers two pages past the one being added, rounded up
 * to four bytes, and thirty growths across five scenarios agree on that.
 */

/** The global map is extended a whole byte-step at a time, 64 pages. */
const INLINE_BITMAP_STEP = 8;
/**
 * A map's bitmap is sized up to a multiple of this to cover a new page,
 * measured on maps the engine grew: 3756 bytes for 30,048 pages, which is a
 * multiple of four and not of eight.
 */
const INLINE_BITMAP_ROUND = 4;
/** A reference map's row: a kind byte and seventeen chunk pointers. */
const REFERENCE_CHUNKS = 17;
/** Pages waiting to come back into use are let go once this many have piled up. */
const PENDING_FLUSH = 5;

/** Who gave a page back, which decides when it may be taken again. */
export type ReleaseKind = 'object' | 'value' | 'rewrite';

/** Write an inline map's row back, resizing it when the bitmap grew. */
function rewriteInlineRow(store: AccessPageStore, map: AccessUsageMap): void {
	const row = Buffer.alloc(5 + map.bitmap.length);
	row[0] = USAGE_MAP_INLINE;
	row.writeUInt32LE(map.startPage, 1);
	map.bitmap.copy(row, 5);
	const page = new AccessDataPage(store.read(map.page));
	page.replaceRow(map.row, row);
	store.write(map.page, page.toBuffer());
}

/**
 * Make an inline map able to hold `page`. An empty map re-bases to the page's
 * 8-aligned start; a map that holds pages grows to cover TWO pages past the one
 * being added, rounded up to four bytes. The two spare pages are what a page
 * taken from the end of the file leaves ahead of it.
 */
function reach(map: AccessUsageMap, page: number): void {
	const withinBitmap = page >= map.startPage
		&& Math.floor((page - map.startPage) / 8) < map.bitmap.length;
	if (withinBitmap) {
		return;
	}
	if (!map.bitmap.some((byte) => byte !== 0)) {
		map.startPage = page & ~7;
		return;
	}
	if (page < map.startPage) {
		throw new AccessFormatError(
			`Page ${page} lies below the start (${map.startPage}) of the usage map at `
			+ `(${map.page}, ${map.row}); the engine's answer to that has not been measured.`,
		);
	}
	const neededBytes = Math.floor((page + 2 - map.startPage) / 8) + 1;
	const rounded = Math.ceil(neededBytes / INLINE_BITMAP_ROUND) * INLINE_BITMAP_ROUND;
	if (rounded > map.bitmap.length) {
		map.bitmap = Buffer.concat([map.bitmap, Buffer.alloc(rounded - map.bitmap.length)]);
	}
}

/** Whether the inline row, at its current bitmap size, still fits its page. */
function inlineRowFits(store: AccessPageStore, map: AccessUsageMap): boolean {
	const page = new AccessDataPage(store.read(map.page));
	const [start, end] = page.span(map.row);
	return 5 + map.bitmap.length - (end - start) <= page.freeSpace;
}

function flip(bitmap: Buffer, index: number, present: boolean): void {
	if (index < 0 || index >> 3 >= bitmap.length) {
		return;
	}
	if (present) {
		bitmap[index >> 3] |= 1 << (index % 8);
	} else {
		bitmap[index >> 3] &= ~(1 << (index % 8)) & 0xff;
	}
}

/** Read a map back after its shape changed. */
function refresh(store: AccessPageStore, map: AccessUsageMap): AccessUsageMap {
	return readUsageMap(store, map.page, map.row);
}

function writeReferenceSlot(
	store: AccessPageStore,
	map: AccessUsageMap,
	chunk: number,
	bitmapPage: number,
): void {
	const raw = Buffer.from(store.read(map.page));
	const [start] = new AccessDataPage(raw).span(map.row);
	raw.writeUInt32LE(bitmapPage, start + 1 + 4 * chunk);
	store.write(map.page, raw);
}

/** Set or clear a page in a map and write the change through. */
export function setUsageBit(
	store: AccessPageStore,
	map: AccessUsageMap,
	page: number,
	present: boolean,
): AccessUsageMap {
	if (map.kind === USAGE_MAP_INLINE) {
		if (present) {
			reach(map, page);
			if (!inlineRowFits(store, map)) {
				const converted = toReference(store, map);
				return setUsageBit(store, converted, page, true);
			}
		}
		const index = page - map.startPage;
		if (index < 0 || index >> 3 >= map.bitmap.length) {
			if (!present) {
				// Clearing a page the map cannot hold changes nothing.
				return map;
			}
			throw new AccessFormatError(
				`Page ${page} is outside the usage map at (${map.page}, ${map.row}).`,
			);
		}
		flip(map.bitmap, index, present);
		rewriteInlineRow(store, map);
		return map;
	}
	if (map.kind === USAGE_MAP_REFERENCE) {
		const chunk = Math.floor(page / PAGES_PER_BITMAP_PAGE);
		const within = page % PAGES_PER_BITMAP_PAGE;
		if (chunk >= map.referencePages.length) {
			throw new AccessFormatError(
				`Page ${page} is beyond the reference usage map at (${map.page}, ${map.row}).`,
			);
		}
		let bitmapPage = map.referencePages[chunk];
		if (bitmapPage === 0) {
			if (!present) {
				return map;
			}
			bitmapPage = newBitmapPage(store);
			map.referencePages[chunk] = bitmapPage;
			writeReferenceSlot(store, map, chunk, bitmapPage);
		}
		const raw = Buffer.from(store.read(bitmapPage));
		const byteIndex = USAGE_BITMAP_PAGE_DATA + (within >> 3);
		if (present) {
			raw[byteIndex] |= 1 << (within % 8);
		} else {
			raw[byteIndex] &= ~(1 << (within % 8)) & 0xff;
		}
		store.write(bitmapPage, raw);
		flip(map.bitmap, page, present);
		return map;
	}
	throw new AccessFormatError(`Usage map kind ${map.kind} cannot be edited.`);
}

/**
 * Turn an inline map into the reference form: the pages it holds move onto
 * bitmap pages of their own and the row becomes a list of those pages. The
 * engine does this when growing the inline bitmap would push its row off the
 * page it lives on.
 */
function toReference(
	store: AccessPageStore,
	map: AccessUsageMap,
	globalMap = false,
): AccessUsageMap {
	const held = map.pages();
	const row = Buffer.alloc(1 + 4 * REFERENCE_CHUNKS);
	row[0] = USAGE_MAP_REFERENCE;
	const page = new AccessDataPage(store.read(map.page));
	page.replaceRow(map.row, row);
	store.write(map.page, page.toBuffer());
	let converted = refresh(store, map);
	if (globalMap) {
		// The global map lists free pages, so its first chunk takes over what
		// was free and everything the file has not reached yet.
		converted = globalChunk(store, converted);
	}
	for (const number of held) {
		converted = setUsageBit(store, converted, number, true);
	}
	return converted;
}

/**
 * Give the global free map its next bitmap page and mark every page of that
 * chunk the file has not reached as free. The bitmap page is the one just past
 * the end of the file, which is where the engine put both measured.
 */
function globalChunk(store: AccessPageStore, map: AccessUsageMap): AccessUsageMap {
	const chunk = map.referencePages.findIndex((page) => page === 0);
	if (chunk === -1) {
		throw new AccessFormatError('The global usage map has no room for another chunk.');
	}
	const bitmapPage = store.append();
	const raw = Buffer.alloc(ACCESS_PAGE_SIZE);
	raw[0] = PAGE_USAGE_BITMAP;
	raw[1] = 0x01;
	const base = chunk * PAGES_PER_BITMAP_PAGE;
	for (let page = Math.max(store.pageCount, base); page < base + PAGES_PER_BITMAP_PAGE; page += 1) {
		const index = page - base;
		raw[USAGE_BITMAP_PAGE_DATA + (index >> 3)] |= 1 << (index % 8);
	}
	store.write(bitmapPage, raw);
	writeReferenceSlot(store, map, chunk, bitmapPage);
	return refresh(store, map);
}

function newBitmapPage(store: AccessPageStore): number {
	const page = allocatePage(store);
	const raw = Buffer.alloc(ACCESS_PAGE_SIZE);
	raw[0] = PAGE_USAGE_BITMAP;
	raw[1] = 0x01;
	store.write(page, raw);
	return page;
}

/**
 * Take the lowest free page from the global map, growing the file if that page
 * lies past its end, and mark it used. Pages released during this session are
 * passed over, as the engine passes them over until the database is reopened.
 */
export function allocatePage(store: AccessPageStore): number {
	let free = readUsageMap(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW);
	const held = new Set([...store.released, ...store.pending]);
	let candidates = free.pages().filter((page) => !held.has(page));
	if (candidates.length === 0) {
		if (free.kind === USAGE_MAP_INLINE) {
			const grown = new AccessUsageMap(
				free.page, free.row, free.kind, free.startPage,
				Buffer.concat([free.bitmap, Buffer.alloc(INLINE_BITMAP_STEP, 0xff)]), [],
			);
			if (inlineRowFits(store, grown)) {
				free.bitmap = grown.bitmap;
				rewriteInlineRow(store, free);
			} else {
				free = toReference(store, free, true);
			}
		} else {
			free = globalChunk(store, free);
		}
		candidates = free.pages().filter((page) => !held.has(page));
	}
	if (candidates.length === 0) {
		throw new AccessFormatError('The database has no free page and the map could not be grown.');
	}
	const page = candidates[0];
	while (store.pageCount <= page) {
		store.append();
	}
	setUsageBit(store, free, page, false);
	store.allocated.add(page);
	return page;
}

/**
 * Return a page to the global free map. When the session may take it again
 * depends on what gave it back, all measured with DAO:
 *
 * - `object` (a dropped table, a retired page): never again this session.
 * - `value` (a freed long-value chain): at once when the page predates the
 *   session, else it waits like a rewrite's page.
 * - `rewrite`: waits with the others, and they all come back once five wait.
 */
export function releasePage(store: AccessPageStore, page: number, kind: ReleaseKind = 'object'): void {
	const free = readUsageMap(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW);
	setUsageBit(store, free, page, true);
	if (kind === 'object') {
		store.released.add(page);
		return;
	}
	if (kind === 'value' && !store.allocated.has(page)) {
		store.released.delete(page);
		return;
	}
	const waiting = store.pending;
	waiting.push(page);
	if (waiting.length >= PENDING_FLUSH) {
		waiting.length = 0;
	}
}

export function addToMap(store: AccessPageStore, reference: number, page: number): void {
	setUsageBit(store, readUsageMapRef(store, reference), page, true);
}

export function removeFromMap(store: AccessPageStore, reference: number, page: number): void {
	setUsageBit(store, readUsageMapRef(store, reference), page, false);
}
