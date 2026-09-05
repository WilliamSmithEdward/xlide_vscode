import { ACCESS_PAGE_SIZE, AccessFormatError, pageCount } from './accessFormat';
import { AccessDataPage } from './accessDataPage';

/**
 * The database as a mutable set of pages, and the usage maps that say which
 * pages are in use.
 *
 * A usage map is a bitmap over page numbers. It comes in two shapes: INLINE,
 * where the bitmap sits in the row itself after a start page, and REFERENCE,
 * where the row holds a list of page numbers and each of those pages carries
 * one chunk of the bitmap. The global map on page 1 row 0 marks FREE pages -
 * including pages past the end of the file, which is how the engine grows.
 * Every other map marks pages that are in use by something.
 *
 * Ported from pyOpenVBA's `_pages.py`.
 */

export const PAGE_USAGE_BITMAP = 0x05;
export const USAGE_MAP_INLINE = 0;
export const USAGE_MAP_REFERENCE = 1;
/** A bitmap page's own header, before its bitmap. */
export const USAGE_BITMAP_PAGE_DATA = 4;
export const PAGES_PER_BITMAP_PAGE = (ACCESS_PAGE_SIZE - USAGE_BITMAP_PAGE_DATA) * 8;

export const GLOBAL_USAGE_MAP_PAGE = 1;
export const GLOBAL_USAGE_MAP_ROW = 0;

/** Split a u32 usage-map reference into the page and row that hold it. */
export function usageMapLocation(reference: number): { page: number; row: number } {
	return { page: reference >>> 8, row: reference & 0xff };
}

/** A decoded usage map, plus where it lives so it can be written back. */
export class AccessUsageMap {
	constructor(
		readonly page: number,
		readonly row: number,
		readonly kind: number,
		public startPage: number,
		public bitmap: Buffer,
		readonly referencePages: number[],
	) {}

	/** Every page the map marks. */
	pages(): number[] {
		const out: number[] = [];
		if (this.kind === USAGE_MAP_INLINE) {
			for (let byteIndex = 0; byteIndex < this.bitmap.length; byteIndex += 1) {
				const byte = this.bitmap[byteIndex];
				if (byte === 0) { continue; }
				for (let bit = 0; bit < 8; bit += 1) {
					if (byte & (1 << bit)) {
						out.push(this.startPage + byteIndex * 8 + bit);
					}
				}
			}
			return out;
		}
		const span = ACCESS_PAGE_SIZE - USAGE_BITMAP_PAGE_DATA;
		for (let chunk = 0; chunk < this.referencePages.length; chunk += 1) {
			const base = chunk * PAGES_PER_BITMAP_PAGE;
			const bytes = this.bitmap.subarray(chunk * span, (chunk + 1) * span);
			for (let byteIndex = 0; byteIndex < bytes.length; byteIndex += 1) {
				const byte = bytes[byteIndex];
				if (byte === 0) { continue; }
				for (let bit = 0; bit < 8; bit += 1) {
					if (byte & (1 << bit)) {
						out.push(base + byteIndex * 8 + bit);
					}
				}
			}
		}
		return out;
	}

	contains(page: number): boolean {
		const index = this.kind === USAGE_MAP_INLINE ? page - this.startPage : page;
		if (index < 0 || Math.floor(index / 8) >= this.bitmap.length) {
			return false;
		}
		return (this.bitmap[index >> 3] & (1 << (index % 8))) !== 0;
	}
}

/**
 * The pages of a database, read and written by number. Growing the file is
 * appending a zeroed page, which is what the engine does when it takes a free
 * page that lies past the end.
 */
export class AccessPageStore {
	private pages: Buffer[];
	/** Pages this session gave back, which the engine passes over until reopen. */
	readonly released = new Set<number>();
	readonly allocated = new Set<number>();
	/** Pages given back this session that are waiting to come round again. */
	readonly pending: number[] = [];

	constructor(data: Buffer) {
		if (data.length === 0 || data.length % ACCESS_PAGE_SIZE !== 0) {
			throw new AccessFormatError('A database is a whole number of 4096-byte pages.');
		}
		this.pages = [];
		for (let page = 0; page < pageCount(data); page += 1) {
			this.pages.push(Buffer.from(data.subarray(page * ACCESS_PAGE_SIZE, (page + 1) * ACCESS_PAGE_SIZE)));
		}
	}

	get pageCount(): number {
		return this.pages.length;
	}

	read(page: number): Buffer {
		if (page < 0 || page >= this.pages.length) {
			throw new AccessFormatError(`Page ${page} is outside the database.`);
		}
		return this.pages[page];
	}

	write(page: number, bytes: Buffer): void {
		if (bytes.length !== ACCESS_PAGE_SIZE) {
			throw new AccessFormatError(`A page is ${ACCESS_PAGE_SIZE} bytes, got ${bytes.length}.`);
		}
		if (page < 0 || page >= this.pages.length) {
			throw new AccessFormatError(`Page ${page} is outside the database.`);
		}
		this.pages[page] = Buffer.from(bytes);
	}

	/** Grow the file by one zeroed page and return its number. */
	append(): number {
		this.pages.push(Buffer.alloc(ACCESS_PAGE_SIZE));
		return this.pages.length - 1;
	}

	toBuffer(): Buffer {
		return Buffer.concat(this.pages);
	}
}

/** A row's bytes on a data page, or undefined when the slot holds none. */
export function rowBytes(page: Buffer, row: number): Buffer | undefined {
	return new AccessDataPage(page).row(row);
}

export function readUsageMap(store: AccessPageStore, page: number, row: number): AccessUsageMap {
	const raw = rowBytes(store.read(page), row);
	if (raw === undefined) {
		throw new AccessFormatError(`Usage map row (${page}, ${row}) is deleted.`);
	}
	if (raw.length === 0) {
		throw new AccessFormatError(`Usage map row (${page}, ${row}) is empty.`);
	}
	const kind = raw[0];
	if (kind === USAGE_MAP_INLINE) {
		if (raw.length < 5) {
			throw new AccessFormatError(`Inline usage map (${page}, ${row}) is truncated.`);
		}
		return new AccessUsageMap(
			page, row, kind, raw.readUInt32LE(1), Buffer.from(raw.subarray(5)), [],
		);
	}
	if (kind === USAGE_MAP_REFERENCE) {
		const count = Math.floor((raw.length - 1) / 4);
		const refs: number[] = [];
		for (let i = 0; i < count; i += 1) {
			refs.push(raw.readUInt32LE(1 + i * 4));
		}
		const span = ACCESS_PAGE_SIZE - USAGE_BITMAP_PAGE_DATA;
		const chunks: Buffer[] = [];
		for (const ref of refs) {
			if (ref === 0) {
				chunks.push(Buffer.alloc(span));
				continue;
			}
			const chunk = store.read(ref);
			if (chunk[0] !== PAGE_USAGE_BITMAP) {
				throw new AccessFormatError(
					`Usage map (${page}, ${row}) references page ${ref}, which is type `
					+ `0x${chunk[0].toString(16)}, not a usage bitmap.`,
				);
			}
			chunks.push(Buffer.from(chunk.subarray(USAGE_BITMAP_PAGE_DATA)));
		}
		return new AccessUsageMap(page, row, kind, 0, Buffer.concat(chunks), refs);
	}
	throw new AccessFormatError(`Usage map (${page}, ${row}) has unknown kind ${kind}.`);
}

export function readUsageMapRef(store: AccessPageStore, reference: number): AccessUsageMap {
	const { page, row } = usageMapLocation(reference);
	return readUsageMap(store, page, row);
}
