// A MultiPage's `x` stream: the page bookkeeping.
//
// Layout ([MS-OFORMS] 2.1.2.3): an array of PageProperties - one more entry
// than there are pages, the first ignored - immediately followed by one
// MultiPageProperties whose DataBlock carries PageCount and ID, and whose
// tail is the PageIDs array naming each Page's site ID in page order.
//
// Measured against the fixture MultiPage: three 8-byte PageProperties
// (mask 0), a 16-byte MultiPageProperties (PageCount=2, ID), and 2 * 4 bytes
// of PageIDs. Everything here is regenerated rather than patched, because a
// page add or remove moves every structure at once - and the serialization
// is byte-identical on the unmodified fixture, which the suite pins.

import { OformsReader, OformsWriter } from './bytes';

export interface PageBookkeeping {
	/** Raw PageProperties records, first entry ignored per the spec. */
	pageProps: Buffer[];
	multiPagePropsMask: number;
	pageCount: number;
	/** The MultiPageProperties ID value, preserved verbatim. */
	id: number;
	/** Page site IDs, in page order. */
	pageIds: number[];
}

export function parsePageBookkeeping(x: Buffer): PageBookkeeping {
	const r = new OformsReader(x);
	const records: Array<{ raw: Buffer; mask: number; body: Buffer }> = [];
	// Records are self-sized; PageProperties entries come first and the
	// MultiPageProperties is the record whose tail (PageIDs) runs to the end.
	while (r.pos < x.length) {
		const start = r.pos;
		const minor = r.u8();
		const major = r.u8();
		if (minor !== 0x00 || major !== 0x02) {
			throw new RangeError(`x stream: not a record at ${start}`);
		}
		const cb = r.u16();
		const mask = r.u32();
		const body = r.bytes(cb - 4);
		const raw = Buffer.from(x.subarray(start, r.pos));
		records.push({ raw, mask, body });
		// The MultiPageProperties is recognisable by its PageCount bit; its
		// PageIDs tail is everything that remains.
		if (mask & (1 << 1)) {
			const bodyReader = new OformsReader(body);
			const pageCount = bodyReader.i32();
			const id = (mask & (1 << 2)) ? bodyReader.i32() : 0;
			const pageIds: number[] = [];
			while (r.pos + 4 <= x.length) { pageIds.push(r.i32()); }
			if (r.pos !== x.length) {
				throw new RangeError('x stream: trailing bytes after PageIDs');
			}
			if (pageIds.length !== pageCount) {
				throw new RangeError(`x stream: ${pageIds.length} PageIDs for PageCount ${pageCount}`);
			}
			return {
				pageProps: records.slice(0, -1).map((rec) => rec.raw),
				multiPagePropsMask: mask,
				pageCount,
				id,
				pageIds,
			};
		}
	}
	throw new RangeError('x stream has no MultiPageProperties');
}

export function serializePageBookkeeping(book: PageBookkeeping): Buffer {
	const w = new OformsWriter();
	for (const raw of book.pageProps) { w.bytes(raw); }
	w.u8(0x00);
	w.u8(0x02);
	w.u16(4 + 4 + ((book.multiPagePropsMask & (1 << 2)) ? 4 : 0));
	w.u32(book.multiPagePropsMask >>> 0);
	w.i32(book.pageCount);
	if (book.multiPagePropsMask & (1 << 2)) { w.i32(book.id); }
	for (const id of book.pageIds) { w.i32(id); }
	return w.toBuffer();
}

/** An empty PageProperties record: header, zero mask, no data. */
export function emptyPageProperties(): Buffer {
	const w = new OformsWriter();
	w.u8(0x00);
	w.u8(0x02);
	w.u16(0x0004);
	w.u32(0);
	return w.toBuffer();
}
