// Read-only VBA extraction from Access databases (.accdb / .mdb).
//
// Access does not store its VBA project as a CFB. The database is a Jet/ACE
// page file, and the project's MS-OVBA streams (the dir catalog and one
// module stream per module) live inside Long Value (LVAL) pages, each stream
// held in one LVAL row or chained across several. The layout here is
// transcribed from pyOpenVBA's corpus-verified reader (access_read.py) and
// re-verified against XLIDE's own Access-authored fixtures.
//
// The output is a synthetic in-memory CFB with the discovered streams at the
// root - VbaProject.parse falls back to root-level 'dir' and module-stream
// lookups, so the rest of the engine works on an Access database unchanged.
// The extraction is content-based throughout: rows are recognized by
// decompressing and checking what they contain, never by slot coordinates.

import { Cfb } from './cfb';
import { decompress } from './ovba';
import { VbaProject } from './vbaProject';

export class AccessDatabaseError extends Error {}

const PAGE_SIZE = 4096;
const SIGNATURES = ['Standard ACE DB\0', 'Standard Jet DB\0'];
/** PROJECTSYSKIND record header: how a decompressed dir stream begins. */
const DIR_STREAM_MAGIC = Buffer.from([0x01, 0x00, 0x04, 0x00, 0x00, 0x00]);
const LVAL_TAG = Buffer.from('LVAL', 'latin1');
const MAX_CHAIN_CHUNKS = 4096;

export function isAccessDatabase(data: Buffer): boolean {
	if (data.length < PAGE_SIZE || data.length % PAGE_SIZE !== 0) {
		return false;
	}
	const signature = data.subarray(4, 20).toString('latin1');
	return SIGNATURES.includes(signature);
}

interface LvalRow {
	page: number;
	slot: number;
	bytes: Buffer;
}

/**
 * Builds the synthetic CFB carrying the database's VBA project: the 'dir'
 * catalog stream plus each dir-declared module's carrier stream, all still in
 * their on-disk MS-OVBA form, placed at the CFB root. Throws when the file
 * holds no locatable VBA project.
 */
export function accessVbaCfb(data: Buffer): Cfb {
	if (!isAccessDatabase(data)) {
		throw new AccessDatabaseError('Not an Access database (no ACE/Jet signature).');
	}
	const rows = collectLvalRows(data);
	const catalog = findCatalogBlob(data, rows);
	if (!catalog) {
		throw new AccessDatabaseError(
			'No VBA project catalog found in the database; it may contain no VBA.',
		);
	}

	// Pass 1: a dir-only CFB tells us the declared modules (name, stream
	// name, text offset) through the one dir parser the project has.
	const dirOnly = Cfb.createEmpty();
	dirOnly.addStream('dir', catalog);
	const declared = VbaProject.parse(dirOnly).modules;

	// Pass 2: find each module's carrier blob - the LVAL row (or assembled
	// chain) into which the dir's MODULEOFFSET indexes so that the bytes at
	// that offset decompress to the module's own attribute header. Access
	// keeps shadow copies of edited modules; the last match in page order is
	// the current one.
	const cfb = Cfb.createEmpty();
	cfb.addStream('dir', catalog);
	for (const module of declared) {
		const streamName = module.streamName || module.name;
		const probe = Buffer.from(`Attribute VB_Name = "${module.name}"`, 'latin1');
		let carrier: Buffer | undefined;
		for (const blob of candidateBlobs(data, rows)) {
			if (module.textOffset >= blob.length) {
				continue;
			}
			let head: Buffer;
			try {
				head = decompress(blob.subarray(module.textOffset), streamName, probe.length);
			} catch {
				continue;
			}
			if (head.subarray(0, probe.length).equals(probe)) {
				carrier = blob;
			}
		}
		if (carrier) {
			cfb.addStream(streamName, carrier);
		}
	}
	return cfb;
}

/** Every non-tombstone row on every LVAL page, in page-then-slot order. */
function collectLvalRows(data: Buffer): LvalRow[] {
	const rows: LvalRow[] = [];
	const pageCount = data.length / PAGE_SIZE;
	for (let page = 1; page < pageCount; page++) {
		const base = page * PAGE_SIZE;
		if (data[base] !== 0x01 || !data.subarray(base + 4, base + 8).equals(LVAL_TAG)) {
			continue;
		}
		const slots = lvalSlotOffsets(data, page);
		for (let slot = 0; slot < slots.length; slot++) {
			if ((slots[slot] & 0xf000) === 0xd000) {
				continue; // tombstone
			}
			rows.push({ page, slot, bytes: lvalRowBytes(data, page, slots, slot) });
		}
	}
	return rows;
}

/** The raw slot table of an LVAL page (tombstone flags included). */
function lvalSlotOffsets(data: Buffer, page: number): number[] {
	const base = page * PAGE_SIZE;
	const count = data.readUInt16LE(base + 12);
	const out: number[] = [];
	for (let i = 0; i < count && base + 16 + 2 * i <= data.length; i++) {
		out.push(data.readUInt16LE(base + 14 + 2 * i));
	}
	return out;
}

/**
 * One LVAL row's bytes. Rows grow downward from the page end; a row runs
 * from its slot offset to the next-higher live offset (or the page end).
 */
function lvalRowBytes(data: Buffer, page: number, slots: number[], slot: number): Buffer {
	const start = slots[slot] & 0x0fff;
	let end = PAGE_SIZE;
	for (const other of slots) {
		if ((other & 0xf000) === 0xd000) {
			continue;
		}
		const offset = other & 0x0fff;
		if (offset > start && offset < end) {
			end = offset;
		}
	}
	const base = page * PAGE_SIZE;
	return data.subarray(base + start, base + end);
}

/**
 * Treats (page, slot) as the head of a chained long value: each chunk opens
 * with a `u8 next_slot, u24 next_page` prefix, (0, 0) marking the last, and
 * the payload is the concatenation with prefixes stripped.
 */
function walkLvalChain(data: Buffer, headPage: number, headSlot: number): Buffer {
	const parts: Buffer[] = [];
	const seen = new Set<number>();
	let page = headPage;
	let slot = headSlot;
	for (let i = 0; i < MAX_CHAIN_CHUNKS; i++) {
		const key = page * 65536 + slot;
		if (seen.has(key)) {
			throw new AccessDatabaseError(`LVAL chain cycle at (${page}, ${slot})`);
		}
		seen.add(key);
		const slots = lvalSlotOffsets(data, page);
		if (slot >= slots.length || (slots[slot] & 0xf000) === 0xd000) {
			throw new AccessDatabaseError(`LVAL chain references missing slot (${page}, ${slot})`);
		}
		const row = lvalRowBytes(data, page, slots, slot);
		if (row.length < 4) {
			throw new AccessDatabaseError(`LVAL chunk (${page}, ${slot}) too short for a chain prefix`);
		}
		const nextSlot = row[0];
		const nextPage = row.readUIntLE(1, 3);
		parts.push(row.subarray(4));
		if (nextPage === 0 && nextSlot === 0) {
			return Buffer.concat(parts);
		}
		if (nextPage * PAGE_SIZE >= data.length) {
			throw new AccessDatabaseError(`LVAL chain references out-of-range page ${nextPage}`);
		}
		page = nextPage;
		slot = nextSlot;
	}
	throw new AccessDatabaseError(`LVAL chain exceeded ${MAX_CHAIN_CHUNKS} chunks`);
}

/** True when a row's first 4 bytes plausibly chain to another LVAL page. */
function looksLikeChainHead(data: Buffer, row: Buffer): boolean {
	if (row.length < 4) {
		return false;
	}
	const nextPage = row.readUIntLE(1, 3);
	if (nextPage === 0 || nextPage * PAGE_SIZE >= data.length) {
		return false;
	}
	const base = nextPage * PAGE_SIZE;
	return data[base] === 0x01 && data.subarray(base + 4, base + 8).equals(LVAL_TAG);
}

/**
 * Candidate MS-OVBA stream blobs: each row interpreted standalone and, when
 * its head reads as a chain prefix, the assembled chain. Yielded in page
 * order so a later (more recent) shadow copy naturally wins a scan.
 */
function* candidateBlobs(data: Buffer, rows: LvalRow[]): Generator<Buffer> {
	for (const row of rows) {
		yield row.bytes;
		if (looksLikeChainHead(data, row.bytes)) {
			try {
				yield walkLvalChain(data, row.page, row.slot);
			} catch {
				// Not actually a chain: the standalone interpretation stands.
			}
		}
	}
}

/**
 * The compressed dir stream: the blob whose OVBA-signature offset
 * decompresses to bytes opening with the PROJECTSYSKIND record. Later
 * matches supersede earlier ones (shadow-copy rule).
 */
function findCatalogBlob(data: Buffer, rows: LvalRow[]): Buffer | undefined {
	let found: Buffer | undefined;
	for (const blob of candidateBlobs(data, rows)) {
		for (const offset of ovbaSignatureOffsets(blob)) {
			let head: Buffer;
			try {
				head = decompress(blob.subarray(offset), 'accdb-dir-probe', DIR_STREAM_MAGIC.length);
			} catch {
				continue;
			}
			if (head.subarray(0, DIR_STREAM_MAGIC.length).equals(DIR_STREAM_MAGIC)) {
				found = blob.subarray(offset);
				break;
			}
		}
	}
	return found;
}

/**
 * Offsets at which a blob plausibly begins an MS-OVBA compressed container:
 * the 0x01 signature byte followed by a chunk header whose signature bits
 * are 0b011.
 */
function ovbaSignatureOffsets(blob: Buffer): number[] {
	const out: number[] = [];
	for (let i = 0; i + 3 <= blob.length; i++) {
		if (blob[i] !== 0x01) {
			continue;
		}
		const header = blob.readUInt16LE(i + 1);
		if (((header >> 12) & 0x7) === 0b011) {
			out.push(i);
		}
	}
	return out;
}
